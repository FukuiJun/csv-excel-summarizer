"""GRAPHTEC GL240 などデータロガーの CSV の読み込み。

仕様 3.2：
- Shift-JIS（CP932）
- ヘッダ部から モデル・測定間隔・アンプ設定（CHごとの単位）を取得
- `測定値` の行の次の2行がヘッダ（項目名・単位）、その次からデータ
- 内部時刻（日付/時間, ms）は使わず、経過時間は (番号-1) × 間隔 とする
"""

from __future__ import annotations

import csv
import io
import os
import re
from dataclasses import dataclass, field
from datetime import datetime

_CH_RE = re.compile(r"^CH\d+$", re.IGNORECASE)
_INTERVAL_RE = re.compile(r"^\s*([0-9]*\.?[0-9]+)\s*(ms|msec|s|sec|min|m|h|hour)?\s*$", re.IGNORECASE)
_INTERVAL_UNIT_MS = {
    None: 1000.0,  # 単位なしは秒とみなす
    "ms": 1.0,
    "msec": 1.0,
    "s": 1000.0,
    "sec": 1000.0,
    "min": 60_000.0,
    "m": 60_000.0,
    "h": 3_600_000.0,
    "hour": 3_600_000.0,
}


class LoggerFormatError(ValueError):
    """ロガー CSV が想定した形式でないときの例外。"""


@dataclass
class ChannelInfo:
    name: str  # 例: CH1
    unit: str  # 例: mV（角括弧なし）


@dataclass
class LoggerData:
    path: str
    model: str
    interval_ms: float | None
    channels: list[ChannelInfo]
    numbers: list[int]  # 各データ行の 番号
    values: dict[str, list[float | None]]  # CH名 -> 値（数値化できなければ None）
    raw_rows: list[list[str]] = field(default_factory=list)
    start_time: datetime | None = None
    end_time: datetime | None = None

    @property
    def row_count(self) -> int:
        return len(self.numbers)

    @property
    def name(self) -> str:
        return os.path.splitext(os.path.basename(self.path))[0]

    def unit_of(self, ch: str) -> str:
        for c in self.channels:
            if c.name == ch:
                return c.unit
        return ""


def parse_interval_ms(text: str) -> float | None:
    """'200ms'、'1s'、'1min' などを ms に換算する。"""
    m = _INTERVAL_RE.match(text or "")
    if not m:
        return None
    unit = m.group(2).lower() if m.group(2) else None
    return float(m.group(1)) * _INTERVAL_UNIT_MS[unit]


def parse_number(text: str) -> float | None:
    """'+  0.28' のような値を数値化する。できなければ None。"""
    s = (text or "").replace(" ", "").replace("　", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _strip_unit(text: str) -> str:
    s = (text or "").strip().strip('"').strip()
    if s.startswith("[") and s.endswith("]"):
        s = s[1:-1]
    return s.strip()


def _parse_datetime(date_s: str, time_s: str) -> datetime | None:
    text = f"{date_s.strip()} {time_s.strip()}"
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def _read_text(path: str) -> str:
    with open(path, "rb") as f:
        data = f.read()
    try:
        return data.decode("cp932")
    except UnicodeDecodeError as e:
        raise LoggerFormatError(f"ロガー CSV を Shift-JIS(CP932) として読み込めません: {e}") from e


def read_logger_csv(path: str) -> LoggerData:
    text = _read_text(path)
    raw_rows = list(csv.reader(io.StringIO(text, newline="")))
    if not raw_rows:
        raise LoggerFormatError("ロガー CSV が空です。")

    model = ""
    interval_ms: float | None = None
    amp_units: dict[str, str] = {}
    start_time = end_time = None
    data_header_idx: int | None = None

    i = 0
    n = len(raw_rows)
    while i < n:
        row = raw_rows[i]
        key = row[0].strip() if row else ""
        if key == "モデル" and len(row) > 1:
            model = row[1].strip()
        elif key == "測定間隔" and len(row) > 1:
            interval_ms = parse_interval_ms(row[1])
        elif key == "開始時刻" and len(row) > 2:
            start_time = _parse_datetime(row[1], row[2])
        elif key == "終了時刻" and len(row) > 2:
            end_time = _parse_datetime(row[1], row[2])
        elif key == "アンプ設定":
            # 次の行がセクションの見出し（CH,信号名,...,単位,...）
            if i + 1 < n:
                sec_header = [c.strip() for c in raw_rows[i + 1]]
                unit_idx = sec_header.index("単位") if "単位" in sec_header else None
                j = i + 2
                while j < n and raw_rows[j] and _CH_RE.match(raw_rows[j][0].strip()):
                    r = raw_rows[j]
                    ch = r[0].strip().upper()
                    unit = ""
                    if unit_idx is not None and unit_idx < len(r):
                        unit = _strip_unit(r[unit_idx])
                    else:  # 見出しがない場合は [..] 形式のセルを探す
                        for c in r:
                            c = c.strip()
                            if c.startswith("[") and c.endswith("]"):
                                unit = _strip_unit(c)
                                break
                    amp_units[ch] = unit
                    j += 1
                i = j
                continue
        elif key == "測定値":
            data_header_idx = i + 1
            break
        i += 1

    if data_header_idx is None or data_header_idx + 1 >= n:
        raise LoggerFormatError(
            "ロガー CSV に『測定値』の行が見つかりません。GL240 の CSV か確認してください。"
        )

    item_row = [c.strip() for c in raw_rows[data_header_idx]]
    unit_row = [c.strip() for c in raw_rows[data_header_idx + 1]]

    if "番号" in item_row:
        num_idx = item_row.index("番号")
    else:
        num_idx = 0

    channels: list[ChannelInfo] = []
    ch_idx: dict[str, int] = {}
    for idx, name in enumerate(item_row):
        if _CH_RE.match(name):
            ch = name.upper()
            unit = amp_units.get(ch)
            if not unit:
                unit = _strip_unit(unit_row[idx]) if idx < len(unit_row) else ""
            channels.append(ChannelInfo(ch, unit))
            ch_idx[ch] = idx
    if not channels:
        raise LoggerFormatError("ロガー CSV のデータ部に CH 列がありません。")

    numbers: list[int] = []
    values: dict[str, list[float | None]] = {c.name: [] for c in channels}
    for row in raw_rows[data_header_idx + 2:]:
        if not row or not any(c.strip() for c in row):
            continue
        num_v = parse_number(row[num_idx]) if num_idx < len(row) else None
        if num_v is None:
            continue
        numbers.append(int(num_v))
        for ch, idx in ch_idx.items():
            values[ch].append(parse_number(row[idx]) if idx < len(row) else None)

    if not numbers:
        raise LoggerFormatError("ロガー CSV に有効なデータ行がありません。")

    return LoggerData(
        path=path,
        model=model,
        interval_ms=interval_ms,
        channels=channels,
        numbers=numbers,
        values=values,
        raw_rows=raw_rows,
        start_time=start_time,
        end_time=end_time,
    )
