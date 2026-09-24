"""UART CSV（マイコンからUART経由で受信したCSV）の読み込み。

仕様 3.1：
- UTF-8、1行目がヘッダ
- `pc_timestamp` を時刻の基準として使う
- 列が欠けている・数値に変換できない行はスキップし、件数を数える
"""

from __future__ import annotations

import csv
import os
from dataclasses import dataclass, field
from datetime import datetime

TIMESTAMP_COLUMN = "pc_timestamp"

_TIMESTAMP_FORMATS = (
    "%Y-%m-%d %H:%M:%S.%f",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S.%f",
    "%Y/%m/%d %H:%M:%S",
)


class UartFormatError(ValueError):
    """UART CSV が想定した形式でないときの例外。"""


@dataclass
class UartData:
    path: str
    columns: list[str]  # pc_timestamp 以外の数値列（CSVの並び順）
    timestamps: list[datetime]  # 有効行の pc_timestamp
    values: dict[str, list[float]]  # 列名 -> 有効行の値
    line_numbers: list[int]  # 有効行のCSV上の行番号（ヘッダ=1）
    raw_rows: list[list[str]] = field(default_factory=list)  # ヘッダを含む全行（元データシート用）
    skipped: int = 0

    @property
    def row_count(self) -> int:
        return len(self.timestamps)

    @property
    def name(self) -> str:
        return os.path.splitext(os.path.basename(self.path))[0]


def parse_timestamp(text: str) -> datetime | None:
    text = text.strip()
    for fmt in _TIMESTAMP_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _to_float(text: str) -> float | None:
    text = text.strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def read_uart_csv(path: str) -> UartData:
    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            raw_rows = [row for row in csv.reader(f)]
    except UnicodeDecodeError as e:
        raise UartFormatError(f"UART CSV を UTF-8 として読み込めません: {e}") from e

    if not raw_rows:
        raise UartFormatError("UART CSV が空です。")

    header = [h.strip() for h in raw_rows[0]]
    if TIMESTAMP_COLUMN not in header:
        raise UartFormatError(
            f"UART CSV の1行目に '{TIMESTAMP_COLUMN}' 列がありません。"
            "UART CSV とロガー CSV を取り違えていないか確認してください。"
        )
    ts_idx = header.index(TIMESTAMP_COLUMN)
    columns = [h for i, h in enumerate(header) if i != ts_idx and h]
    col_idx = {h: header.index(h) for h in columns}

    timestamps: list[datetime] = []
    values: dict[str, list[float]] = {c: [] for c in columns}
    line_numbers: list[int] = []
    skipped = 0

    for line_no, row in enumerate(raw_rows[1:], start=2):
        if not any(cell.strip() for cell in row):
            continue  # 空行は数えない
        if len(row) < len(header):
            skipped += 1
            continue
        ts = parse_timestamp(row[ts_idx])
        if ts is None:
            skipped += 1
            continue
        parsed = {}
        ok = True
        for c in columns:
            v = _to_float(row[col_idx[c]])
            if v is None:
                ok = False
                break
            parsed[c] = v
        if not ok:
            skipped += 1
            continue
        timestamps.append(ts)
        line_numbers.append(line_no)
        for c in columns:
            values[c].append(parsed[c])

    if not timestamps:
        raise UartFormatError("UART CSV に有効なデータ行がありません。")

    return UartData(
        path=path,
        columns=columns,
        timestamps=timestamps,
        values=values,
        line_numbers=line_numbers,
        raw_rows=raw_rows,
        skipped=skipped,
    )
