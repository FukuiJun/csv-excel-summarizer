"""欠落補完・時刻合わせ・換算（仕様 4章）。GUI に依存しない。"""

from __future__ import annotations

import bisect
import math
import statistics
from dataclasses import dataclass, field
from datetime import datetime

from logger_reader import LoggerData
from models import SOURCE_LOGGER, SOURCE_UART, ItemSetting
from uart_reader import UartData


ELAPSED_MS_COLUMN = "elapsed_ms"


def round_half_up(x: float) -> int:
    """四捨五入（Python の round は偶数丸めのため使わない）。"""
    return int(math.floor(x + 0.5))


def elapsed_ms(ts: datetime, ts0: datetime) -> float:
    return (ts - ts0).total_seconds() * 1000.0


def default_uart_interval(timestamps: list[datetime]) -> int:
    """pc_timestamp の隣接行差の中央値を 100ms 単位に丸めた値（仕様 5.2.2）。"""
    diffs = [elapsed_ms(b, a) for a, b in zip(timestamps, timestamps[1:])]
    diffs = [d for d in diffs if d > 0]
    if not diffs:
        return 1000
    med = statistics.median(diffs)
    return max(100, round_half_up(med / 100.0) * 100)


# ---------------------------------------------------------------------------
# 4.2 UART受信の欠落の検出と補完
# ---------------------------------------------------------------------------


@dataclass
class Gap:
    prev_line: int  # 欠落直前の行（CSV上の行番号）
    prev_time: datetime
    next_line: int  # 欠落直後の行
    next_time: datetime
    missing: int  # 補完した行数


@dataclass
class UartTimeline:
    t_ms: list[float] = field(default_factory=list)  # 各行の経過時間 tU
    src: list[int | None] = field(default_factory=list)  # UartData の有効行インデックス（欠落行は None）
    gaps: list[Gap] = field(default_factory=list)
    duplicates: int = 0

    @property
    def gap_rows(self) -> int:
        return sum(g.missing for g in self.gaps)

    def __len__(self) -> int:
        return len(self.t_ms)


def build_uart_timeline(uart: UartData, interval_ms: float) -> UartTimeline:
    tl = UartTimeline()
    ts0 = uart.timestamps[0]
    last_t: float | None = None
    last_i: int | None = None
    for i, ts in enumerate(uart.timestamps):
        t = elapsed_ms(ts, ts0)
        if last_t is not None:
            diff = t - last_t
            if diff < interval_ms * 0.5:
                tl.duplicates += 1  # 二重受信など：後の行を捨てる
                continue
            if diff > interval_ms * 1.5:
                n = round_half_up(diff / interval_ms) - 1
                if n > 0:
                    for k in range(1, n + 1):
                        tl.t_ms.append(last_t + interval_ms * k)
                        tl.src.append(None)
                    tl.gaps.append(
                        Gap(
                            prev_line=uart.line_numbers[last_i],
                            prev_time=uart.timestamps[last_i],
                            next_line=uart.line_numbers[i],
                            next_time=ts,
                            missing=n,
                        )
                    )
        tl.t_ms.append(t)
        tl.src.append(i)
        last_t, last_i = t, i
    return tl


# ---------------------------------------------------------------------------
# 4.3 間隔の違いの吸収
# ---------------------------------------------------------------------------


@dataclass
class MergedRow:
    t_ms: float  # 出力する経過時間（4.4）
    uart_idx: int | None  # UartData の有効行インデックス
    logger_idx: int | None  # LoggerData の行インデックス
    is_gap: bool  # 欠落で補完した行


@dataclass
class MergeResult:
    rows: list[MergedRow]
    uart_based: bool  # True: UART基準（ロガーを間引き）、False: ロガー基準
    timeline: UartTimeline
    uart_offset_ms: float = 0.0  # rows の t_ms に含まれる UART の時間オフセット


def is_uart_based(uart_interval_ms: float, logger_interval_ms: float) -> bool:
    return uart_interval_ms >= logger_interval_ms


def merge(
    uart: UartData,
    logger: LoggerData,
    uart_interval_ms: float,
    logger_interval_ms: float,
    timeline: UartTimeline | None = None,
    uart_offset_ms: float = 0.0,
    logger_offset_ms: float = 0.0,
) -> MergeResult:
    """UART とロガーを時刻で対応付ける。

    uart_offset_ms / logger_offset_ms：それぞれのデータ全体の時刻をずらす量（ms、＋で遅らせる）。
    UART の各行は tU + uart_offset、ロガーの各行は (番号-1)×間隔 + logger_offset の時刻として扱う。
    """
    if uart_interval_ms <= 0 or logger_interval_ms <= 0:
        raise ValueError("間隔は正の値にしてください。")
    if timeline is None:
        timeline = build_uart_timeline(uart, uart_interval_ms)

    # 番号 -> 行インデックス
    num_to_idx: dict[int, int] = {}
    for idx, num in enumerate(logger.numbers):
        num_to_idx.setdefault(num, idx)
    max_num = max(logger.numbers)

    rows: list[MergedRow] = []
    if is_uart_based(uart_interval_ms, logger_interval_ms):
        for t0, src in zip(timeline.t_ms, timeline.src):
            t = t0 + uart_offset_ms
            li = round_half_up((t - logger_offset_ms) / logger_interval_ms)  # 0始まりの行番号
            if li + 1 > max_num:
                break  # ロガーのデータが尽きた
            logger_idx = num_to_idx.get(li + 1) if li >= 0 else None  # ロガーがまだ始まっていない
            rows.append(MergedRow(t, src, logger_idx, src is None))
        return MergeResult(rows, True, timeline, uart_offset_ms)

    # ロガー基準：各ロガー行に最も近い UART 行（欠落行を含む）を探す
    tol = uart_interval_ms * 0.5
    ts = [t + uart_offset_ms for t in timeline.t_ms]
    end_t = ts[-1] + tol
    for idx, num in enumerate(logger.numbers):
        t_l = (num - 1) * logger_interval_ms + logger_offset_ms
        if t_l > end_t:
            break  # UARTのデータが尽きた
        pos = bisect.bisect_left(ts, t_l)
        best = None
        for cand in (pos - 1, pos):
            if 0 <= cand < len(ts):
                if best is None or abs(ts[cand] - t_l) < abs(ts[best] - t_l):
                    best = cand
        if best is not None and abs(ts[best] - t_l) <= tol:
            src = timeline.src[best]
            rows.append(MergedRow(ts[best], src, idx, src is None))
        else:
            rows.append(MergedRow(t_l, None, idx, False))
    return MergeResult(rows, False, timeline, uart_offset_ms)


def describe_mode(uart_interval_ms: float, logger_interval_ms: float) -> str:
    ratio = max(uart_interval_ms, logger_interval_ms) / min(uart_interval_ms, logger_interval_ms)
    ratio_s = f"{ratio:.0f}" if abs(ratio - round(ratio)) < 1e-9 else f"{ratio:.1f}"
    if is_uart_based(uart_interval_ms, logger_interval_ms):
        if ratio == 1:
            return "UART基準・ロガーを時刻で対応付け（間隔が同じ）"
        return f"UART基準・ロガーを時刻で間引き（約{ratio_s}行ごと）"
    return f"ロガー基準・UARTを時刻で間引き（約{ratio_s}行ごと）"


# ---------------------------------------------------------------------------
# 出力テーブルの作成（換算）
# ---------------------------------------------------------------------------


@dataclass
class AnalysisTable:
    elapsed_label: str
    uart_items: list[ItemSetting]
    logger_items: list[ItemSetting]
    rows: list[tuple[float, list[float | None], list[float | None], bool]]
    # 各行: (経過秒数, UART値, ロガー値, 欠落行か)
    elapsed_ms_raw: list[float | None] = field(default_factory=list)  # 各行の係数前の elapsed_ms（時間単位の換算列用）


def build_analysis_table(
    result: MergeResult,
    uart: UartData,
    logger: LoggerData,
    items: list[ItemSetting],
    elapsed_label: str,
) -> AnalysisTable:
    uart_items = [it for it in items if it.enabled and it.source == SOURCE_UART and it.key in uart.values]
    logger_items = [it for it in items if it.enabled and it.source == SOURCE_LOGGER and it.key in logger.values]
    # UART 行がない行（欠落補完行など）の elapsed_ms は、直前の実 UART 行から推定する
    tl = result.timeline
    real_t = [t + result.uart_offset_ms for t, s in zip(tl.t_ms, tl.src) if s is not None]
    real_src = [s for s in tl.src if s is not None]
    has_em = ELAPSED_MS_COLUMN in uart.values
    out = []
    em_raw: list[float | None] = []
    for r in result.rows:
        sec = round(r.t_ms / 1000.0, 3)
        if r.uart_idx is None:
            uv = [None] * len(uart_items)
            est = None
            if has_em and real_t:
                p = max(bisect.bisect_right(real_t, r.t_ms) - 1, 0)
                est = round(uart.values[ELAPSED_MS_COLUMN][real_src[p]] + (r.t_ms - real_t[p]), 3)
                for i, it in enumerate(uart_items):
                    if it.key == ELAPSED_MS_COLUMN:
                        uv[i] = it.convert(est)
            em_raw.append(est)
        else:
            uv = [it.convert(uart.values[it.key][r.uart_idx]) for it in uart_items]
            em_raw.append(uart.values[ELAPSED_MS_COLUMN][r.uart_idx] if has_em else None)
        if r.logger_idx is None:
            lv = [None] * len(logger_items)
        else:
            lv = [it.convert(logger.values[it.key][r.logger_idx]) for it in logger_items]
        out.append((sec, uv, lv, r.is_gap))
    return AnalysisTable(elapsed_label, uart_items, logger_items, out, em_raw)


# ---------------------------------------------------------------------------
# 7章 行数の制限
# ---------------------------------------------------------------------------

ROWS_OK = "ok"
ROWS_WARN = "warn"
ROWS_ERROR = "error"

ANALYSIS_HEADER_ROWS = 6  # 解析シートのデータ開始前の行数


def max_sheet_rows(merged_rows: int, uart: UartData, logger: LoggerData) -> int:
    return max(ANALYSIS_HEADER_ROWS + merged_rows, len(uart.raw_rows), len(logger.raw_rows))


def check_row_limit(rows: int, warn_threshold: int, limit: int) -> str:
    if rows > limit:
        return ROWS_ERROR
    if rows > warn_threshold:
        return ROWS_WARN
    return ROWS_OK
