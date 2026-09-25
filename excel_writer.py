"""Excel（.xlsx）の出力（仕様 6章）。"""

from __future__ import annotations

import math
import os
import re
from typing import Callable

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.axis import NumericAxis
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from logger_reader import LoggerData
from merger import AnalysisTable
from models import ELAPSED_KEY, TIME_UNITS, GraphSetting, ItemSetting
from uart_reader import TIMESTAMP_COLUMN, UartData, parse_timestamp

SHEET_ANALYSIS = "解析"
SHEET_GRAPH = "グラフ"

HEADER_ROW = 6
DATA_START_ROW = 7

GAP_FILL = PatternFill("solid", fgColor="D9D9D9")
HEADER_FILL = PatternFill("solid", fgColor="E2EFDA")  # 見出し行（薄い緑）
TITLE_FONT = Font(bold=True, size=12)
THIN = Side(style="thin", color="000000")
MEDIUM = Side(style="medium", color="000000")  # 表（UART・ロガー）の外周
# シート見出し（タブ）の色
TAB_COLOR_ANALYSIS = "A9D08E"  # 緑
TAB_COLOR_GRAPH = "F4B084"  # オレンジ

# グラフの大きさ（cm）と縦の間隔（行）
CHART_WIDTH_CM = 24
CHART_HEIGHT_CM = 12
CHART_ROW_STEP = 26

_INVALID_SHEET_CHARS = re.compile(r"[\\/?*\[\]:]")
_INVALID_FILE_CHARS = set('\\/:*?"<>|')

ProgressFn = Callable[[str], None]


def sanitize_sheet_name(name: str, used: set[str]) -> str:
    """Excel のシート名制限（31文字、\\ / ? * [ ] : 不可、重複不可）に合わせる。"""
    s = _INVALID_SHEET_CHARS.sub("_", name).strip("'") or "Sheet"
    s = s[:31]
    base = s
    n = 2
    while s.lower() in {u.lower() for u in used}:
        suffix = f"({n})"
        s = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(s)
    return s


def invalid_filename_chars(name: str) -> list[str]:
    return sorted({c for c in name if c in _INVALID_FILE_CHARS})


def normalize_filename(name: str) -> str:
    name = name.strip()
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def _raw_value(text: str):
    """CSV のセルを、Excel で開いたときと同じように数値化できれば数値にする。"""
    s = text.strip()
    if not s:
        return None
    t = s.replace(" ", "")
    if t[0] in "+-.0123456789" and "_" not in t:
        if t.isdigit():
            return int(t)
        try:
            v = float(t)
        except ValueError:
            return text
        if math.isfinite(v):
            return v
    return text


def _write_uart_raw(ws, uart: UartData) -> None:
    ts_idx = None
    for r, row in enumerate(uart.raw_rows):
        if r == 0:
            header = [h.strip() for h in row]
            ts_idx = header.index(TIMESTAMP_COLUMN) if TIMESTAMP_COLUMN in header else None
            ws.append(row)
            continue
        out = []
        for c, cell in enumerate(row):
            if c == ts_idx:
                ts = parse_timestamp(cell)
                if ts is not None:
                    wc = WriteOnlyCell(ws, value=ts)
                    wc.number_format = "yyyy/mm/dd hh:mm:ss.000"
                    out.append(wc)
                    continue
            out.append(_raw_value(cell))
        ws.append(out)


def _write_logger_raw(ws, logger: LoggerData, progress: ProgressFn | None) -> None:
    total = len(logger.raw_rows)
    for n, row in enumerate(logger.raw_rows):
        ws.append([_raw_value(c) for c in row])
        if progress and n % 50000 == 0 and n:
            progress(f"ロガーの元データを書き込み中… {n:,}/{total:,} 行")


def time_unit_label(label: str, unit: str) -> str:
    """換算列の見出し：末尾の「(単位)」を置き換える（なければ付け足す）。例 time(ms) → time(min)"""
    if re.search(r"\([^()]*\)\s*$", label):
        return re.sub(r"\([^()]*\)\s*$", f"({unit})", label)
    return f"{label}({unit})"


def needed_time_columns(graphs: list[GraphSetting]) -> set[tuple[str, str]]:
    """グラフの横軸に使う時間の換算列 (元のキー, 単位) の一覧。"""
    return {(g.x, g.effective_x_unit()) for g in graphs if g.effective_x_unit()}


def _write_analysis(ws, table: AnalysisTable, uart: UartData | None, logger: LoggerData | None,
                    progress: ProgressFn | None,
                    time_cols: set[tuple[str, str]] = frozenset()) -> tuple[dict, dict]:
    """解析シートを書き、グラフ作成用に (項目キー -> 列番号, 換算列のキー -> 見出し) を返す。

    time_cols の換算列（例 ("elapsed_ms", "min")）は元の列のすぐ右に入れる。キーは "elapsed_ms@min"。
    """
    # UART ブロックの列：(キー, 見出し, 値を取り出す関数(行番号, 行))
    def unit_cols(key: str, label: str, to_ms):
        cols = []
        for u in TIME_UNITS:
            if (key, u) in time_cols:
                f = TIME_UNITS[u]
                cols.append((f"{key}@{u}", time_unit_label(label, u),
                             lambda n, row, g=to_ms, f=f: None if g(n, row) is None else round(g(n, row) / f, 9)))
        return cols

    specs = [(ELAPSED_KEY, table.elapsed_label, lambda n, row: row[0])]
    specs += unit_cols(ELAPSED_KEY, table.elapsed_label, lambda n, row: row[0] * 1000)
    raw = table.elapsed_ms_raw
    for i, it in enumerate(table.uart_items):
        specs.append((it.key, it.label, lambda n, row, i=i: row[1][i]))
        if it.key == "elapsed_ms":
            specs += unit_cols(it.key, it.label, lambda n, row: raw[n] if n < len(raw) else None)
    logger_specs = [(it.key, it.label, lambda n, row, i=i: row[2][i]) for i, it in enumerate(table.logger_items)]
    extra_labels = {k: lab for k, lab, _ in specs if "@" in k[1:]}

    # 表のかたまり（ブロック）：(見出し, 列)。ブロックの間は 1 列空ける
    if uart is None:
        # ロガーだけ：経過時間の列もロガーの表に入れる
        block_defs = [("測定値", specs + logger_specs)]
    else:
        block_defs = [("マイコン内部データ(UART)", specs)]
        if logger_specs:
            block_defs.append(("測定値", logger_specs))
    layout = []  # (見出し, 開始列, 列の一覧)
    col = 1
    for title, cols in block_defs:
        layout.append((title, col, cols))
        col += len(cols) + 1
    total_cols = col - 2

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 20
    for c in range(3, total_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 12

    start = uart.timestamps[0] if uart is not None else (logger.start_time if logger is not None else None)
    c_start = WriteOnlyCell(ws, value=start.replace(microsecond=0) if start else None)
    c_start.number_format = "yyyy/mm/dd hh:mm:ss"
    ws.append(["測定日時", c_start])
    ws.append(["測定機", logger.model if logger is not None else "（ロガーなし）"])
    ws.append([])
    ws.append([])

    blocks = [range(c0, c0 + len(cols)) for _, c0, cols in layout]

    block_of = {c: b for b in blocks for c in b}
    _borders: dict = {}

    def grid_border(col: int, top: bool = False, bottom: bool = False) -> Border:
        """表の罫線：セルごとに細線の格子、ブロック（UART・ロガー）の外周は中太線。"""
        key = (col, top, bottom)
        if key not in _borders:
            b = block_of[col]
            _borders[key] = Border(
                left=MEDIUM if col == b.start else THIN,
                right=MEDIUM if col == b.stop - 1 else THIN,
                top=MEDIUM if top else THIN,
                bottom=MEDIUM if bottom else THIN,
            )
        return _borders[key]

    def styled(value, col: int, *, font=None, fill=None, top=False, bottom=False, grid=True):
        c = WriteOnlyCell(ws, value=value)
        if font is not None:
            c.font = font
        if fill is not None:
            c.fill = fill
        if grid and col in block_of:
            c.border = grid_border(col, top, bottom)
        return c

    row5 = [None] * total_cols
    for title, c0, _ in layout:
        row5[c0 - 1] = styled(title, c0, font=TITLE_FONT, grid=False)
    ws.append(row5)

    row6: list = [None] * total_cols
    col_of: dict[str, int] = {}
    getters: list = [None] * total_cols
    for _, c0, cols in layout:
        for k, (key, lab, g) in enumerate(cols):
            row6[c0 - 1 + k] = styled(lab, c0 + k, fill=HEADER_FILL, top=True)
            col_of.setdefault(key, c0 + k)
            getters[c0 - 1 + k] = g
    ws.append(row6)

    total = len(table.rows)
    for n, row in enumerate(table.rows):
        is_gap = row[3]
        vals: list = [g(n, row) if g is not None else None for g in getters]
        is_last = n == total - 1
        for idx in range(total_cols):
            col = idx + 1
            if is_gap:
                vals[idx] = styled(vals[idx], col, fill=GAP_FILL, bottom=is_last)
            elif col in block_of:
                vals[idx] = styled(vals[idx], col, bottom=is_last)
        ws.append(vals)
        if progress and n % 20000 == 0 and n:
            progress(f"解析シートを書き込み中… {n:,}/{total:,} 行")
    return col_of, extra_labels


def _series(ref_y, ref_x, label: str, color: str) -> Series:
    s = Series(ref_y, ref_x, title=label)
    s.marker.symbol = "none"
    s.smooth = False
    s.graphicalProperties.line.solidFill = color
    s.graphicalProperties.line.width = 19050
    return s


def _add_charts(ws, ws_data, table: AnalysisTable, graphs: list[GraphSetting], col_of: dict[str, int],
                extra_labels: dict[str, str] | None = None) -> None:
    """グラフシートに散布図を並べる。データは解析シート(ws_data)のセル範囲を参照する。"""
    first = DATA_START_ROW
    last = DATA_START_ROW + len(table.rows) - 1
    items = {it.key: it for it in table.uart_items + table.logger_items}
    labels = {k: it.label for k, it in items.items()}
    labels[ELAPSED_KEY] = table.elapsed_label
    labels.update(extra_labels or {})
    col_of = {ELAPSED_KEY: 1, **col_of}

    def ref(col: int):
        return Reference(ws_data, min_col=col, min_row=first, max_row=last)

    for gi, g in enumerate(graphs):
        x_key = g.x or ELAPSED_KEY
        if g.effective_x_unit():
            x_key = f"{x_key}@{g.effective_x_unit()}"  # 時間の単位を換算した列
        xref = ref(col_of[x_key])
        prim = [k for k in g.primary if k]
        sec = [k for k in g.secondary if k]

        c1 = ScatterChart()
        c1.scatterStyle = "lineMarker"
        c1.display_blanks = "gap"
        c1.width = CHART_WIDTH_CM
        c1.height = CHART_HEIGHT_CM
        for k in prim:
            c1.series.append(_series(ref(col_of[k]), xref, labels[k], items[k].color))
        p_title = " / ".join(labels[k] for k in prim)
        c1.x_axis.title = labels[x_key]
        c1.y_axis.title = p_title
        c1.x_axis.axPos = "b"
        c1.x_axis.delete = False
        c1.y_axis.delete = False
        title = p_title
        c2 = None
        if sec:
            s_title = " / ".join(labels[k] for k in sec)
            c2 = ScatterChart()
            c2.scatterStyle = "lineMarker"
            for k in sec:
                c2.series.append(_series(ref(col_of[k]), xref, labels[k], items[k].color))
            c2.y_axis.axId = 200
            c2.y_axis.title = s_title
            c2.y_axis.crosses = "max"
            c2.y_axis.axPos = "r"
            c2.y_axis.majorGridlines = None
            c2.y_axis.delete = False
            # 第2軸のグループには非表示の横軸を別に持たせる（Excel が作る2軸グラフと同じ構成）
            c2.x_axis = NumericAxis(axId=500, crossAx=200, delete=True, axPos="b")
            c2.x_axis.title = None
            c2.x_axis.majorGridlines = None
            c2.y_axis.crossAx = 500
            c1 += c2
            title = f"{p_title}  |  {s_title}"
        # 系列番号を通し番号にする（凡例・色の割り当てが重ならないように）
        for n, s_ in enumerate(c1.series + (c2.series if c2 else [])):
            s_.idx = s_.order = n
        c1.title = title
        c1.legend.position = "b"
        # タイトル・軸タイトル・凡例がグラフ本体や目盛の数値と重ならないよう、重ねない配置にする
        c1.title.overlay = False
        c1.legend.overlay = False
        c1.x_axis.title.overlay = False
        c1.y_axis.title.overlay = False
        # 目盛の数値は軸の外側（左端・下端・右端）に置き、軸タイトルとぶつからないようにする
        c1.x_axis.tickLblPos = "low"
        c1.x_axis.crosses = "min"  # 横軸は常にグラフの下端（縦軸に負の値があっても中央に来ない）
        c1.y_axis.crosses = "min"
        c1.y_axis.tickLblPos = "low"
        if c2 is not None:
            c2.y_axis.title.overlay = False
            c2.y_axis.tickLblPos = "high"
        anchor_row = 1 + gi * CHART_ROW_STEP
        ws.add_chart(c1, f"A{anchor_row}")


def validate_graphs(graphs: list[GraphSetting], items: list[ItemSetting]) -> list[str]:
    """グラフ設定の検査。エラーメッセージの一覧を返す（空なら OK）。"""
    enabled = {it.key for it in items if it.enabled}
    errors = []
    for i, g in enumerate(graphs, start=1):
        if not g.x or (g.x != ELAPSED_KEY and g.x not in enabled):
            errors.append(f"グラフ{i}：横軸を選択してください。")
        if not g.primary[0] or g.primary[0] not in enabled:
            errors.append(f"グラフ{i}：第1軸を選択してください。")
        for axis, keys in (("第1軸", g.primary[1:]), ("第2軸", g.secondary)):
            for k in keys:
                if k == "":
                    errors.append(f"グラフ{i}：{axis}の項目を選択してください（「なし」も選べます）。")
                elif k is not None and k not in enabled:
                    errors.append(f"グラフ{i}：{axis}の項目が採用されていません。")
        chosen = [k for k in [g.x] + g.series_keys() if k]
        if len(chosen) != len(set(chosen)):
            errors.append(f"グラフ{i}：横軸・縦軸で同じ項目を2回以上選んでいます。")
    return errors


def write_workbook(
    path: str,
    table: AnalysisTable,
    uart: UartData | None,
    logger: LoggerData | None,
    graphs: list[GraphSetting],
    progress: ProgressFn | None = None,
) -> None:
    """Excel ファイルを出力する。保存に失敗した場合は OSError（PermissionError）を送出する。"""
    if not table.rows:
        raise ValueError("出力するデータ行がありません。")
    folder = os.path.dirname(os.path.abspath(path))
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"出力先フォルダがありません: {folder}")
    wb = Workbook(write_only=True)
    used: set[str] = set()

    if progress:
        progress("解析シートを書き込み中…")
    ws_a = wb.create_sheet(sanitize_sheet_name(SHEET_ANALYSIS, used))
    ws_a.sheet_properties.tabColor = TAB_COLOR_ANALYSIS
    col_of, extra_labels = _write_analysis(ws_a, table, uart, logger, progress, needed_time_columns(graphs))

    if graphs:
        ws_g = wb.create_sheet(sanitize_sheet_name(SHEET_GRAPH, used))
        ws_g.sheet_properties.tabColor = TAB_COLOR_GRAPH
        _add_charts(ws_g, ws_a, table, graphs, col_of, extra_labels)

    if uart is not None:  # 読み込んだファイルの元データだけシートにする
        if progress:
            progress("UART の元データを書き込み中…")
        ws_u = wb.create_sheet(sanitize_sheet_name(uart.name, used))
        _write_uart_raw(ws_u, uart)

    if logger is not None:
        if progress:
            progress("ロガーの元データを書き込み中…")
        ws_l = wb.create_sheet(sanitize_sheet_name(logger.name, used))
        _write_logger_raw(ws_l, logger, progress)

    if progress:
        progress("ファイルを保存中…")
    wb.save(path)
