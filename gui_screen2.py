"""画面2：出力内容の調整（仕様書 3 章）。

左：サイドバー（読み込み結果／間隔・時間オフセット／出力先）
右：タブ（出力する項目／グラフ）
下：フッター（エラー一覧・初期値に戻す・戻る・Excelに出力）
"""

from __future__ import annotations

import math
import os
import queue
import threading
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk

import config as cfgmod
import design_tokens as T
import exporter
from excel_writer import invalid_filename_chars, normalize_filename, validate_graphs
from gui_dialogs import APP_TITLE, ColorPalette, DoneDialog, GapDetailDialog, ProgressDialog
from gui_theme import mark
from gui_widgets import CheckIcon, ScrollArea, TabBar, boxed, frame, hline, label, sized
from logger_reader import LoggerData
from merger import (
    ROWS_ERROR,
    ROWS_WARN,
    MergeResult,
    build_analysis_table,
    build_uart_timeline,
    check_row_limit,
    default_uart_interval,
    describe_mode,
    max_sheet_rows,
    merge,
)
from models import ELAPSED_KEY, MAX_SERIES_PER_AXIS, SOURCE_LOGGER, SOURCE_UART, GraphSetting, ItemSetting
from uart_reader import UartData

NONE_LABEL = "なし"
TAB_ITEMS, TAB_GRAPHS = 0, 1


def _parse_float(text: str) -> float | None:
    try:
        v = float(text.strip())
    except ValueError:
        return None
    return v if math.isfinite(v) else None


def _fmt_num(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else repr(float(v))


@dataclass
class UiError:
    """1 件のエラー。入力欄・行・タブ・フッター一覧を同じ語でつなぐ。"""

    place: str  # フッターの場所チップ（例：項目 › voltage_mV）
    message: str  # チップの右に出すメッセージ（現行の文言から場所の前置きを除いたもの）
    full: str  # 現行と同じ文言（validate() の戻り値）
    widget: tk.Misc | None = None  # クリック時にフォーカスする部品
    tab: int | None = None  # 切り替えるタブ
    section: str | None = None  # "offset" なら間隔セクションを開く


# ---------------------------------------------------------------------------
# 出力する項目の 1 行
# ---------------------------------------------------------------------------


class ItemRow:
    """出力する項目の表の 1 行（行ごとに背景色を変えられるよう tk.Frame に入れる）。"""

    def __init__(self, table: "AdjustFrame", parent, item: ItemSetting, on_change):
        th = self.th = table.th
        self.table = table
        self.key = item.key
        self.source = item.source
        self.unit = item.unit
        self.enabled = tk.BooleanVar(value=item.enabled)
        self.label = tk.StringVar(value=item.label)
        self.coef = tk.StringVar(value=_fmt_num(item.coef))
        self.state_ = "normal"

        self.frame = frame(parent, th, "panel")
        table.setup_columns(self.frame)
        f = self.frame
        pad = th.px(T.SIZE["table_col_gap"])
        edge = th.px(T.SIZE["table_pad_x"])
        vpad = th.px(2)
        self.check = ttk.Checkbutton(f, variable=self.enabled, style="panel.TCheckbutton")
        self.check.grid(row=0, column=0, padx=(edge, 0), pady=vpad, sticky="w")
        name = item.key if item.source == SOURCE_UART else f"{item.key} ({item.unit})"
        self.name_lb = label(f, th, name, bg="panel")
        self.name_lb.grid(row=0, column=1, padx=(pad, 0), sticky="w")
        self.label_e = ttk.Entry(f, textvariable=self.label, width=1)
        self.label_e.grid(row=0, column=2, padx=(pad, 0), pady=vpad, sticky="ew")
        self.coef_e = ttk.Entry(f, textvariable=self.coef, width=1, justify="right")
        self.coef_e.grid(row=0, column=3, padx=(pad, 0), pady=vpad, sticky="ew")
        self.color = item.color
        sw = frame(f, th, "panel", width=th.size("swatch_w"), height=th.size("swatch_h"))
        sw.pack_propagate(False)
        sw.grid(row=0, column=4, padx=(pad, 0), sticky="w")
        self._sw_box = sw
        self.color_btn = tk.Button(sw, relief="flat", bd=0, highlightthickness=1, cursor="hand2", command=self.choose_color)
        th.paint(self.color_btn, highlightbackground="swatch_border", highlightcolor="swatch_border")
        self.color_btn.pack(fill="both", expand=True)
        self._set_color(item.color)
        self.status = label(f, th, "", font="caption", fg="text2", bg="panel")
        self.status.grid(row=0, column=5, padx=(pad, edge), sticky="w")
        self._on_change = on_change
        self._palette: ColorPalette | None = None
        for v in (self.enabled, self.label, self.coef):
            v.trace_add("write", lambda *_: on_change())

    def grid(self, **kw) -> None:
        self.frame.pack(fill="x", **kw)

    def set_state(self, state: str, reason: str = "") -> None:
        """行の状態：normal / error / warning。背景と状態列を切り替える。"""
        bg = {"normal": "panel", "error": "error_row", "warning": "warning_row"}[state]
        fg = {"normal": "text2", "error": "error", "warning": "warning"}[state]
        th = self.th
        if state != self.state_:
            for w in (self.frame, self._sw_box):
                th.paint(w, bg=bg)
            th.paint(self.name_lb, bg=bg)
            self.check.configure(style=f"{bg}.TCheckbutton")
            self.state_ = state
        th.paint(self.status, bg=bg, fg=fg)
        self.status.configure(text=("⚠ " + reason) if reason and state != "normal" else reason)
        th.paint(self.name_lb, fg="text" if self.enabled.get() else "disabled_name")

    def _set_color(self, color: str) -> None:
        self.color = color.upper()
        c = "#" + self.color
        self.color_btn.configure(bg=c, activebackground=c, disabledforeground=c)

    def choose_color(self) -> None:
        if str(self.color_btn.cget("state")) == "disabled":
            return
        if self._palette is not None and self._palette.winfo_exists():
            self._palette.close()  # 開いているときにもう一度押したら閉じる
            self._palette = None
            return
        self._palette = ColorPalette(self.color_btn, self.color, on_pick=self._picked, th=self.th)

    def _picked(self, color: str) -> None:
        self._set_color(color)
        self._on_change()

    def to_setting(self) -> ItemSetting:
        return ItemSetting(
            color=self.color,
            key=self.key,
            source=self.source,
            enabled=self.enabled.get(),
            label=self.label.get().strip(),
            coef=_parse_float(self.coef.get()) if _parse_float(self.coef.get()) is not None else 1.0,
            unit=self.unit,
        )


# ---------------------------------------------------------------------------
# グラフ 1 つ分のブロック
# ---------------------------------------------------------------------------


class GraphRow:
    """グラフ1つ分の設定ブロック：横軸・第1軸×2・第2軸×2 のプルダウン。"""

    # 各欄の種類：x=横軸（必須）、req=必須、opt=「なし」を選べる
    KINDS = ["x", "req"] + ["opt"] * (MAX_SERIES_PER_AXIS - 1) + ["opt"] * MAX_SERIES_PER_AXIS

    def __init__(self, frame_: "AdjustFrame", graph: GraphSetting):
        th = self.th = frame_.th
        self.frame = frame_
        # 欄の値：None = 未選択（x/req）または「なし」（opt）、"" = 未選択（opt）
        self.values: list[str | None] = [graph.x] + list(graph.primary) + list(graph.secondary)
        parent = frame_.graph_list
        pad_y, pad_x = T.SIZE["graph_block_pad"]
        self.outer, body = boxed(parent, th, "panel", "border", pad=(pad_y, pad_x))
        head = frame(body, th, "panel", height=th.px(28))
        head.pack(fill="x")
        self.title = label(head, th, "", font="body_bold", bg="panel")
        self.title.pack(side="left")
        box, self.del_btn = sized(
            head, th, lambda p: ttk.Button(p, text="🗑 削除", style="Small.TButton", command=lambda: frame_.remove_graph(self)),
            72, T.SIZE["btn_small_h"], bg="panel",
        )
        box.pack(side="right")
        self.reason = label(head, th, "", font="caption", fg="error", bg="panel")
        self.reason.pack(side="right", padx=(0, th.px(12)))

        grid = frame(body, th, "panel")
        grid.pack(fill="x", pady=(th.px(8), 0))
        grid.columnconfigure(0, minsize=th.size("graph_label_w"))
        grid.columnconfigure(1, weight=1, uniform="cb")
        grid.columnconfigure(2, weight=1, uniform="cb")
        m = MAX_SERIES_PER_AXIS
        rows = [("横軸", "", [0]), ("第1軸（左）", f"{m}つまで", list(range(1, 1 + m))),
                ("第2軸（右）", f"{m}つまで", list(range(1 + m, 1 + 2 * m)))]
        self.cbs: list[ttk.Combobox] = [None] * len(self.KINDS)  # type: ignore[list-item]
        gap = th.px(T.SIZE["table_col_gap"])
        for r, (name, sub, idxs) in enumerate(rows):
            lf = frame(grid, th, "panel")
            lf.grid(row=r, column=0, sticky="w", pady=(0 if r == 0 else th.px(8), 0))
            label(lf, th, name, bg="panel").pack(side="left")
            if sub:
                label(lf, th, sub, font="caption", fg="text2", bg="panel").pack(side="left", padx=(th.px(8), 0))
            for c, i in enumerate(idxs):
                cb = ttk.Combobox(grid, state="readonly", width=1)
                cb.bind("<<ComboboxSelected>>", lambda e: self._on_select())
                cb.grid(row=r, column=1 + c, sticky="ew", padx=(0 if c == 0 else gap, 0),
                        pady=(0 if r == 0 else th.px(8), 0))
                self.cbs[i] = cb
        self._keys: list[list[str | None]] = [[] for _ in self.KINDS]

    # 互換用の参照
    @property
    def x(self) -> str | None:
        return self.values[0]

    @property
    def primary(self) -> list[str | None]:
        return self.values[1 : 1 + MAX_SERIES_PER_AXIS]

    @property
    def secondary(self) -> list[str | None]:
        return self.values[1 + MAX_SERIES_PER_AXIS :]

    @property
    def x_cb(self):
        return self.cbs[0]

    def place(self, no: int) -> None:
        self.no = no
        self.title.configure(text=f"グラフ {no}")
        self.outer.pack_forget()
        self.outer.pack(fill="x", pady=(0, self.th.px(T.SIZE["graph_block_gap"])))

    def destroy(self) -> None:
        self.outer.destroy()

    def refresh(self, choices: list[tuple[str, str]], elapsed_label: str) -> None:
        """選択肢（(key, label) の一覧）を更新し、採用が外れた項目は未選択にする。"""
        for n, (kind, cb) in enumerate(zip(self.KINDS, self.cbs)):
            if kind == "x":
                opts = [(ELAPSED_KEY, elapsed_label)] + list(choices)
            elif kind == "req":
                opts = list(choices)
            else:
                opts = [(None, NONE_LABEL)] + list(choices)
            okeys = [k for k, _ in opts]
            v = self.values[n]
            if v not in okeys:
                if kind == "opt" and v not in (None, ""):
                    self.values[n] = ""  # 選んでいた項目の採用が外れた → 未選択（エラー）
                elif kind != "opt":
                    self.values[n] = None
            self._keys[n] = okeys
            cb.configure(values=[lab for _, lab in opts])
            v = self.values[n]
            cb.set(opts[okeys.index(v)][1] if (v in okeys and (v is not None or kind == "opt")) else "")

    def _on_select(self) -> None:
        for n, cb in enumerate(self.cbs):
            i = cb.current()
            if i >= 0:
                self.values[n] = self._keys[n][i]
        self.frame.validate()

    def set_errors(self, errs: list[bool], reason: str = "") -> None:
        for cb, err in zip(self.cbs, errs):
            mark(cb, not err)
        self.reason.configure(text=("⚠ " + reason) if reason else "")

    def to_setting(self) -> GraphSetting:
        return GraphSetting(list(self.primary), list(self.secondary), self.x)


# ---------------------------------------------------------------------------
# 画面2 本体
# ---------------------------------------------------------------------------


class AdjustFrame(tk.Frame):
    def __init__(self, app, uart: UartData, logger: LoggerData):
        super().__init__(app.content, bd=0, highlightthickness=0)
        th = self.th = app.theme
        th.paint(self, bg="bg")
        self.app = app
        self.uart = uart
        self.logger = logger
        self.result: MergeResult | None = None
        self.item_rows: list[ItemRow] = []
        self.graph_rows: list[GraphRow] = []
        self.errors: list[UiError] = []
        self.warnings: list[str] = []
        self._recalc_job = None
        self._busy = False
        self._label_w = th.px(210)
        self._fixed_row = None

        items, self.unit_warnings = cfgmod.build_items(app.cfg, uart, logger)
        self.unit_changed = {w.split(" ", 1)[0] for w in self.unit_warnings}
        graphs = cfgmod.build_graphs(app.cfg, items)
        self.elapsed_var = tk.StringVar(value=cfgmod.elapsed_label(app.cfg))

        self._build_footer()
        body = frame(self, th, "bg")
        body.pack(fill="both", expand=True)
        self._build_sidebar(body)
        vline = frame(body, th, "border", width=1)
        vline.pack(side="left", fill="y")
        self._build_main(body)

        self.set_items(items, graphs)
        self.recalc()

    # =====================================================================
    # サイドバー
    # =====================================================================
    def _build_sidebar(self, parent) -> None:
        th = self.th
        side = frame(parent, th, "surface", width=th.size("sidebar_w"))
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        self.side = ScrollArea(side, th, "surface", pad_x=T.SIZE["sidebar_pad_x"], pad_y=T.SIZE["sidebar_pad_y"])
        self.side.pack(fill="both", expand=True)
        s = self.side.inner
        self._build_result_section(s)
        self._section_sep(s)
        self._build_offset_section(s)
        self._section_sep(s)
        self._build_output_section(s)

    def _section_sep(self, parent) -> None:
        gap = self.th.size("sidebar_section_gap")
        hline(parent, self.th, "divider").pack(fill="x", pady=(gap, gap))

    def _heading(self, parent, text: str, bg: str = "surface") -> tk.Label:
        return label(parent, self.th, text, font="heading", fg="text2", bg=bg)

    # ---- 読み込み結果 ----
    def _build_result_section(self, s) -> None:
        th = self.th
        self._heading(s, "読み込み結果").pack(anchor="w")
        tiles = frame(s, th, "surface")
        tiles.pack(fill="x", pady=(th.px(8), 0))
        tiles.columnconfigure(0, weight=1, uniform="t")
        tiles.columnconfigure(1, weight=1, uniform="t")
        u, lg = self.uart, self.logger
        t1 = self._tile(tiles, "UART", u.row_count,
                        [f"{u.timestamps[0]:%H:%M:%S}〜{u.timestamps[-1]:%H:%M:%S}"])
        self.uart_dup_lb = label(t1, th, "", font="caption", fg="text2", bg="tile")
        self.uart_dup_lb.pack(anchor="w")
        self._uart_skip = u.skipped
        t1.master.grid(row=0, column=0, sticky="nsew", padx=(0, th.px(4)))
        t2 = self._tile(tiles, f"ロガー {lg.model or '(モデル不明)'}", lg.row_count,
                        [f"{len(lg.channels)}CH", f"測定間隔 {lg.interval_ms:g}ms" if lg.interval_ms else "測定間隔 不明"])
        t2.master.grid(row=0, column=1, sticky="nsew", padx=(th.px(4), 0))
        self.bands = frame(s, th, "surface")
        self.bands.pack(fill="x")

    def _tile(self, parent, title: str, rows: int, lines: list[str]) -> tk.Frame:
        th = self.th
        _, inner = boxed(parent, th, "tile", "border", pad=(8, 12))
        label(inner, th, title, font="caption", fg="text2", bg="tile").pack(anchor="w")
        st = frame(inner, th, "tile")
        st.pack(anchor="w")
        label(st, th, f"{rows:,}", font="stat", bg="tile").pack(side="left")
        label(st, th, " 行", font="small_body", bg="tile").pack(side="left", anchor="s", pady=(0, th.px(3)))
        for ln in lines:
            label(inner, th, ln, font="caption", fg="text2", bg="tile").pack(anchor="w")
        return inner

    def _band(self, kind: str, text: str, bold: str = "", button: tuple[str, object] | None = None):
        """状態の帯（success / warning）。"""
        th = self.th
        if kind == "success":
            outer = frame(self.bands, th, "success_bg")
            inner = frame(outer, th, "success_bg")
        else:
            outer, inner = boxed(self.bands, th, "warning_bg", "warning_border", pad=(6, 12))
        outer.pack(fill="x", pady=(th.px(8), 0))
        bg = "success_bg" if kind == "success" else "warning_bg"
        if kind == "success":
            inner.pack(fill="x", padx=th.px(12), pady=th.px(6))
            CheckIcon(inner, th, 16, "success", bg).pack(side="left", padx=(0, th.px(8)))
        else:
            label(inner, th, "⚠", font="status", fg="warning", bg=bg).pack(side="left", anchor="n", padx=(0, th.px(8)))
        txt = frame(inner, th, bg)
        txt.pack(side="left", fill="x", expand=True)
        fg = "success" if kind == "success" else "warning"
        if bold:
            label(txt, th, bold, font="status", fg=fg, bg=bg).pack(anchor="w")
        if text:
            label(txt, th, text, font="small_body" if bold else "status", fg=fg, bg=bg,
                  wraplength=th.px(220)).pack(anchor="w")
        if button:
            box, btn = sized(inner, th, lambda p: ttk.Button(p, text=button[0], style="Small.TButton", command=button[1]),
                             60, T.SIZE["btn_small_h"], bg=bg)
            box.pack(side="right", anchor="n")
            self.gap_btn = btn
        return outer

    def _update_bands(self, tl) -> None:
        for w in self.bands.winfo_children():
            w.destroy()
        self.gap_btn = None
        if tl is None:
            self.gap_label_text = ""
        elif tl.gaps:
            self.gap_label_text = f"⚠ 欠落 {len(tl.gaps):,}箇所（{tl.gap_rows:,}行を空欄で補完）"
            self._band("warning", f"（{tl.gap_rows:,}行を空欄で補完）", bold=f"欠落 {len(tl.gaps):,}箇所",
                       button=("詳細…", self.show_gap_detail))
        else:
            self.gap_label_text = "欠落なし"
            self._band("success", "欠落なし")
        for w in self.unit_warnings:
            self._band("warning", w)

    # ---- 間隔・時間オフセット ----
    def _build_offset_section(self, s) -> None:
        th = self.th
        head = frame(s, th, "surface", height=th.px(24))
        head.pack(fill="x")
        self.offset_open = bool(self.app.cfg.get("offset_section_open", True))
        self.toggle_lb = label(head, th, "", font="heading", fg="text", bg="surface", cursor="hand2", takefocus=1)
        self.toggle_lb.pack(side="left")
        self.toggle_lb.bind("<Button-1>", lambda e: self.toggle_offset())
        self.toggle_lb.bind("<Return>", lambda e: self.toggle_offset())
        self.toggle_lb.bind("<space>", lambda e: self.toggle_offset())
        self.offset_status = label(head, th, "たまに変更", font="caption", fg="text2", bg="surface")
        self.offset_status.pack(side="right")

        self.offset_body = frame(s, th, "surface")
        g = frame(self.offset_body, th, "surface")
        g.pack(fill="x", pady=(th.px(8), 0))
        og = T.OFFSET_GRID
        g.columnconfigure(0, minsize=th.px(og["label_w"]))
        g.columnconfigure(1, weight=1, uniform="o")
        g.columnconfigure(2, weight=1, uniform="o")
        g.columnconfigure(3, minsize=th.px(og["unit_w"]))
        cg, rg = th.px(og["col_gap"]), th.px(og["row_gap"])
        label(g, th, "UART", font="caption", fg="text2", bg="surface").grid(row=0, column=1, sticky="w")
        label(g, th, "ロガー", font="caption", fg="text2", bg="surface").grid(row=0, column=2, sticky="w", padx=(cg, 0))
        self.uart_int = tk.StringVar(value=str(default_uart_interval(self.uart.timestamps)))
        li = self.logger.interval_ms
        self.logger_int = tk.StringVar(value=_fmt_num(li) if li else "")
        self.uart_off = tk.StringVar(value="0")
        self.logger_off = tk.StringVar(value="0")
        entries = []
        for r, (name, a, b) in enumerate((("間隔", self.uart_int, self.logger_int),
                                          ("オフセット", self.uart_off, self.logger_off)), start=1):
            label(g, th, name, bg="surface").grid(row=r, column=0, sticky="w", pady=(rg, 0))
            e1 = ttk.Entry(g, textvariable=a, width=1, justify="right")
            e1.grid(row=r, column=1, sticky="ew", pady=(rg, 0))
            e2 = ttk.Entry(g, textvariable=b, width=1, justify="right")
            e2.grid(row=r, column=2, sticky="ew", pady=(rg, 0), padx=(cg, 0))
            label(g, th, "ms", font="caption", fg="text2", bg="surface").grid(row=r, column=3, sticky="w",
                                                                           pady=(rg, 0), padx=(th.px(6), 0))
            entries += [e1, e2]
        self.uart_int_e, self.logger_int_e, self.uart_off_e, self.logger_off_e = entries
        self._wrap_labels = [label(self.offset_body, th, "＋で遅らせる　例：ロガーが1秒早い → ロガーに 1000",
                                   font="caption", fg="text2", bg="surface")]
        self._wrap_labels[0].pack(anchor="w", pady=(th.px(6), 0))
        self.offset_summary = label(s, th, "", font="caption", fg="text2", bg="surface")
        self._offset_anchor = frame(s, th, "surface")  # 本体・要約の差し込み位置
        self._offset_anchor.pack(fill="x")

        # 結果表示の枠（開閉どちらでも表示）
        self.result_outer, rb = boxed(s, th, "info_bg", "info_border", pad=(8, 12))
        self.result_outer.pack(fill="x", pady=(th.px(8), 0))
        self._result_inner = rb
        self.mode_line1 = label(rb, th, "", font="caption", fg="text", bg="info_bg")
        self.mode_line1.pack(anchor="w")
        self.mode_label = label(rb, th, "", font="body_bold", fg="text", bg="info_bg")
        self.mode_label.pack(anchor="w", pady=(th.px(2), 0))
        self._wrap_labels.append(self.mode_line1)
        # サイドバーの幅に合わせて折り返す
        s.bind("<Configure>", lambda e: [lb.configure(wraplength=max(50, e.width - th.px(26))) for lb in self._wrap_labels],
               add="+")
        self._apply_offset_open()

        for v in (self.uart_int, self.logger_int, self.uart_off, self.logger_off):
            v.trace_add("write", lambda *_: self.schedule_recalc())

    def toggle_offset(self, open_: bool | None = None) -> None:
        self.offset_open = (not self.offset_open) if open_ is None else open_
        self.app.cfg["offset_section_open"] = self.offset_open  # 次に出力が成功したときに保存される
        self._apply_offset_open()

    def _apply_offset_open(self) -> None:
        th = self.th
        self.toggle_lb.configure(text=("▾  " if self.offset_open else "▸  ") + "間隔・時間オフセット")
        self.offset_body.pack_forget()
        self.offset_summary.pack_forget()
        if self.offset_open:
            self.offset_body.pack(fill="x", before=self._offset_anchor)
        else:
            self.offset_summary.pack(anchor="w", padx=(th.px(20), 0), pady=(th.px(6), 0), before=self._offset_anchor)
        self._update_offset_summary()

    def _update_offset_summary(self) -> None:
        self.offset_summary.configure(
            text=f"間隔 {self.uart_int.get()} / {self.logger_int.get()} ms ・ "
                 f"オフセット {self.uart_off.get()} / {self.logger_off.get()} ms"
        )

    def _set_result_box(self, state: str, line1: str, line2: str) -> None:
        """結果表示の枠：info / warning / error。"""
        th = self.th
        bg, border, fg = {"info": ("info_bg", "info_border", "text"),
                          "warning": ("warning_bg", "warning_border", "warning"),
                          "error": ("error_bg", "error_border", "error")}[state]
        th.paint(self.result_outer, bg=bg, highlightbackground=border, highlightcolor=border)
        th.paint(self._result_inner, bg=bg)
        th.paint(self.mode_line1, bg=bg, fg=fg)
        th.paint(self.mode_label, bg=bg, fg="text" if state == "info" else fg)
        self.mode_line1.configure(text=line1)
        self.mode_label.configure(text=line2, font=th.fonts["body_bold"] if state != "error" or "出力予定 " in line2
                                  else th.fonts["caption"])

    # ---- 出力先 ----
    def _build_output_section(self, s) -> None:
        th = self.th
        self._heading(s, "出力先").pack(anchor="w")
        label(s, th, "フォルダ", font="caption", fg="text2", bg="surface").pack(anchor="w", pady=(th.px(6), th.px(2)))
        row = frame(s, th, "surface")
        row.pack(fill="x")
        self.folder = tk.StringVar(value=os.path.dirname(os.path.abspath(self.uart.path)))
        self.filename = tk.StringVar(value=cfgmod.make_output_filename(self.app.cfg, self.uart.timestamps[0]))
        box, _ = sized(row, th, lambda p: ttk.Button(p, text="参照…", command=self.browse_folder), 76, 30, bg="surface")
        box.pack(side="right", padx=(th.px(8), 0))
        self.folder_e = ttk.Entry(row, textvariable=self.folder, width=1)
        self.folder_e.pack(side="left", fill="x", expand=True)
        label(s, th, "ファイル名", font="caption", fg="text2", bg="surface").pack(anchor="w", pady=(th.px(6), th.px(2)))
        self.filename_e = ttk.Entry(s, textvariable=self.filename, width=1)
        self.filename_e.pack(fill="x")
        self.folder.trace_add("write", lambda *_: self.validate())
        self.filename.trace_add("write", lambda *_: self.validate())

    def browse_folder(self) -> None:
        cur = self.folder.get()
        path = filedialog.askdirectory(parent=self, initialdir=cur if os.path.isdir(cur) else None)
        if path:
            self.folder.set(os.path.normpath(path))

    # =====================================================================
    # メイン（タブ）
    # =====================================================================
    def _build_main(self, parent) -> None:
        th = self.th
        main = frame(parent, th, "bg")
        main.pack(side="left", fill="both", expand=True)
        wrap = frame(main, th, "bg")
        wrap.pack(fill="both", expand=True, padx=th.size("main_pad_x"), pady=th.size("main_pad_y"))
        self.nb = TabBar(wrap, th)
        self.nb.pack(fill="both", expand=True)
        self.tab_hint = self.nb.hint
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._update_tab_hint())

        # ---- 出力する項目 ----
        page = frame(self.nb.panel, th, "panel")
        self.nb.add(page, text="出力する項目")
        self.items_page = page
        head = frame(page, th, "panel", height=th.px(T.SIZE["table_header_h"]))
        head.pack(fill="x", pady=(th.px(4), 0))
        self.items_head = head
        self.setup_columns(head)
        pad = th.px(T.SIZE["table_col_gap"])
        edge = th.px(T.SIZE["table_pad_x"])
        for c, (name, _) in enumerate(T.ITEM_COLUMNS):
            if name:
                label(head, th, name, font="heading", fg="text2", bg="panel").grid(
                    row=0, column=c, sticky="w", padx=(edge if c == 0 else pad, 0), pady=(th.px(6), th.px(6)))
        hline(page, th, "row_line").pack(fill="x", padx=edge)
        self.items_area = ScrollArea(page, th, "panel", pad_y=2)
        self.items_area.pack(fill="both", expand=True, pady=(0, th.px(8)))
        self.items_inner = self.items_area.inner
        page.bind("<Configure>", self._on_items_resize)

        # ---- グラフ ----
        gpage = frame(self.nb.panel, th, "panel")
        self.nb.add(gpage, text="グラフ")
        self.graph_area = ScrollArea(gpage, th, "panel", pad_x=16, pad_y=16)
        self.graph_area.pack(fill="both", expand=True)
        self.graph_list = frame(self.graph_area.inner, th, "panel")
        self.graph_list.pack(fill="x")
        box, self.add_graph_btn = sized(
            self.graph_area.inner, th, lambda p: ttk.Button(p, text="＋  グラフを追加", command=self.add_graph),
            140, 32, bg="panel",
        )
        box.pack(anchor="w")
        self._update_tab_hint()

    def setup_columns(self, f: tk.Frame) -> None:
        """項目の表の列幅（見出し行と各行で同じにする）。"""
        th = self.th
        pad = th.px(T.SIZE["table_col_gap"])
        edge = th.px(T.SIZE["table_pad_x"])
        for c, (_, w) in enumerate(T.ITEM_COLUMNS):
            if c == 2:
                f.columnconfigure(c, minsize=self._label_w + pad)
            elif w:
                f.columnconfigure(c, minsize=th.px(w) + (edge if c == 0 else pad))
            else:
                f.columnconfigure(c, weight=1)

    def _on_items_resize(self, e) -> None:
        """幅が広いときだけラベル列を伸ばす（210〜320）。"""
        th = self.th
        fixed = sum(th.px(w) for n, w in T.ITEM_COLUMNS if w and n != T.ITEM_COLUMNS[2][0])
        gaps = th.px(T.SIZE["table_col_gap"]) * 5 + th.px(T.SIZE["table_pad_x"]) * 2
        status_w = th.px(90)
        w = max(th.px(120), min(th.px(320), e.width - fixed - gaps - status_w))
        if w != self._label_w:
            self._label_w = w
            for f in [self.items_head, self._fixed_row] + [r.frame for r in self.item_rows]:
                if f is None:
                    continue
                self.setup_columns(f)

    def _update_tab_hint(self) -> None:
        try:
            cur = self.nb.index(self.nb.select())
        except tk.TclError:
            return
        self.tab_hint.configure(text="色見本をクリックで色を変更" if cur == TAB_ITEMS else "選択肢は「採用中の項目のラベル」")

    # ---- 項目の表 ----
    def set_items(self, items: list[ItemSetting], graphs: list[GraphSetting]) -> None:
        th = self.th
        inner = self.items_inner
        for w in inner.winfo_children():
            w.destroy()
        self.item_rows = []
        pad = th.px(T.SIZE["table_col_gap"])
        edge = th.px(T.SIZE["table_pad_x"])

        # 固定行「経過時間(s)」
        fr = frame(inner, th, "panel")
        self._fixed_row = fr
        self.setup_columns(fr)
        fr.pack(fill="x")
        fixed = ttk.Checkbutton(fr, style="panel.TCheckbutton")  # 常に出力（外せない）
        fixed.state(["!alternate", "selected", "disabled"])
        fixed.grid(row=0, column=0, padx=(edge, 0), pady=th.px(2), sticky="w")
        label(fr, th, "経過時間(s)", bg="panel").grid(row=0, column=1, padx=(pad, 0), sticky="w")
        self.elapsed_e = ttk.Entry(fr, textvariable=self.elapsed_var, width=1)
        self.elapsed_e.grid(row=0, column=2, padx=(pad, 0), pady=th.px(2), sticky="ew")
        label(fr, th, "―", font="caption", fg="text2", bg="panel", anchor="center").grid(row=0, column=3, padx=(pad, 0))
        label(fr, th, "―", font="caption", fg="text2", bg="panel", anchor="center").grid(row=0, column=4, padx=(pad, 0))
        label(fr, th, "常に出力", font="caption", fg="text2", bg="panel").grid(row=0, column=5, padx=(pad, edge), sticky="w")
        if not getattr(self, "_elapsed_traced", False):
            self.elapsed_var.trace_add("write", lambda *_: self.on_items_changed())
            self._elapsed_traced = True

        self.group_labels = {}
        for src, title in ((SOURCE_UART, "UART"), (SOURCE_LOGGER, f"ロガー {self.logger.model}".strip())):
            group = [it for it in items if it.source == src]
            if not group:
                continue
            gh = frame(inner, th, "panel", height=th.px(T.SIZE["table_group_h"]))
            gh.pack(fill="x", padx=edge, pady=(th.px(6), 0))
            label(gh, th, title, font="heading", fg="text", bg="panel").pack(side="left", anchor="s")
            cnt = label(gh, th, "", font="caption", fg="text2", bg="panel")
            cnt.pack(side="left", anchor="s", padx=(th.px(8), 0))
            self.group_labels[src] = cnt
            hline(inner, th, "row_line").pack(fill="x", padx=edge, pady=(th.px(2), th.px(2)))
            for it in group:
                row = ItemRow(self, inner, it, self.on_items_changed)
                row.grid()
                self.item_rows.append(row)

        for g in self.graph_rows:
            g.destroy()
        self.graph_rows = [GraphRow(self, g) for g in graphs]
        self._regrid_graphs()
        self.on_items_changed()

    # ---- グラフ ----
    def _regrid_graphs(self) -> None:
        for i, g in enumerate(self.graph_rows, start=1):
            g.place(i)

    def add_graph(self) -> None:
        g = GraphRow(self, GraphSetting([None], [None]))  # 横軸は初期値（elapsed_ms）
        self.graph_rows.append(g)
        self._regrid_graphs()
        g.refresh(self._graph_choices(), self.elapsed_var.get().strip())
        self.validate()

    def remove_graph(self, g: GraphRow) -> None:
        g.destroy()
        self.graph_rows.remove(g)
        self._regrid_graphs()
        self.validate()

    def _graph_choices(self) -> list[tuple[str, str]]:
        return [(r.key, r.label.get().strip()) for r in self.item_rows if r.enabled.get()]

    # =====================================================================
    # フッター
    # =====================================================================
    def _build_footer(self) -> None:
        th = self.th
        self.footer = frame(self, th, "surface")
        self.footer.pack(side="bottom", fill="x")
        # エラー一覧パネル（エラーがあるときだけ表示）
        self.err_panel = frame(self.footer, th, "error_row")
        hline(self.err_panel, th, "error_border").pack(fill="x")
        ep = frame(self.err_panel, th, "error_row")
        ep.pack(fill="x", padx=th.size("page_pad_x"), pady=(th.px(10), th.px(8)))
        eh = frame(ep, th, "error_row")
        eh.pack(fill="x")
        self.err_title = label(eh, th, "", font="status", fg="error", bg="error_row")
        self.err_title.pack(side="left")
        label(eh, th, "赤い見出しをクリックすると該当の欄へ移動します", font="caption", fg="text2", bg="error_row").pack(
            side="left", padx=(th.px(10), 0))
        self.err_grid = frame(ep, th, "error_row")
        self.err_grid.pack(fill="x", pady=(th.px(4), 0))
        self.err_grid.columnconfigure(0, weight=1, uniform="e")
        self.err_grid.columnconfigure(1, weight=1, uniform="e")
        # 旧来のエラー表示（テスト・互換用。画面には出さない）
        self.error_label = tk.Label(self)

        self.footer_line = hline(self.footer, th, "border")
        self.footer_line.pack(fill="x")
        bar = frame(self.footer, th, "surface", height=th.size("footer_h"))
        bar.pack(fill="x")
        bar.pack_propagate(False)
        inner = frame(bar, th, "surface")
        inner.pack(fill="both", expand=True, padx=th.size("page_pad_x"))
        box, self.reset_btn = sized(inner, th, lambda p: ttk.Button(p, text="初期値に戻す", command=self.reset_defaults),
                                    118, 32, bg="surface")
        box.pack(side="left", pady=th.px(16))
        self.status_icon = label(inner, th, "", font="status", fg="success", bg="surface")
        self.status_icon.pack(side="left", padx=(th.px(16), th.px(6)))
        self.status_lb = label(inner, th, "", font="small_body", fg="success", bg="surface")
        self.status_lb.pack(side="left")
        box, self.export_btn = sized(
            inner, th, lambda p: ttk.Button(p, text="Excelに出力", style="Accent.TButton", command=self.export),
            T.SIZE["btn_primary_w"], 32, bg="surface",
        )
        box.pack(side="right")
        box, self.back_btn = sized(inner, th, lambda p: ttk.Button(p, text="戻る", command=self.app.show_select),
                                   T.SIZE["btn_min_w"], 32, bg="surface")
        box.pack(side="right", padx=(0, th.px(8)))

    def _update_footer(self) -> None:
        th = self.th
        errs = self.errors
        for w in self.err_grid.winfo_children():
            w.destroy()
        if errs:
            self.err_title.configure(text=f"⚠  {len(errs)}件のエラーを直すと出力できます")
            shown = errs[: T.ERROR_LIST_MAX]
            rows = math.ceil(len(shown) / T.ERROR_LIST_COLUMNS) if len(errs) <= T.ERROR_LIST_MAX else T.ERROR_LIST_MAX // 2
            for i, e in enumerate(shown):
                r, c = i % rows, i // rows  # 列優先（左列を上から、次に右列）
                cell = frame(self.err_grid, th, "error_row")
                cell.grid(row=r, column=c, sticky="w", padx=(0, th.px(24)), pady=(th.px(2), 0))
                chip = tk.Label(cell, text=e.place, font=(th.family, -th.px(12), "bold"), padx=th.px(8), pady=th.px(1),
                                cursor="hand2", takefocus=1, bd=0, highlightthickness=1)
                th.paint(chip, bg="panel", fg="error", highlightbackground="error_chip_border",
                         highlightcolor="accent")
                chip.pack(side="left")
                chip.bind("<Button-1>", lambda ev, err=e: self.goto_error(err))
                chip.bind("<Return>", lambda ev, err=e: self.goto_error(err))
                label(cell, th, e.message, font="small_body", fg="text", bg="error_row").pack(side="left", padx=(th.px(8), 0))
            if len(errs) > T.ERROR_LIST_MAX:
                label(self.err_grid, th, f"…ほか {len(errs) - T.ERROR_LIST_MAX} 件", font="caption", fg="text2",
                      bg="error_row").grid(row=rows, column=0, sticky="w", pady=(th.px(2), 0))
            if not self.err_panel.winfo_manager():
                self.err_panel.pack(fill="x", before=self.footer_line)
            self.status_icon.configure(text="")
            self.status_lb.configure(text="")
        else:
            self.err_panel.pack_forget()
            if self.warnings:
                th.paint(self.status_icon, fg="warning")
                th.paint(self.status_lb, fg="warning")
                self.status_icon.configure(text="⚠")
                self.status_lb.configure(text=f"警告 {len(self.warnings)}件（出力はできます）")
            else:
                th.paint(self.status_icon, fg="success")
                th.paint(self.status_lb, fg="success")
                self.status_icon.configure(text="✓")
                self.status_lb.configure(text="エラーはありません")

    def goto_error(self, err: UiError) -> None:
        """エラー一覧の場所チップ：①タブ切替 ②スクロール ③フォーカス。"""
        if err.section == "offset" and not self.offset_open:
            self.toggle_offset(True)
        if err.tab is not None:
            self.nb.select(err.tab)
        w = err.widget
        if w is None:
            return
        self.update_idletasks()
        area = ScrollArea.under(w)
        if area is not None:
            area.see(w)
        try:
            w.focus_set()
        except tk.TclError:
            pass

    # =====================================================================
    # 再計算・検査
    # =====================================================================
    def intervals(self) -> tuple[float | None, float | None]:
        u = _parse_float(self.uart_int.get())
        lg = _parse_float(self.logger_int.get())
        return (u if u is not None and u > 0 else None, lg if lg is not None and lg > 0 else None)

    def offsets(self) -> tuple[float | None, float | None]:
        return _parse_float(self.uart_off.get()), _parse_float(self.logger_off.get())

    def schedule_recalc(self) -> None:
        self._update_offset_summary()
        if self._recalc_job is not None:
            self.after_cancel(self._recalc_job)
        self._recalc_job = self.after(250, self.recalc)

    def recalc(self) -> None:
        self._recalc_job = None
        self._update_offset_summary()
        u, lg = self.intervals()
        mark(self.uart_int_e, u)
        mark(self.logger_int_e, lg)
        uo, lo = self.offsets()
        mark(self.uart_off_e, uo is not None)
        mark(self.logger_off_e, lo is not None)
        self.row_level = "ok"
        if not u or not lg or uo is None or lo is None:
            self.result = None
            msg = "間隔を正の数値で入力してください。" if not u or not lg else "時間オフセットを数値で入力してください。"
            self._set_result_box("error", "⚠ " + msg, "出力予定の行数は計算できません")
            self.uart_dup_lb.configure(text=f"スキップ {self._uart_skip:,}")
            if not self.bands.winfo_children():
                self._update_bands(None)
            self.validate()
            return
        tl = build_uart_timeline(self.uart, u)
        self.result = merge(self.uart, self.logger, u, lg, timeline=tl, uart_offset_ms=uo, logger_offset_ms=lo)
        self.uart_dup_lb.configure(text=f"スキップ {self._uart_skip:,} ・ 重複 {tl.duplicates:,}")
        self._update_bands(tl)

        n = len(self.result.rows)
        rows = max_sheet_rows(n, self.uart, self.logger)
        level = check_row_limit(rows, self.app.cfg["row_warn_threshold"], self.app.cfg["row_limit"])
        self.row_level = level
        line1 = describe_mode(u, lg)
        line2 = f"出力予定 {n:,} 行"
        if level != "ok":
            line2 += f"（最大シート {rows:,} 行）"
        state = {"ok": "info", ROWS_WARN: "warning", ROWS_ERROR: "error"}[level]
        self._set_result_box(state, ("⚠ " + line1) if state != "info" else line1, line2)
        self.validate()

    def on_items_changed(self) -> None:
        for r in self.item_rows:
            state = "normal" if r.enabled.get() else "disabled"
            for e in (r.label_e, r.coef_e):
                e.configure(state=state)
            r.color_btn.configure(state=state)
        choices = self._graph_choices()
        for g in self.graph_rows:
            g.refresh(choices, self.elapsed_var.get().strip())
        # 件数（グループ見出し）
        for src, lb in getattr(self, "group_labels", {}).items():
            rows = [r for r in self.item_rows if r.source == src]
            on = sum(r.enabled.get() for r in rows)
            unit = "列" if src == SOURCE_UART else "CH"
            lb.configure(text=f"{len(rows)}{unit}中 {on}{unit}を出力")
        self.validate()

    def validate(self) -> list[str]:
        """入力を検査し、エラー表示（入力欄・行・タブ・セクション・フッター）を更新する。

        戻り値は現行と同じ文言のエラーメッセージの一覧。
        """
        errs: list[UiError] = []
        u, lg = self.intervals()
        if not u or not lg:
            w = self.uart_int_e if not u else self.logger_int_e
            errs.append(UiError("間隔・オフセット", "間隔が正しくありません。", "間隔が正しくありません。", w, None, "offset"))
        uo, lo = self.offsets()
        if uo is None or lo is None:
            w = self.uart_off_e if uo is None else self.logger_off_e
            errs.append(UiError("間隔・オフセット", "時間オフセットが数値ではありません。", "時間オフセットが数値ではありません。",
                                w, None, "offset"))

        # 出力する項目
        elabel = self.elapsed_var.get().strip()
        labels: dict[str, int] = {}
        if elabel:
            labels[elabel] = 1
        for r in self.item_rows:
            if r.enabled.get() and r.label.get().strip():
                labels[r.label.get().strip()] = labels.get(r.label.get().strip(), 0) + 1
        mark(self.elapsed_e, bool(elabel) and labels.get(elabel, 0) <= 1)
        if not elabel:
            errs.append(UiError("項目 › 経過時間(s)", "ラベルが空欄です。", "経過時間のラベルが空欄です。",
                                self.elapsed_e, TAB_ITEMS))
        for r in self.item_rows:
            reasons = []
            if not r.enabled.get():
                for e in (r.label_e, r.coef_e):
                    mark(e, True)
            else:
                lab = r.label.get().strip()
                bad_label = not lab or labels.get(lab, 0) > 1
                mark(r.label_e, not bad_label)
                if not lab:
                    errs.append(UiError(f"項目 › {r.key}", "ラベルが空欄です。", f"{r.key}：ラベルが空欄です。", r.label_e, TAB_ITEMS))
                    reasons.append("ラベルが空欄")
                elif labels.get(lab, 0) > 1:
                    msg = f"ラベル「{lab}」が重複しています。"
                    errs.append(UiError(f"項目 › {r.key}", msg, f"{r.key}：{msg}", r.label_e, TAB_ITEMS))
                    reasons.append("ラベル重複")
                ok = _parse_float(r.coef.get()) is not None
                mark(r.coef_e, ok)
                if not ok:
                    errs.append(UiError(f"項目 › {r.key}", "係数が数値ではありません。", f"{r.key}：係数が数値ではありません。",
                                        r.coef_e, TAB_ITEMS))
                    reasons.append("係数が数値でない")
            if reasons:
                r.set_state("error", "・".join(reasons))
            elif r.key in self.unit_changed and r.source == SOURCE_LOGGER:
                r.set_state("warning", "単位が前回と違う")
            else:
                r.set_state("normal")
        if elabel and labels.get(elabel, 0) > 1:
            errs.append(UiError("項目 › 経過時間(s)", f"ラベル「{elabel}」が他の項目と重複しています。",
                                f"経過時間のラベル「{elabel}」が他の項目と重複しています。", self.elapsed_e, TAB_ITEMS))

        # グラフ
        items = [r.to_setting() for r in self.item_rows]
        graphs = [g.to_setting() for g in self.graph_rows]
        enabled = {it.key for it in items if it.enabled}
        graph_msgs = validate_graphs(graphs, items)
        for i, g in enumerate(self.graph_rows, start=1):
            chosen = [v for v in g.values if v]
            flags = []
            for kind, v in zip(g.KINDS, g.values):
                if kind == "x":
                    bad = not v or (v != ELAPSED_KEY and v not in enabled)
                elif kind == "req":
                    bad = not v or v not in enabled
                else:
                    bad = v == "" or (v is not None and v not in enabled)
                flags.append(bad or (bool(v) and chosen.count(v) > 1))
            mine = [m.split("：", 1)[1] for m in graph_msgs if m.startswith(f"グラフ{i}：")]
            g.set_errors(flags, mine[0].rstrip("。") if mine else "")
            first_bad = next((cb for cb, f in zip(g.cbs, flags) if f), g.cbs[0])
            for m in mine:
                errs.append(UiError(f"グラフ {i}", m, f"グラフ{i}：{m}", first_bad, TAB_GRAPHS))

        # 出力先
        folder = self.folder.get().strip()
        folder_ok = bool(folder) and os.path.isdir(folder)
        mark(self.folder_e, folder_ok)
        if not folder_ok:
            errs.append(UiError("出力先", "出力先フォルダがありません。", "出力先フォルダがありません。", self.folder_e))
        fname = self.filename.get().strip()
        bad = invalid_filename_chars(fname)
        mark(self.filename_e, bool(fname) and not bad)
        if not fname:
            errs.append(UiError("出力先", "ファイル名が空欄です。", "ファイル名が空欄です。", self.filename_e))
        elif bad:
            msg = "ファイル名に使えない文字が含まれています：" + " ".join(bad)
            errs.append(UiError("出力先", msg, msg, self.filename_e))

        if self.result is not None and not self.result.rows:
            msg = "出力するデータ行がありません（間隔の設定を確認してください）。"
            errs.append(UiError("間隔・オフセット", msg, msg, self.uart_int_e, None, "offset"))

        self.errors = errs
        # 警告（出力はできる）
        warns = []
        if self.result is not None and self.result.timeline.gaps:
            warns.append("欠落")
        warns += self.unit_warnings
        if getattr(self, "row_level", "ok") == ROWS_WARN:
            warns.append("行数")
        self.warnings = warns

        # タブ名・セクション見出し
        n_items = sum(1 for e in errs if e.tab == TAB_ITEMS)
        n_graph = sum(1 for e in errs if e.tab == TAB_GRAPHS)
        total = len(self.item_rows) + 1
        on = sum(r.enabled.get() for r in self.item_rows) + 1
        self.nb.tab(TAB_ITEMS, text="出力する項目", count=f"エラー {n_items}" if n_items else f"{on} / {total}",
                    error=bool(n_items))
        self.nb.tab(TAB_GRAPHS, text="グラフ", count=f"エラー {n_graph}" if n_graph else f"{len(self.graph_rows)}",
                    error=bool(n_graph))
        n_off = sum(1 for e in errs if e.section == "offset")
        if n_off:
            self.offset_status.configure(text=f"⚠ {n_off}件")
            self.th.paint(self.offset_status, fg="error")
            if not self.offset_open:
                self.toggle_offset(True)  # エラーがある場合は自動で開く
        else:
            self.offset_status.configure(text="たまに変更")
            self.th.paint(self.offset_status, fg="text2")

        full = [e.full for e in errs]
        self.error_label.configure(text="\n".join(full))
        self._update_footer()
        if not self._busy:
            self.export_btn.state(["disabled"] if errs else ["!disabled"])
        return full

    # =====================================================================
    # 詳細・初期値に戻す
    # =====================================================================
    def show_gap_detail(self) -> None:
        if self.result is None:
            return
        return GapDetailDialog(self, self.th, self.result.timeline.gaps)

    def reset_defaults(self) -> None:
        if not messagebox.askokcancel(
            APP_TITLE, "出力する項目とグラフの設定を初期値に戻します。よろしいですか？", parent=self
        ):
            return
        items, _ = cfgmod.build_items(self.app.cfg, self.uart, self.logger, use_saved=False)
        graphs = cfgmod.build_graphs(self.app.cfg, items, use_saved=False)
        self.elapsed_var.set(cfgmod.elapsed_label(self.app.cfg, use_saved=False))
        self.set_items(items, graphs)

    # =====================================================================
    # 出力
    # =====================================================================
    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        for b in (self.export_btn, self.back_btn, self.reset_btn):
            b.state(["disabled"] if busy else ["!disabled"])
        if not busy:
            self.validate()

    def export(self) -> None:
        if self.validate() or self.result is None:
            return
        fname = normalize_filename(self.filename.get())
        path = os.path.join(self.folder.get().strip(), fname)
        if os.path.exists(path) and not messagebox.askyesno(
            APP_TITLE, f"同名のファイルがあります。上書きしますか？\n{path}", parent=self
        ):
            return

        rows = max_sheet_rows(len(self.result.rows), self.uart, self.logger)
        level = check_row_limit(rows, self.app.cfg["row_warn_threshold"], self.app.cfg["row_limit"])
        if level == ROWS_ERROR:
            messagebox.showerror(
                APP_TITLE,
                f"出力する行数が上限を超えています（最大 {rows:,}行 / 上限 {self.app.cfg['row_limit']:,}行）。\n出力しません。",
                parent=self,
            )
            return
        if level == ROWS_WARN and not messagebox.askokcancel(
            APP_TITLE,
            f"出力する行数が多いです（最大 {rows:,}行）。\n出力に時間がかかる可能性があります。\n\n"
            "続行する場合は［OK］、やめる場合は［キャンセル］を押してください。",
            parent=self,
        ):
            return

        items = [r.to_setting() for r in self.item_rows]
        graphs = [g.to_setting() for g in self.graph_rows]
        elapsed = self.elapsed_var.get().strip()
        table = build_analysis_table(self.result, self.uart, self.logger, items, elapsed)

        progress = ProgressDialog(self, self.th)
        self._set_busy(True)
        q: queue.Queue = queue.Queue()

        cfg = self.app.cfg
        config_path = self.app.config_path
        uart_dir = os.path.dirname(os.path.abspath(self.uart.path))
        logger_dir = os.path.dirname(os.path.abspath(self.logger.path))

        def worker():
            try:
                res = exporter.export(
                    path, table, self.uart, self.logger, items, graphs, elapsed, cfg, config_path,
                    uart_dir=uart_dir, logger_dir=logger_dir, progress=lambda m: q.put(("msg", m)),
                )
                q.put(("done", res))
            except Exception as e:  # noqa: BLE001
                q.put(("error", e))

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                while True:
                    kind, val = q.get_nowait()
                    if kind == "msg":
                        progress.set_message(val)
                    else:
                        progress.close()
                        self._set_busy(False)
                        if kind == "done":
                            self._on_export_success(path, val)
                        else:
                            self._on_export_error(path, val)
                        return
            except queue.Empty:
                pass
            self.after(100, poll)

        self.after(100, poll)

    def _on_export_success(self, path: str, res: "exporter.ExportResult") -> None:
        self.app.cfg = res.config  # 保存に失敗しても、このセッション中は反映する
        self.app.dirs = dict(res.config.get("last_dirs", {}))
        note = f"※設定を保存できませんでした（{res.config_error}）" if res.config_error else ""
        DoneDialog(self, path, note, self.th)

    def _on_export_error(self, path: str, e: Exception) -> None:
        if isinstance(e, PermissionError):
            msg = (
                f"ファイルを保存できませんでした。\n{path}\n\n"
                "ファイルが Excel で開かれている場合は閉じてから、もう一度出力してください。"
            )
        else:
            msg = f"出力に失敗しました。\n{type(e).__name__}: {e}"
        messagebox.showerror(APP_TITLE, msg, parent=self)
