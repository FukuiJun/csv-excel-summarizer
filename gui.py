"""GUI（tkinter）。画面1：ファイル選択、画面2：出力内容の調整（仕様 5章）。"""

from __future__ import annotations

import math
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import config as cfgmod
import exporter
from excel_writer import invalid_filename_chars, normalize_filename, validate_graphs
from logger_reader import LoggerData, LoggerFormatError, read_logger_csv
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
from models import SOURCE_LOGGER, SOURCE_UART, GraphSetting, ItemSetting
from uart_reader import UartData, UartFormatError, read_uart_csv

APP_TITLE = "ロガーCSV・UART CSV 統合ツール"
ERR_BG = "#ffc0c0"
OK_BG = "white"
DISABLED_BG = "#e8e8e8"
NONE_LABEL = "なし"
LABEL_WIDTH = 20  # ラベル入力欄の幅（文字数）
CSV_TYPES = [("CSV ファイル", "*.csv *.CSV"), ("すべてのファイル", "*.*")]


def _parse_float(text: str) -> float | None:
    try:
        v = float(text.strip())
    except ValueError:
        return None
    return v if math.isfinite(v) else None


def open_folder(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError as e:
        messagebox.showerror(APP_TITLE, f"フォルダを開けませんでした: {e}")


class App(tk.Tk):
    def __init__(self, config_path: str | None = None):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("920x860")
        self.minsize(480, 320)

        style = ttk.Style(self)
        style.map("Error.TCombobox", fieldbackground=[("readonly", ERR_BG), ("!disabled", ERR_BG)])

        self.config_path = config_path or cfgmod.default_config_path()
        res = cfgmod.load_config(self.config_path)
        self.cfg = res.config
        # 「参照」ダイアログの初期フォルダ（出力成功時に config.json へ保存）
        self.dirs = dict(self.cfg.get("last_dirs", {}))

        self.uart_path = tk.StringVar()
        self.logger_path = tk.StringVar()

        self.select_frame = FileSelectFrame(self)
        self.adjust_frame: AdjustFrame | None = None
        self.select_frame.pack(fill="both", expand=True)

        if res.status == "broken":
            self.after(
                100,
                lambda: messagebox.showwarning(
                    APP_TITLE, f"config.json が壊れていたため、既定値で作り直しました。\n{res.message}"
                ),
            )
        elif res.message:
            self.after(100, lambda: messagebox.showwarning(APP_TITLE, res.message))

    def show_select(self) -> None:
        if self.adjust_frame is not None:
            self.adjust_frame.destroy()
            self.adjust_frame = None
        self.select_frame.pack(fill="both", expand=True)

    def show_adjust(self, uart: UartData, logger: LoggerData) -> None:
        self.select_frame.pack_forget()
        self.adjust_frame = AdjustFrame(self, uart, logger)
        self.adjust_frame.pack(fill="both", expand=True)


# ---------------------------------------------------------------------------
# 画面1：ファイル選択
# ---------------------------------------------------------------------------


class FileSelectFrame(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app, padding=16)
        self.app = app
        box = ttk.LabelFrame(self, text="ファイル選択", padding=12)
        box.pack(fill="x", anchor="n")
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="UART CSV").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(box, textvariable=app.uart_path).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(box, text="参照", command=lambda: self.browse("uart", app.uart_path)).grid(
            row=0, column=2, padx=(8, 0)
        )
        ttk.Label(box, text="ロガー CSV").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(box, textvariable=app.logger_path).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(box, text="参照", command=lambda: self.browse("logger", app.logger_path)).grid(
            row=1, column=2, padx=(8, 0)
        )
        self.load_btn = ttk.Button(box, text="読み込み", command=self.load)
        self.load_btn.grid(row=2, column=2, sticky="e", pady=(12, 0))

        app.uart_path.trace_add("write", lambda *_: self.update_state())
        app.logger_path.trace_add("write", lambda *_: self.update_state())
        self.update_state()

    def browse(self, kind: str, var: tk.StringVar) -> None:
        initial = os.path.dirname(var.get()) if var.get() else self.app.dirs.get(kind, "")
        if not initial or not os.path.isdir(initial):
            initial = None
        path = filedialog.askopenfilename(
            parent=self, title="UART CSV を選択" if kind == "uart" else "ロガー CSV を選択",
            initialdir=initial, filetypes=CSV_TYPES,
        )
        if path:
            var.set(os.path.normpath(path))
            self.app.dirs[kind] = os.path.dirname(os.path.normpath(path))

    def update_state(self) -> None:
        ok = bool(self.app.uart_path.get().strip() and self.app.logger_path.get().strip())
        self.load_btn.state(["!disabled"] if ok else ["disabled"])

    def load(self) -> None:
        up = self.app.uart_path.get().strip().strip('"')
        lp = self.app.logger_path.get().strip().strip('"')
        self.app.config(cursor="watch")
        self.update_idletasks()
        try:
            try:
                uart = read_uart_csv(up)
            except FileNotFoundError:
                raise UartFormatError(f"UART CSV が見つかりません:\n{up}")
            except UartFormatError:
                raise
            except Exception as e:  # noqa: BLE001
                raise UartFormatError(f"UART CSV を読み込めません: {e}")
            try:
                logger = read_logger_csv(lp)
            except FileNotFoundError:
                raise LoggerFormatError(f"ロガー CSV が見つかりません:\n{lp}")
            except LoggerFormatError:
                raise
            except Exception as e:  # noqa: BLE001
                raise LoggerFormatError(f"ロガー CSV を読み込めません: {e}")
        except (UartFormatError, LoggerFormatError) as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
            return
        finally:
            self.app.config(cursor="")
        self.app.show_adjust(uart, logger)


# ---------------------------------------------------------------------------
# 画面2：出力内容の調整
# ---------------------------------------------------------------------------


class ScrollableFrame(ttk.Frame):
    """縦スクロールできる領域。マウスホイールは、この領域が表示されている間はどこの上でも効く。"""

    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self._update_region())
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        top = self.winfo_toplevel()
        top.bind_all("<MouseWheel>", self._on_wheel, add="+")
        top.bind_all("<Button-4>", self._on_wheel, add="+")
        top.bind_all("<Button-5>", self._on_wheel, add="+")
        # コンボボックス上のホイールで選択値が変わらないようにし、画面のスクロールだけ行う
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            top.bind_class("TCombobox", seq, lambda e: (self._on_wheel(e), "break")[1])
        self.bind("<Destroy>", self._on_destroy)

    def _on_canvas_configure(self, e) -> None:
        self.canvas.itemconfigure(self._win, width=e.width)
        self._update_region()

    def _update_region(self) -> None:
        h = max(self.inner.winfo_reqheight(), self.canvas.winfo_height())
        self.canvas.configure(scrollregion=(0, 0, self.inner.winfo_reqwidth(), h))

    def _can_scroll(self) -> bool:
        return self.inner.winfo_reqheight() > self.canvas.winfo_height()

    def _on_wheel(self, e) -> None:
        if not self.winfo_exists() or not self.winfo_ismapped() or not self._can_scroll():
            return
        w = e.widget
        if isinstance(w, str) or not str(w).startswith(str(self.winfo_toplevel())) or "popdown" in str(w):
            return  # ダイアログやコンボボックスのドロップダウン上では何もしない
        if w.winfo_toplevel() is not self.winfo_toplevel():
            return
        if getattr(e, "num", None) == 4:
            step = -1
        elif getattr(e, "num", None) == 5:
            step = 1
        else:
            step = int(-e.delta / 120) or (-1 if e.delta > 0 else 1)
        self.canvas.yview_scroll(step * 2, "units")

    def _on_destroy(self, e) -> None:
        if e.widget is not self:
            return
        top = self.winfo_toplevel()
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            top.unbind_all(seq)
            top.unbind_class("TCombobox", seq)


class ItemRow:
    """出力する項目の1行分のウィジェット。"""

    def __init__(self, parent, row: int, item: ItemSetting, on_change):
        self.key = item.key
        self.source = item.source
        self.unit = item.unit
        self.enabled = tk.BooleanVar(value=item.enabled)
        self.label = tk.StringVar(value=item.label)
        self.coef = tk.StringVar(value=_fmt_num(item.coef))
        self.offset = tk.StringVar(value=_fmt_num(item.offset))

        name = item.key if item.source == SOURCE_UART else f"{item.key} ({item.unit})"
        self.check = ttk.Checkbutton(parent, variable=self.enabled)
        self.check.grid(row=row, column=0, padx=4)
        ttk.Label(parent, text=name).grid(row=row, column=1, sticky="w", padx=4)
        self.label_e = tk.Entry(parent, textvariable=self.label, width=LABEL_WIDTH)
        self.label_e.grid(row=row, column=2, sticky="w", padx=4, pady=1)
        self.coef_e = tk.Entry(parent, textvariable=self.coef, width=10)
        self.coef_e.grid(row=row, column=3, padx=4)
        self.offset_e = tk.Entry(parent, textvariable=self.offset, width=10)
        self.offset_e.grid(row=row, column=4, padx=4)

        for v in (self.enabled, self.label, self.coef, self.offset):
            v.trace_add("write", lambda *_: on_change())

    def to_setting(self) -> ItemSetting:
        return ItemSetting(
            key=self.key,
            source=self.source,
            enabled=self.enabled.get(),
            label=self.label.get().strip(),
            coef=_parse_float(self.coef.get()) if _parse_float(self.coef.get()) is not None else 1.0,
            offset=_parse_float(self.offset.get()) if _parse_float(self.offset.get()) is not None else 0.0,
            unit=self.unit,
        )


def _fmt_num(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else repr(float(v))


class GraphRow:
    def __init__(self, frame: "AdjustFrame", graph: GraphSetting):
        self.frame = frame
        self.primary: str | None = graph.primary
        # None = 「なし」、"" = 未選択
        self.secondary: str | None = graph.secondary
        parent = frame.graph_inner
        self.no = ttk.Label(parent, width=3)
        self.p_cb = ttk.Combobox(parent, state="readonly", width=24)
        self.s_cb = ttk.Combobox(parent, state="readonly", width=24)
        self.del_btn = ttk.Button(parent, text="削除", command=lambda: frame.remove_graph(self))
        self.p_cb.bind("<<ComboboxSelected>>", lambda e: self._on_select())
        self.s_cb.bind("<<ComboboxSelected>>", lambda e: self._on_select())

    def grid(self, row: int) -> None:
        self.no.configure(text=str(row))
        self.no.grid(row=row, column=0, padx=4, pady=2)
        self.p_cb.grid(row=row, column=1, padx=4, pady=2)
        self.s_cb.grid(row=row, column=2, padx=4, pady=2)
        self.del_btn.grid(row=row, column=3, padx=4, pady=2)

    def destroy(self) -> None:
        for w in (self.no, self.p_cb, self.s_cb, self.del_btn):
            w.destroy()

    def refresh(self, choices: list[tuple[str, str]]) -> None:
        """選択肢（(key, label) の一覧）を更新し、採用が外れた項目は未選択にする。"""
        keys = [k for k, _ in choices]
        labels = [lab for _, lab in choices]
        if self.primary not in keys:
            self.primary = None
        if self.secondary not in (None, "") and self.secondary not in keys:
            self.secondary = ""
        self._keys = keys
        self.p_cb.configure(values=labels)
        self.s_cb.configure(values=[NONE_LABEL] + labels)
        self.p_cb.set(labels[keys.index(self.primary)] if self.primary else "")
        if self.secondary is None:
            self.s_cb.set(NONE_LABEL)
        elif self.secondary == "":
            self.s_cb.set("")
        else:
            self.s_cb.set(labels[keys.index(self.secondary)])

    def _on_select(self) -> None:
        pi = self.p_cb.current()
        self.primary = self._keys[pi] if pi >= 0 else None
        si = self.s_cb.current()
        if si == 0:
            self.secondary = None
        elif si > 0:
            self.secondary = self._keys[si - 1]
        self.frame.validate()

    def set_error(self, p_err: bool, s_err: bool) -> None:
        self.p_cb.configure(style="Error.TCombobox" if p_err else "TCombobox")
        self.s_cb.configure(style="Error.TCombobox" if s_err else "TCombobox")

    def to_setting(self) -> GraphSetting:
        return GraphSetting(self.primary, self.secondary)


class AdjustFrame(ttk.Frame):
    def __init__(self, app: App, uart: UartData, logger: LoggerData):
        super().__init__(app, padding=8)
        self.app = app
        self.uart = uart
        self.logger = logger
        self.result: MergeResult | None = None
        self.item_rows: list[ItemRow] = []
        self.graph_rows: list[GraphRow] = []
        self._recalc_job = None
        self._busy = False

        items, self.unit_warnings = cfgmod.build_items(app.cfg, uart, logger)
        graphs = cfgmod.build_graphs(app.cfg, items)
        self.elapsed_var = tk.StringVar(value=cfgmod.elapsed_label(app.cfg))

        # エラー表示とボタンは常に見えるよう下に固定し、それ以外をスクロール領域に入れる
        self._build_bottom_bar()
        self.page = ScrollableFrame(self)
        self.page.pack(fill="both", expand=True)
        self.body = self.page.inner
        self.body.configure(padding=(0, 0, 8, 0))

        self._build_result_section()
        self._build_interval_section()
        self._build_items_section()
        self._build_graph_section()
        self._build_output_section()

        self.set_items(items, graphs)
        self.recalc()

    # ---- 読み込み結果 ------------------------------------------------------
    def _build_result_section(self) -> None:
        box = ttk.LabelFrame(self.body, text="読み込み結果", padding=6)
        box.pack(fill="x")
        u = self.uart
        self.uart_info = ttk.Label(box)
        self.uart_info.pack(anchor="w")
        lg = self.logger
        ttk.Label(
            box,
            text=f"ロガー: {lg.model or '(モデル不明)'}  {lg.row_count:,}行  {len(lg.channels)}CH"
            + (f"  測定間隔 {lg.interval_ms:g}ms" if lg.interval_ms else "  測定間隔 不明"),
        ).pack(anchor="w")
        gap_line = ttk.Frame(box)
        gap_line.pack(anchor="w", fill="x")
        self.gap_label = tk.Label(gap_line, fg="#b36b00")
        self.gap_label.pack(side="left")
        self.gap_btn = ttk.Button(gap_line, text="詳細", command=self.show_gap_detail)
        for w in self.unit_warnings:
            tk.Label(box, text="⚠ " + w, fg="#b36b00", justify="left", wraplength=860).pack(anchor="w")
        self._uart_summary_base = (
            f"UART : {u.row_count:,}行  {u.timestamps[0]:%H:%M:%S}〜{u.timestamps[-1]:%H:%M:%S}  "
            f"スキップ{u.skipped:,}行"
        )

    # ---- 間隔 --------------------------------------------------------------
    def _build_interval_section(self) -> None:
        box = ttk.LabelFrame(self.body, text="間隔", padding=6)
        box.pack(fill="x", pady=(6, 0))
        line = ttk.Frame(box)
        line.pack(anchor="w")
        self.uart_int = tk.StringVar(value=str(default_uart_interval(self.uart.timestamps)))
        li = self.logger.interval_ms
        self.logger_int = tk.StringVar(value=_fmt_num(li) if li else "")
        ttk.Label(line, text="UART間隔").pack(side="left")
        self.uart_int_e = tk.Entry(line, textvariable=self.uart_int, width=8)
        self.uart_int_e.pack(side="left", padx=4)
        ttk.Label(line, text="ms     ロガー間隔").pack(side="left")
        self.logger_int_e = tk.Entry(line, textvariable=self.logger_int, width=8)
        self.logger_int_e.pack(side="left", padx=4)
        ttk.Label(line, text="ms").pack(side="left")
        self.mode_label = tk.Label(box, anchor="w")
        self.mode_label.pack(anchor="w", pady=(4, 0))
        self._mode_bg = self.mode_label.cget("bg")
        self.uart_int.trace_add("write", lambda *_: self.schedule_recalc())
        self.logger_int.trace_add("write", lambda *_: self.schedule_recalc())

    # ---- 出力する項目 ------------------------------------------------------
    def _build_items_section(self) -> None:
        box = ttk.LabelFrame(self.body, text="出力する項目", padding=6)
        box.pack(fill="x", pady=(6, 0))
        self.items_inner = ttk.Frame(box)
        self.items_inner.pack(fill="x")

    def set_items(self, items: list[ItemSetting], graphs: list[GraphSetting]) -> None:
        inner = self.items_inner
        for w in inner.winfo_children():
            w.destroy()
        self.item_rows = []
        inner.columnconfigure(5, weight=1)  # 余白は右端に寄せ、ラベル欄は伸ばさない
        for c, text in enumerate(["採用", "元の列名", "ラベル", "係数", "オフセット"]):
            ttk.Label(inner, text=text, font=("", 9, "bold")).grid(row=0, column=c, sticky="w", padx=4)

        fixed = ttk.Checkbutton(inner)  # 経過時間は常に出力（外せない）
        fixed.state(["!alternate", "selected", "disabled"])
        fixed.grid(row=1, column=0)
        ttk.Label(inner, text="経過時間(s)").grid(row=1, column=1, sticky="w", padx=4)
        self.elapsed_e = tk.Entry(inner, textvariable=self.elapsed_var, width=LABEL_WIDTH)
        self.elapsed_e.grid(row=1, column=2, sticky="w", padx=4, pady=1)
        ttk.Label(inner, text="―").grid(row=1, column=3)
        ttk.Label(inner, text="―").grid(row=1, column=4)
        if not getattr(self, "_elapsed_traced", False):
            self.elapsed_var.trace_add("write", lambda *_: self.on_items_changed())
            self._elapsed_traced = True

        r = 2
        logger_header_done = False
        for it in items:
            if it.source == SOURCE_LOGGER and not logger_header_done:
                ttk.Label(inner, text="── ロガー ──").grid(row=r, column=0, columnspan=5, sticky="w", pady=(4, 0))
                r += 1
                logger_header_done = True
            self.item_rows.append(ItemRow(inner, r, it, self.on_items_changed))
            r += 1

        for g in self.graph_rows:
            g.destroy()
        self.graph_rows = [GraphRow(self, g) for g in graphs]
        self._regrid_graphs()
        self.on_items_changed()

    # ---- グラフ ------------------------------------------------------------
    def _build_graph_section(self) -> None:
        box = ttk.LabelFrame(self.body, text="グラフ", padding=6)
        box.pack(fill="x", pady=(6, 0))
        self.graph_inner = ttk.Frame(box)
        self.graph_inner.pack(anchor="w")
        for c, text in enumerate(["No", "第1軸", "第2軸"]):
            ttk.Label(self.graph_inner, text=text, font=("", 9, "bold")).grid(row=0, column=c, sticky="w", padx=4)
        ttk.Button(box, text="＋グラフを追加", command=self.add_graph).pack(anchor="w", pady=(4, 0))

    def _regrid_graphs(self) -> None:
        for i, g in enumerate(self.graph_rows, start=1):
            g.grid(i)

    def add_graph(self) -> None:
        g = GraphRow(self, GraphSetting(None, None))
        self.graph_rows.append(g)
        self._regrid_graphs()
        g.refresh(self._graph_choices())
        self.validate()

    def remove_graph(self, g: GraphRow) -> None:
        g.destroy()
        self.graph_rows.remove(g)
        self._regrid_graphs()
        self.validate()

    def _graph_choices(self) -> list[tuple[str, str]]:
        return [(r.key, r.label.get().strip()) for r in self.item_rows if r.enabled.get()]

    # ---- 出力先・ボタン ----------------------------------------------------
    def _build_output_section(self) -> None:
        box = ttk.LabelFrame(self.body, text="出力先", padding=6)
        box.pack(fill="x", pady=(6, 0))
        box.columnconfigure(1, weight=1)
        self.folder = tk.StringVar(value=os.path.dirname(os.path.abspath(self.uart.path)))
        self.filename = tk.StringVar(value=cfgmod.make_output_filename(self.app.cfg, self.uart.timestamps[0]))
        ttk.Label(box, text="フォルダ").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.folder_e = tk.Entry(box, textvariable=self.folder)
        self.folder_e.grid(row=0, column=1, sticky="ew")
        ttk.Button(box, text="参照", command=self.browse_folder).grid(row=0, column=2, padx=(6, 0))
        ttk.Label(box, text="ファイル名").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(4, 0))
        self.filename_e = tk.Entry(box, textvariable=self.filename)
        self.filename_e.grid(row=1, column=1, sticky="ew", pady=(4, 0))
        self.folder.trace_add("write", lambda *_: self.validate())
        self.filename.trace_add("write", lambda *_: self.validate())

    def _build_bottom_bar(self) -> None:
        btns = ttk.Frame(self)
        btns.pack(side="bottom", fill="x", pady=(4, 0))
        self.error_label = tk.Label(self, fg="#c00000", justify="left", anchor="w", wraplength=880)
        self.error_label.pack(side="bottom", fill="x", pady=(4, 0))
        self.reset_btn = ttk.Button(btns, text="初期値に戻す", command=self.reset_defaults)
        self.reset_btn.pack(side="left")
        self.export_btn = ttk.Button(btns, text="Excelに出力", command=self.export)
        self.export_btn.pack(side="right")
        self.back_btn = ttk.Button(btns, text="戻る", command=self.app.show_select)
        self.back_btn.pack(side="right", padx=6)

    def browse_folder(self) -> None:
        cur = self.folder.get()
        path = filedialog.askdirectory(parent=self, initialdir=cur if os.path.isdir(cur) else None)
        if path:
            self.folder.set(os.path.normpath(path))

    # ---- 再計算・検査 ------------------------------------------------------
    def intervals(self) -> tuple[float | None, float | None]:
        u = _parse_float(self.uart_int.get())
        lg = _parse_float(self.logger_int.get())
        return (u if u is not None and u > 0 else None, lg if lg is not None and lg > 0 else None)

    def schedule_recalc(self) -> None:
        if self._recalc_job is not None:
            self.after_cancel(self._recalc_job)
        self._recalc_job = self.after(250, self.recalc)

    def recalc(self) -> None:
        self._recalc_job = None
        u, lg = self.intervals()
        self.uart_int_e.configure(bg=OK_BG if u else ERR_BG)
        self.logger_int_e.configure(bg=OK_BG if lg else ERR_BG)
        if not u or not lg:
            self.result = None
            self.mode_label.configure(text="間隔を正の数値で入力してください。", fg="#c00000", bg=self._mode_bg)
            self.uart_info.configure(text=self._uart_summary_base)
            self.gap_label.configure(text="")
            self.gap_btn.pack_forget()
            self.validate()
            return
        tl = build_uart_timeline(self.uart, u)
        self.result = merge(self.uart, self.logger, u, lg, timeline=tl)
        self.uart_info.configure(text=f"{self._uart_summary_base} / 重複{tl.duplicates:,}行")
        if tl.gaps:
            self.gap_label.configure(
                text=f"⚠ 欠落：{len(tl.gaps):,}箇所（{tl.gap_rows:,}行を空欄で補完）", fg="#b36b00"
            )
            self.gap_btn.pack(side="left", padx=6)
        else:
            self.gap_label.configure(text="欠落なし", fg="black")
            self.gap_btn.pack_forget()

        n = len(self.result.rows)
        rows = max_sheet_rows(n, self.uart, self.logger)
        level = check_row_limit(rows, self.app.cfg["row_warn_threshold"], self.app.cfg["row_limit"])
        text = f"→ {describe_mode(u, lg)} / 出力予定 {n:,}行"
        if level != "ok":
            text += f"（最大シート {rows:,}行）"
        color = {"ok": "black", ROWS_WARN: "#b38f00", ROWS_ERROR: "#c00000"}[level]
        bg = {"ok": self._mode_bg, ROWS_WARN: "#fff3a0", ROWS_ERROR: "#ffc0c0"}[level]
        self.mode_label.configure(text=text, fg=color, bg=bg)
        self.validate()

    def on_items_changed(self) -> None:
        for r in self.item_rows:
            state = "normal" if r.enabled.get() else "disabled"
            for e in (r.label_e, r.coef_e, r.offset_e):
                e.configure(state=state)
        choices = self._graph_choices()
        for g in self.graph_rows:
            g.refresh(choices)
        self.validate()

    def validate(self) -> list[str]:
        errors: list[str] = []
        u, lg = self.intervals()
        if not u or not lg:
            errors.append("間隔が正しくありません。")

        # 出力する項目
        elabel = self.elapsed_var.get().strip()
        labels: dict[str, int] = {}
        if elabel:
            labels[elabel] = 1
        for r in self.item_rows:
            if r.enabled.get() and r.label.get().strip():
                labels[r.label.get().strip()] = labels.get(r.label.get().strip(), 0) + 1
        self.elapsed_e.configure(bg=OK_BG if elabel and labels.get(elabel, 0) <= 1 else ERR_BG)
        if not elabel:
            errors.append("経過時間のラベルが空欄です。")
        for r in self.item_rows:
            if not r.enabled.get():
                for e in (r.label_e, r.coef_e, r.offset_e):
                    e.configure(disabledbackground=DISABLED_BG)
                continue
            lab = r.label.get().strip()
            bad_label = not lab or labels.get(lab, 0) > 1
            r.label_e.configure(bg=ERR_BG if bad_label else OK_BG)
            if not lab:
                errors.append(f"{r.key}：ラベルが空欄です。")
            elif labels.get(lab, 0) > 1:
                errors.append(f"{r.key}：ラベル「{lab}」が重複しています。")
            for e, v, name in ((r.coef_e, r.coef, "係数"), (r.offset_e, r.offset, "オフセット")):
                ok = _parse_float(v.get()) is not None
                e.configure(bg=OK_BG if ok else ERR_BG)
                if not ok:
                    errors.append(f"{r.key}：{name}が数値ではありません。")
        if elabel and labels.get(elabel, 0) > 1:
            errors.append(f"経過時間のラベル「{elabel}」が他の項目と重複しています。")

        # グラフ
        items = [r.to_setting() for r in self.item_rows]
        graphs = [g.to_setting() for g in self.graph_rows]
        enabled = {it.key for it in items if it.enabled}
        for g in self.graph_rows:
            p_err = not g.primary or g.primary not in enabled
            s_err = g.secondary == "" or (g.secondary is not None and (g.secondary not in enabled or g.secondary == g.primary))
            g.set_error(p_err, s_err)
        errors.extend(validate_graphs(graphs, items))

        # 出力先
        folder = self.folder.get().strip()
        folder_ok = bool(folder) and os.path.isdir(folder)
        self.folder_e.configure(bg=OK_BG if folder_ok else ERR_BG)
        if not folder_ok:
            errors.append("出力先フォルダがありません。")
        fname = self.filename.get().strip()
        bad = invalid_filename_chars(fname)
        fname_ok = bool(fname) and not bad
        self.filename_e.configure(bg=OK_BG if fname_ok else ERR_BG)
        if not fname:
            errors.append("ファイル名が空欄です。")
        elif bad:
            errors.append("ファイル名に使えない文字が含まれています：" + " ".join(bad))

        if self.result is not None and not self.result.rows:
            errors.append("出力するデータ行がありません（間隔の設定を確認してください）。")

        shown = errors[:6] + ([f"…ほか{len(errors) - 6}件"] if len(errors) > 6 else [])
        self.error_label.configure(text="\n".join(shown))
        if not self._busy:
            self.export_btn.state(["disabled"] if errors else ["!disabled"])
        return errors

    # ---- 詳細・初期値に戻す ------------------------------------------------
    def show_gap_detail(self) -> None:
        if self.result is None:
            return
        win = tk.Toplevel(self)
        win.title("欠落の詳細")
        win.geometry("640x360")
        cols = ("no", "prev", "next", "missing")
        tv = ttk.Treeview(win, columns=cols, show="headings")
        for c, text, w in (
            ("no", "No", 50),
            ("prev", "欠落直前の行（行番号・時刻）", 230),
            ("next", "欠落直後の行（行番号・時刻）", 230),
            ("missing", "補完行数", 80),
        ):
            tv.heading(c, text=text)
            tv.column(c, width=w, anchor="w" if c in ("prev", "next") else "e")
        for i, g in enumerate(self.result.timeline.gaps, start=1):
            tv.insert(
                "", "end",
                values=(
                    i,
                    f"{g.prev_line}行目  {g.prev_time:%H:%M:%S.%f}"[:-3],
                    f"{g.next_line}行目  {g.next_time:%H:%M:%S.%f}"[:-3],
                    g.missing,
                ),
            )
        vsb = ttk.Scrollbar(win, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        tv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def reset_defaults(self) -> None:
        if not messagebox.askokcancel(
            APP_TITLE, "出力する項目とグラフの設定を初期値に戻します。よろしいですか？", parent=self
        ):
            return
        items, _ = cfgmod.build_items(self.app.cfg, self.uart, self.logger, use_saved=False)
        graphs = cfgmod.build_graphs(self.app.cfg, items, use_saved=False)
        self.elapsed_var.set(cfgmod.elapsed_label(self.app.cfg, use_saved=False))
        self.set_items(items, graphs)

    # ---- 出力 --------------------------------------------------------------
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

        progress = ProgressDialog(self)
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
        note = f"\n\n※設定を保存できませんでした（{res.config_error}）" if res.config_error else ""
        DoneDialog(self, f"Excel ファイルを出力しました。\n{path}{note}", os.path.dirname(path))

    def _on_export_error(self, path: str, e: Exception) -> None:
        if isinstance(e, PermissionError):
            msg = (
                f"ファイルを保存できませんでした。\n{path}\n\n"
                "ファイルが Excel で開かれている場合は閉じてから、もう一度出力してください。"
            )
        else:
            msg = f"出力に失敗しました。\n{type(e).__name__}: {e}"
        messagebox.showerror(APP_TITLE, msg, parent=self)


class ProgressDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("出力中")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # 出力中は閉じさせない
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        self.msg = ttk.Label(frm, text="出力中…", width=48)
        self.msg.pack(anchor="w")
        pb = ttk.Progressbar(frm, mode="indeterminate", length=360)
        pb.pack(pady=(8, 0))
        pb.start(15)
        self.grab_set()

    def set_message(self, m: str) -> None:
        self.msg.configure(text=m)

    def close(self) -> None:
        self.grab_release()
        self.destroy()


class DoneDialog(tk.Toplevel):
    def __init__(self, parent, message: str, folder: str):
        super().__init__(parent)
        self.title(APP_TITLE)
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=message, justify="left", wraplength=520).pack(anchor="w")
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="OK", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="フォルダを開く", command=lambda: (open_folder(folder), self.destroy())).pack(
            side="right", padx=6
        )
        self.grab_set()
