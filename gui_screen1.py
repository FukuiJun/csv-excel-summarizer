"""画面1：ファイル選択（仕様書 2 章）。"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import design_tokens as T
from file_detect import KIND_LOGGER, KIND_UART, detect_csv_kind
from gui_dialogs import APP_TITLE
from gui_widgets import CheckIcon, frame, hline, label, sized
from logger_reader import LoggerFormatError, read_logger_csv
from uart_reader import UartFormatError, read_uart_csv

try:
    from tkinterdnd2 import DND_FILES
except ImportError:  # pragma: no cover
    DND_FILES = None

CSV_TYPES = [("CSV ファイル", "*.csv *.CSV"), ("すべてのファイル", "*.*")]


class DropZone(tk.Canvas):
    """破線の枠のドロップ領域（高さ 220）。ドラッグ中は accent の実線にする。"""

    def __init__(self, parent, th):
        super().__init__(parent, height=th.size("drop_h"), highlightthickness=0, bd=0)
        th.paint(self, bg="bg")
        self.th = th
        self.active = False
        self.bind("<Configure>", lambda e: self.draw())

    def set_active(self, v: bool) -> None:
        self.active = v
        self.draw()

    def draw(self) -> None:
        th, c = self.th, self.th.c
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10:
            return
        color = c["accent"] if self.active else c["drop_border"]
        self.create_rectangle(2, 2, w - 2, h - 2, fill=c["drop_bg"], outline=color, width=2,
                              dash=() if self.active else (6, 4))
        cx = w / 2
        s = th.px(48)
        top = h / 2 - th.px(58)
        a = c["accent"]
        lw = max(2, th.px(2))
        # 上向き矢印＋トレイ
        self.create_line(cx, top + s * 0.12, cx, top + s * 0.62, fill=a, width=lw, capstyle="round")
        self.create_line(cx - s * 0.2, top + s * 0.3, cx, top + s * 0.1, cx + s * 0.2, top + s * 0.3,
                         fill=a, width=lw, capstyle="round", joinstyle="round")
        self.create_line(cx - s * 0.36, top + s * 0.55, cx - s * 0.36, top + s * 0.8, cx + s * 0.36, top + s * 0.8,
                         cx + s * 0.36, top + s * 0.55, fill=a, width=lw, capstyle="round", joinstyle="round")
        self.create_text(cx, top + s + th.px(22), text="CSV ファイルをここにドラッグ＆ドロップ",
                         font=th.fonts["drop_title"], fill=c["text"])
        self.create_text(cx, top + s + th.px(52),
                         text="2つ同時にドロップできます。中身から UART／ロガーを自動で振り分けます",
                         font=th.fonts["small_body"], fill=c["text2"])


class FileSelectFrame(tk.Frame):
    def __init__(self, app):
        super().__init__(app.content, bd=0, highlightthickness=0)
        th = self.th = app.theme
        th.paint(self, bg="bg")
        self.app = app

        # ---- 本体：幅 760 の列を中央寄せ ----
        body = frame(self, th, "bg")
        body.pack(fill="both", expand=True)
        col = frame(body, th, "bg")
        col.place(relx=0.5, y=th.size("screen1_top_pad"), anchor="n", width=th.size("screen1_content_w"))
        self._col = col

        self.drop = DropZone(col, th)
        self.drop.pack(fill="x")

        rows = frame(col, th, "bg")
        rows.pack(fill="x", pady=(th.px(24), 0))
        rows.columnconfigure(2, weight=1)
        for c, w in ((0, 28), (1, 132), (3, 88)):
            rows.columnconfigure(c, minsize=th.px(w))
        gap = th.px(12)
        self.uart_icon, self.uart_entry = self._file_row(
            rows, 0, "UART CSV", "マイコンからの受信データ", app.uart_path, KIND_UART, gap)
        self.logger_icon, self.logger_entry = self._file_row(
            rows, 1, "ロガー CSV", "GRAPHTEC GL240", app.logger_path, KIND_LOGGER, gap)
        label(col, th, "欄の上にドロップするとその欄に、枠内の他の場所なら中身で自動判別して入ります。",
              font="caption", fg="text2").pack(anchor="w", padx=(th.px(40), 0), pady=(th.px(14), 0))

        # ---- フッター ----
        hline(self, th, "border").pack(fill="x", side="bottom", before=body)
        foot = frame(self, th, "surface", height=th.size("footer_h"))
        foot.pack(fill="x", side="bottom", before=body)
        foot.pack_propagate(False)
        inner = frame(foot, th, "surface")
        inner.pack(fill="both", expand=True, padx=th.size("page_pad_x"))
        box, self.load_btn = sized(
            inner, th, lambda p: ttk.Button(p, text="読み込み  →", style="Accent.TButton", command=self.load),
            T.SIZE["btn_primary_w"], T.SIZE["btn_h"], bg="surface",
        )
        box.pack(side="right", pady=th.px(16))
        self.hint = label(inner, th, "UART・ロガーのどちらか1つだけでも読み込めます", font="caption", fg="text2", bg="surface")
        self.hint.pack(side="right", padx=(0, th.px(16)))

        # ---- ドラッグ＆ドロップ ----
        if app.dnd_enabled and DND_FILES is not None:
            for w, kind in ((self.uart_entry, KIND_UART), (self.logger_entry, KIND_LOGGER)):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", lambda e, k=kind: self._on_drop(e, k))
            for w in (self.drop, self, body, col, rows):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", lambda e: self._on_drop(e, None))
            self.drop.dnd_bind("<<DropEnter>>", lambda e: (self.drop.set_active(True), e.action)[1])
            self.drop.dnd_bind("<<DropLeave>>", lambda e: (self.drop.set_active(False), e.action)[1])
        self.drop.bind("<Button-1>", lambda e: self.browse("uart", app.uart_path) if not app.uart_path.get()
                       else self.browse("logger", app.logger_path))

        app.uart_path.trace_add("write", lambda *_: self.update_state())
        app.logger_path.trace_add("write", lambda *_: self.update_state())
        self.update_state()

    def _file_row(self, parent, r, title, sub, var, kind, gap):
        th = self.th
        icon = CheckIcon(parent, th, 20, "success", "bg")
        icon.grid(row=r, column=0, sticky="w", pady=(0 if r == 0 else gap, 0))
        head = frame(parent, th, "bg")
        head.grid(row=r, column=1, sticky="w", padx=(gap, 0), pady=(0 if r == 0 else gap, 0))
        label(head, th, title, font="body_bold").pack(anchor="w")
        label(head, th, sub, font="caption", fg="text2").pack(anchor="w")
        entry = ttk.Entry(parent, textvariable=var)
        entry.grid(row=r, column=2, sticky="ew", padx=(gap, 0), pady=(0 if r == 0 else gap, 0), ipady=th.px(2))
        box, _ = sized(parent, th, lambda p: ttk.Button(p, text="参照…", command=lambda: self.browse(kind, var)), 88, 32)
        box.grid(row=r, column=3, sticky="e", padx=(gap, 0), pady=(0 if r == 0 else gap, 0))
        return icon, entry

    # ---- ドロップ・参照 ------------------------------------------------------
    def _on_drop(self, event, kind: str | None) -> str:
        self.drop.set_active(False)
        self.set_dropped_files(list(self.tk.splitlist(event.data)), kind)
        return event.action

    def set_dropped_files(self, paths: list[str], kind: str | None = None) -> None:
        """ドロップされたファイルを UART／ロガーの欄に設定する。kind が None なら中身から判別する。"""
        unknown = []
        for path in paths:
            path = os.path.normpath(path)
            k = kind if (kind is not None and len(paths) == 1) else detect_csv_kind(path)
            if k == KIND_UART:
                self.app.uart_path.set(path)
                self.app.dirs["uart"] = os.path.dirname(path)
            elif k == KIND_LOGGER:
                self.app.logger_path.set(path)
                self.app.dirs["logger"] = os.path.dirname(path)
            else:
                unknown.append(os.path.basename(path))
        if unknown:
            messagebox.showwarning(
                APP_TITLE,
                "UART CSV・ロガー CSV のどちらか判別できませんでした：\n" + "\n".join(unknown)
                + "\n\nUART CSV 欄・ロガー CSV 欄の上に直接ドロップするか、［参照］で選んでください。",
                parent=self,
            )

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
        u = bool(self.app.uart_path.get().strip())
        lg = bool(self.app.logger_path.get().strip())
        self.uart_icon.set_visible(u)
        self.logger_icon.set_visible(lg)
        # どちらか 1 つだけでも読み込める（片方だけの Excel を作る）
        self.load_btn.state(["!disabled"] if (u or lg) else ["disabled"])
        if u and lg:
            self.hint.pack_forget()
        else:
            self.hint.configure(text="UART・ロガーのどちらか1つだけでも読み込めます" if not (u or lg)
                                else ("UART だけで読み込みます（ロガーも入れるとまとめます）" if u
                                      else "ロガーだけで読み込みます（UART も入れるとまとめます）"))
            if not self.hint.winfo_manager():
                self.hint.pack(side="right", padx=(0, self.th.px(16)))

    def load(self) -> None:
        up = self.app.uart_path.get().strip().strip('"')
        lp = self.app.logger_path.get().strip().strip('"')
        self.app.config(cursor="watch")
        self.update_idletasks()
        uart = logger = None
        try:
            try:
                uart = read_uart_csv(up) if up else None
            except FileNotFoundError:
                raise UartFormatError(f"UART CSV が見つかりません:\n{up}")
            except UartFormatError:
                raise
            except Exception as e:  # noqa: BLE001
                raise UartFormatError(f"UART CSV を読み込めません: {e}")
            try:
                logger = read_logger_csv(lp) if lp else None
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
        if uart is None and logger is None:
            return
        self.app.show_adjust(uart, logger)

