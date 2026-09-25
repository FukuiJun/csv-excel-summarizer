"""補助ウィンドウ：カラーパレット・欠落の詳細・出力中・完了（仕様書 5 章）。"""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from tkinter import colorchooser, messagebox, ttk

from gui_theme import Theme
from gui_widgets import CheckIcon, frame, hline, label, sized

APP_TITLE = "ロガーCSV・UART CSV 統合ツール"


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


def _center_on(win: tk.Toplevel, parent: tk.Misc) -> None:
    win.update_idletasks()
    top = parent.winfo_toplevel()
    x = top.winfo_rootx() + (top.winfo_width() - win.winfo_reqwidth()) // 2
    y = top.winfo_rooty() + (top.winfo_height() - win.winfo_reqheight()) // 3
    win.geometry(f"+{max(0, x)}+{max(0, y)}")


def _grab_later(win: tk.Toplevel) -> None:
    """表示されてから grab する（表示前の grab は Windows で失敗することがある）。"""

    def grab():
        try:
            if win.winfo_exists():
                win.grab_set()
        except tk.TclError:
            pass

    win.after(50, grab)


# ---------------------------------------------------------------------------
# 5.1 カラーパレット
# ---------------------------------------------------------------------------


class ColorPalette(tk.Toplevel):
    """色見本の真下に出るカラーパレット。選んだ色は result（"RRGGBB"）に入り、on_pick に渡される。"""

    THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]
    STANDARD = ["C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C0", "002060", "7030A0"]

    def __init__(self, anchor: tk.Widget, current: str, on_pick=None, th: Theme | None = None):
        top = anchor.winfo_toplevel()
        super().__init__(top)
        self.th = th or getattr(top, "theme", None)
        th = self.th
        self.result: str | None = None
        self._on_pick = on_pick
        self._anchor = anchor
        self.withdraw()
        # 枠なしのポップアップ。transient にすると Windows では親の状態によって表示されないことがあるため、
        # 代わりに最前面表示にする（外側クリック・Esc・色の選択で閉じる）
        self.overrideredirect(True)
        try:
            self.attributes("-topmost", True)
        except tk.TclError:  # pragma: no cover
            pass
        current = (current or "").upper()

        outer = tk.Frame(self, bd=0, highlightthickness=1, highlightbackground="#C8C8C8", highlightcolor="#C8C8C8")
        th.paint(outer, bg="panel")
        outer.pack(fill="both", expand=True)
        body = frame(outer, th, "panel")
        body.pack(fill="both", expand=True, padx=th.px(12), pady=th.px(12))

        # 上部：現在の色
        head = frame(body, th, "panel")
        head.pack(fill="x")
        sw = tk.Frame(head, width=th.px(36), height=th.px(24), bd=0, highlightthickness=1, bg="#" + (current or "FFFFFF"))
        th.paint(sw, highlightbackground="swatch_border")
        sw.pack(side="left")
        t = frame(head, th, "panel")
        t.pack(side="left", padx=(th.px(10), 0))
        label(t, th, "現在の色", font="caption", fg="text2", bg="panel").pack(anchor="w")
        label(t, th, f"#{current}", font="body_bold", bg="panel").pack(anchor="w")
        label(head, th, "Esc で閉じる", font="caption", fg="text2", bg="panel").pack(side="right", anchor="n")
        hline(body, th, "row_line").pack(fill="x", pady=(th.px(10), th.px(10)))

        label(body, th, "テーマの色", font="heading", fg="text2", bg="panel").pack(anchor="w")
        g1 = frame(body, th, "panel")
        g1.pack(anchor="w", pady=(th.px(6), 0))
        for c, base in enumerate(self.THEME):
            self._swatch(g1, base, current, 0, c)
        g2 = frame(body, th, "panel")
        g2.pack(anchor="w", pady=(th.px(8), 0))
        for c, base in enumerate(self.THEME):
            for r, f in enumerate((0.8, 0.6, 0.4, -0.25, -0.5)):
                self._swatch(g2, self._shade(base, f), current, r, c, vgap=0)
        label(body, th, "標準の色", font="heading", fg="text2", bg="panel").pack(anchor="w", pady=(th.px(10), 0))
        g3 = frame(body, th, "panel")
        g3.pack(anchor="w", pady=(th.px(6), 0))
        for c, col in enumerate(self.STANDARD):
            self._swatch(g3, col, current, 0, c)
        ttk.Button(body, text="その他の色…", command=self._more).pack(fill="x", pady=(th.px(10), 0))

        # クリックした色見本の真下に表示する（位置の計算は _place）
        self._place()
        self.bind("<Escape>", lambda e: self.close())
        self.bind("<Button-1>", self._click_outside, add="+")
        self._set_anchor_highlight(True)
        self.deiconify()
        self._place()  # Windows では表示前の位置指定が効かないことがあるため、表示後にもう一度
        self.after_idle(self._place)
        self.lift()
        self.focus_force()
        _grab_later(self)

    def popup_position(self) -> tuple[int, int]:
        """色見本の真下（入らなければ真上）。アプリのウィンドウからはみ出す分は内側に寄せる。

        モニターが複数ある場合もあるので、画面（プライマリモニター）の範囲ではなく
        アプリのウィンドウを基準にする（座標が負の値になることもある）。
        """
        a = self._anchor
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        top = a.winfo_toplevel()
        left, right = top.winfo_rootx(), top.winfo_rootx() + top.winfo_width()
        upper, lower = top.winfo_rooty(), top.winfo_rooty() + top.winfo_height()
        x = a.winfo_rootx()
        if x + w > right:
            x = max(left, a.winfo_rootx() + a.winfo_width() - w)  # 色見本の右端にそろえる
        y = a.winfo_rooty() + a.winfo_height() + 2
        if y + h > lower and a.winfo_rooty() - h - 2 >= upper:
            y = a.winfo_rooty() - h - 2  # 下に入らなければ上に出す
        elif y + h > lower:
            y = max(upper, lower - h)  # 上にも入らなければウィンドウの下端にそろえる
        return x, y

    def _place(self) -> None:
        if self.winfo_exists():
            x, y = self.popup_position()
            self.geometry(f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}{x:+d}{y:+d}")

    @staticmethod
    def _shade(hex_color: str, f: float) -> str:
        """f>0 は白に近づけ（明るく）、f<0 は黒に近づける（暗く）。"""
        rgb = [int(hex_color[i : i + 2], 16) for i in (0, 2, 4)]
        if f >= 0:
            rgb = [round(v + (255 - v) * f) for v in rgb]
        else:
            rgb = [round(v * (1 + f)) for v in rgb]
        return "".join(f"{v:02X}" for v in rgb)

    def _swatch(self, parent, color: str, current: str, row: int, col: int, vgap: int | None = None) -> None:
        th = self.th
        s = th.px(22)
        gap = th.px(4)
        sel = color.upper() == current
        # 選択中の色は accent 2px の外枠（1px 空ける）
        holder = tk.Frame(parent, bd=0, highlightthickness=2 if sel else 0)
        th.paint(holder, bg="panel", highlightbackground="accent", highlightcolor="accent")
        holder.grid(row=row, column=col, padx=(0, gap), pady=(0, gap if vgap is None else vgap))
        sw = tk.Frame(holder, width=s, height=s, bg="#" + color, cursor="hand2", bd=0,
                      highlightthickness=1, highlightbackground="#D0D0D0")
        sw.pack(padx=1 if sel else 0, pady=1 if sel else 0)
        sw.bind("<Button-1>", lambda e, c=color: self._pick(c))

    def _set_anchor_highlight(self, on: bool) -> None:
        a = self._anchor
        try:
            if on:
                self._anchor_prev = (a.cget("highlightthickness"), a.cget("highlightbackground"))
                a.configure(highlightthickness=2, highlightbackground=self.th.c["accent"])
            elif getattr(self, "_anchor_prev", None):
                a.configure(highlightthickness=self._anchor_prev[0], highlightbackground=self._anchor_prev[1])
        except tk.TclError:
            pass

    def _click_outside(self, e) -> None:
        x, y = e.x_root, e.y_root
        if not (self.winfo_rootx() <= x < self.winfo_rootx() + self.winfo_width()
                and self.winfo_rooty() <= y < self.winfo_rooty() + self.winfo_height()):
            self.close()

    def close(self) -> None:
        self._set_anchor_highlight(False)
        try:
            self.grab_release()
        except tk.TclError:
            pass
        if self.winfo_exists():
            self.destroy()

    def _pick(self, color: str) -> None:
        self.result = color.upper()
        self.close()
        if self._on_pick is not None:
            self._on_pick(self.result)

    def _more(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        rgb, hx = colorchooser.askcolor(parent=self, title="その他の色")
        if hx:
            self._pick(hx.lstrip("#"))
        elif self.winfo_exists():
            _grab_later(self)


# ---------------------------------------------------------------------------
# 5.2 欠落の詳細
# ---------------------------------------------------------------------------


class GapDetailDialog(tk.Toplevel):
    def __init__(self, parent, th: Theme, gaps):
        super().__init__(parent)
        self.title("欠落の詳細")
        self.transient(parent.winfo_toplevel())
        th.paint(self, bg="bg")
        body = frame(self, th, "bg")
        body.pack(fill="both", expand=True, padx=th.px(16), pady=th.px(14))
        head = frame(body, th, "bg")
        head.pack(fill="x")
        label(head, th, f"欠落 {len(gaps):,}箇所", font="title").pack(side="left")
        total = sum(g.missing for g in gaps)
        label(head, th, f"補完した行は合計 {total:,}行（Excel では空欄）", font="caption", fg="text2").pack(
            side="left", padx=(th.px(10), 0), pady=(th.px(3), 0)
        )
        tf = frame(body, th, "bg")
        tf.pack(fill="both", expand=True, pady=(th.px(12), th.px(12)))
        cols = ("no", "prev", "next", "missing")
        self.tree = tv = ttk.Treeview(tf, columns=cols, show="headings", height=6)
        for c, text, w, anchor in (
            ("no", "No", 56, "e"),
            ("prev", "欠落直前の行（行番号・時刻）", 260, "w"),
            ("next", "欠落直後の行（行番号・時刻）", 260, "w"),
            ("missing", "補完行数", 96, "e"),
        ):
            tv.heading(c, text=text, anchor=anchor)
            tv.column(c, width=th.px(w), anchor=anchor, stretch=c in ("prev", "next"))
        for i, g in enumerate(gaps, start=1):
            tv.insert(
                "", "end",
                values=(i, f"{g.prev_line}行目   {g.prev_time:%H:%M:%S.%f}"[:-3],
                        f"{g.next_line}行目   {g.next_time:%H:%M:%S.%f}"[:-3], g.missing),
            )
        vsb = ttk.Scrollbar(tf, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        tv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        btn_box, _ = sized(body, th, lambda p: ttk.Button(p, text="閉じる", command=self.destroy), 96, 32)
        btn_box.pack(anchor="e")
        self.bind("<Escape>", lambda e: self.destroy())
        _center_on(self, parent)


# ---------------------------------------------------------------------------
# 5.3 出力中
# ---------------------------------------------------------------------------


class ProgressDialog(tk.Toplevel):
    def __init__(self, parent, th: Theme | None = None):
        super().__init__(parent)
        th = th or parent.winfo_toplevel().theme
        self.title("出力中")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # 出力中は閉じさせない
        th.paint(self, bg="bg")
        body = frame(self, th, "bg")
        body.pack(fill="both", expand=True, padx=th.px(20), pady=th.px(20))
        label(body, th, "Excel ファイルを出力しています", font="body_bold").pack(anchor="w")
        self.msg = label(body, th, "出力中…", font="caption", fg="text2", width=52)
        self.msg.pack(anchor="w", pady=(th.px(6), th.px(12)))
        pb = ttk.Progressbar(body, mode="indeterminate", length=th.px(320))
        pb.pack(fill="x")
        pb.start(15)
        _center_on(self, parent)
        _grab_later(self)

    def set_message(self, m: str) -> None:
        self.msg.configure(text=m)

    def close(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()


# ---------------------------------------------------------------------------
# 5.4 完了
# ---------------------------------------------------------------------------


class DoneDialog(tk.Toplevel):
    def __init__(self, parent, path: str, note: str = "", th: Theme | None = None):
        super().__init__(parent)
        th = th or parent.winfo_toplevel().theme
        folder = os.path.dirname(path)
        self.title("完了")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        th.paint(self, bg="bg")
        body = frame(self, th, "bg")
        body.pack(fill="both", expand=True, padx=th.px(20), pady=th.px(20))
        CheckIcon(body, th, 32, "success", "bg", filled=True).pack(side="left", anchor="n", padx=(0, th.px(14)))
        txt = frame(body, th, "bg")
        txt.pack(side="left", fill="both", expand=True)
        label(txt, th, "Excel ファイルを出力しました", font="title").pack(anchor="w")
        label(txt, th, os.path.basename(path), font="body_bold").pack(anchor="w", pady=(th.px(4), 0))
        label(txt, th, f"保存先：{folder}{os.sep}", font="caption", fg="text2", wraplength=th.px(420)).pack(
            anchor="w", pady=(th.px(2), 0)
        )
        if note:
            label(txt, th, note, font="caption", fg="warning", wraplength=th.px(420)).pack(anchor="w", pady=(th.px(6), 0))

        hline(self, th, "border").pack(fill="x")
        band = frame(self, th, "surface")
        band.pack(fill="x")
        inner = frame(band, th, "surface")
        inner.pack(fill="x", padx=th.px(16), pady=th.px(12))

        def open_and_close():
            open_folder(folder)
            self.destroy()

        ok_box, _ = sized(inner, th, lambda p: ttk.Button(p, text="OK", command=self.destroy), 80, 32, bg="surface")
        ok_box.pack(side="right")
        open_box, self.open_btn = sized(
            inner, th, lambda p: ttk.Button(p, text="フォルダを開く", style="Accent.TButton", command=open_and_close),
            132, 32, bg="surface",
        )
        open_box.pack(side="right", padx=(0, th.px(8)))
        # Enter の既定は［フォルダを開く］
        self.bind("<Return>", lambda e: open_and_close())
        self.bind("<Escape>", lambda e: self.destroy())
        _center_on(self, parent)
        self.open_btn.focus_set()
        _grab_later(self)
