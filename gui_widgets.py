"""画面で共通に使う部品（スクロール領域・ラベル・区切り線・アイコンなど）。"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from gui_theme import Theme


def label(parent, th: Theme, text: str = "", *, font: str = "body", fg: str = "text", bg: str = "bg", **kw) -> tk.Label:
    """テーマの色で塗る tk.Label。bg / fg はトークン名。"""
    lb = tk.Label(parent, text=text, font=th.fonts[font], anchor=kw.pop("anchor", "w"), justify=kw.pop("justify", "left"),
                  bd=0, padx=0, pady=0, **kw)
    th.paint(lb, bg=bg, fg=fg)
    return lb


def frame(parent, th: Theme, bg: str = "bg", **kw) -> tk.Frame:
    f = tk.Frame(parent, bd=0, highlightthickness=0, **kw)
    th.paint(f, bg=bg)
    return f


def hline(parent, th: Theme, color: str = "border", height: int = 1) -> tk.Frame:
    """1px の区切り線。"""
    f = tk.Frame(parent, height=height, bd=0, highlightthickness=0)
    th.paint(f, bg=color)
    return f


def boxed(parent, th: Theme, bg: str, border: str, pad=(8, 12)) -> tuple[tk.Frame, tk.Frame]:
    """1px の枠付きの箱。戻り値：(外枠, 中身を置く Frame)。"""
    outer = tk.Frame(parent, bd=0, highlightthickness=1)
    th.paint(outer, bg=bg, highlightbackground=border, highlightcolor=border)
    inner = frame(outer, th, bg)
    inner.pack(fill="both", expand=True, padx=th.px(pad[1]), pady=th.px(pad[0]))
    return outer, inner


def sized(parent, th: Theme, widget_factory, width: int, height: int, bg: str = "bg"):
    """px 指定の大きさの箱に部品を入れる（ttk.Button の幅・高さを px でそろえるため）。"""
    box = frame(parent, th, bg, width=th.px(width), height=th.px(height))
    box.pack_propagate(False)
    w = widget_factory(box)
    w.pack(fill="both", expand=True)
    return box, w


class CheckIcon(tk.Canvas):
    """丸にチェック（✓）のアイコン。filled=True なら塗りつぶし＋白いチェック。"""

    def __init__(self, parent, th: Theme, size: int = 20, color: str = "success", bg: str = "bg", filled: bool = False):
        s = th.px(size)
        super().__init__(parent, width=s, height=s, highlightthickness=0, bd=0)
        th.paint(self, bg=bg)
        self.th, self.s, self.color, self.filled = th, s, color, filled
        self.visible = True
        self.draw()

    def set_visible(self, v: bool) -> None:
        self.visible = v
        self.draw()

    def draw(self) -> None:
        self.delete("all")
        if not self.visible:
            return
        s, c = self.s, self.th.c[self.color]
        w = max(1, round(s / 12))
        if self.filled:
            self.create_oval(1, 1, s - 1, s - 1, fill=c, outline=c)
            tick = self.th.c["on_accent"] if self.color == "accent" else "#FFFFFF"
        else:
            self.create_oval(w, w, s - w, s - w, outline=c, width=w + 0.5)
            tick = c
        self.create_line(s * 0.3, s * 0.52, s * 0.45, s * 0.66, s * 0.72, s * 0.36, fill=tick, width=w + 0.8,
                         capstyle="round", joinstyle="round")


class StepIcon(tk.Canvas):
    """ヘッダーの手順表示の丸（24×24）。state: current / done / todo"""

    def __init__(self, parent, th: Theme, number: int):
        s = th.px(24)
        super().__init__(parent, width=s, height=s, highlightthickness=0, bd=0)
        th.paint(self, bg="bg")
        self.th, self.s, self.number = th, s, number
        self.state_ = "todo"

    def set_state(self, state: str) -> None:
        self.state_ = state
        self.draw()

    def draw(self) -> None:
        self.delete("all")
        s, c = self.s, self.th.c
        self.configure(bg=c["bg"])
        if self.state_ == "current":
            self.create_oval(1, 1, s - 1, s - 1, fill=c["accent"], outline=c["accent"])
            self.create_text(s / 2, s / 2, text=str(self.number), fill=c["on_accent"], font=self.th.fonts["heading"])
        else:
            self.create_oval(1, 1, s - 1, s - 1, outline="#C8C8C8" if self.th.name == "light" else "#5A5A5A", width=1.2)
            if self.state_ == "done":
                self.create_line(s * 0.3, s * 0.52, s * 0.45, s * 0.66, s * 0.72, s * 0.36, fill=c["text2"], width=1.6)
            else:
                self.create_text(s / 2, s / 2, text=str(self.number), fill=c["text2"], font=self.th.fonts["caption"])


class ScrollArea(tk.Frame):
    """縦スクロール領域（Canvas ＋ 内側 Frame ＋ ttk.Scrollbar）。

    マウスホイールは App が「カーソルの下にある ScrollArea」だけに振り分ける。
    """

    registry: "list[ScrollArea]" = []

    def __init__(self, parent, th: Theme, bg: str = "bg", pad_x: int = 0, pad_y: int = 0):
        super().__init__(parent, bd=0, highlightthickness=0)
        th.paint(self, bg=bg)
        self.th = th
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        th.paint(self.canvas, bg=bg)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = frame(self.canvas, th, bg)
        self._win = self.canvas.create_window((th.px(pad_x), th.px(pad_y)), window=self.inner, anchor="nw")
        self._pad = (th.px(pad_x), th.px(pad_y))
        self.inner.bind("<Configure>", lambda e: self._update_region())
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=self._on_scroll)
        self.canvas.pack(side="left", fill="both", expand=True)
        self._vsb_shown = False
        ScrollArea.registry.append(self)
        self.bind("<Destroy>", lambda e: e.widget is self and self in ScrollArea.registry and ScrollArea.registry.remove(self))

    def _on_scroll(self, first, last) -> None:
        self.vsb.set(first, last)
        need = not (float(first) <= 0.0 and float(last) >= 1.0)
        if need != self._vsb_shown:
            self._vsb_shown = need
            if need:
                self.vsb.pack(side="right", fill="y", before=self.canvas)
            else:
                self.vsb.pack_forget()

    def _on_canvas_configure(self, e) -> None:
        self.canvas.itemconfigure(self._win, width=max(1, e.width - 2 * self._pad[0]))
        self._update_region()

    def _update_region(self) -> None:
        h = max(self.inner.winfo_reqheight() + 2 * self._pad[1], self.canvas.winfo_height())
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), h))

    def can_scroll(self) -> bool:
        return self.inner.winfo_reqheight() + 2 * self._pad[1] > self.canvas.winfo_height()

    def scroll(self, step: int) -> None:
        if self.can_scroll():
            self.canvas.yview_scroll(step, "units")

    def see(self, widget: tk.Misc) -> None:
        """widget が見える位置までスクロールする。"""
        self.update_idletasks()
        if not self.can_scroll():
            return
        y = widget.winfo_rooty() - self.inner.winfo_rooty() + self._pad[1]
        h = widget.winfo_height()
        total = max(1, self.inner.winfo_reqheight() + 2 * self._pad[1])
        view_h = self.canvas.winfo_height()
        top = self.canvas.canvasy(0)
        if y < top or y + h > top + view_h:
            self.canvas.yview_moveto(max(0.0, (y - view_h / 3) / total))

    @classmethod
    def under(cls, widget: tk.Misc | None) -> "ScrollArea | None":
        """widget を含む一番内側の ScrollArea。"""
        w = widget
        while w is not None:
            if isinstance(w, ScrollArea):
                return w
            w = getattr(w, "master", None)
        return None


class TabBar(tk.Frame):
    """モックアップどおりのタブ（選択中は白いパネルとつながる見た目、件数・エラーのバッジ付き）。

    ttk.Notebook と同じ呼び方（add / select / index / tab）ができる。
    タブを切り替えると <<NotebookTabChanged>> を発生させる。
    """

    def __init__(self, parent, th: Theme):
        super().__init__(parent, bd=0, highlightthickness=0)
        th.paint(self, bg="bg")
        self.th = th
        self.row = frame(self, th, "bg", height=th.size("tab_h"))
        self.row.pack(fill="x")
        self.hint = label(self.row, th, "", font="caption", fg="text2")
        self.hint.pack(side="right")
        self.panel = tk.Frame(self, bd=0, highlightthickness=1)
        th.paint(self.panel, bg="panel", highlightbackground="border", highlightcolor="border")
        self.panel.pack(fill="both", expand=True)
        self._joint = frame(self, th, "panel")  # 選択中タブとパネルの境目の線を隠す
        self.tabs: list[dict] = []
        self.current = 0

    def add(self, page: tk.Misc, text: str = "") -> None:
        th = self.th
        i = len(self.tabs)
        tab = tk.Frame(self.row, bd=0, highlightthickness=1, cursor="hand2", takefocus=1)
        tab.pack(side="left", fill="y")
        inner = frame(tab, th, "bg")
        inner.pack(fill="both", expand=True, padx=th.px(18), pady=th.px(8))
        name = label(inner, th, "", cursor="hand2")
        name.pack(side="left")
        badge = tk.Label(inner, font=th.fonts["caption"], padx=th.px(7), pady=0, bd=0, cursor="hand2")
        badge.pack(side="left", padx=(th.px(8), 0))
        for w in (tab, inner, name, badge):
            w.bind("<Button-1>", lambda e, k=i: self.select(k))
        tab.bind("<Return>", lambda e, k=i: self.select(k))
        self.tabs.append({"frame": tab, "inner": inner, "name": name, "badge": badge, "page": page,
                          "text": "", "count": "", "error": False})
        self.tab(i, text=text)
        if i == 0:
            self.select(0)
        else:
            self._refresh()

    def select(self, i=None):
        if i is None:
            return self.current
        if isinstance(i, str):
            i = int(i)
        changed = i != self.current
        self.current = i
        for k, t in enumerate(self.tabs):
            t["page"].pack_forget()
        self.tabs[i]["page"].pack(in_=self.panel, fill="both", expand=True)
        self._refresh()
        if changed:
            self.event_generate("<<NotebookTabChanged>>")
        return None

    def index(self, i) -> int:
        return int(i)

    def tab(self, i: int, text: str | None = None, count: str | None = None, error: bool | None = None) -> None:
        """text：タブ名。count：バッジの文字（例「9 / 12」「エラー 3」）。error：バッジを赤にする。"""
        t = self.tabs[i]
        if text is not None:
            t["text"] = text
        if count is not None:
            t["count"] = count
        if error is not None:
            t["error"] = error
        self._refresh()

    def _refresh(self) -> None:
        th = self.th
        for k, t in enumerate(self.tabs):
            sel = k == self.current
            bg = "panel" if sel else "bg"
            th.paint(t["frame"], bg=bg, highlightbackground="border" if sel else "bg",
                     highlightcolor="border" if sel else "bg")
            th.paint(t["inner"], bg=bg)
            th.paint(t["name"], bg=bg, fg="text" if sel else "text2")
            t["name"].configure(text=t["text"], font=th.fonts["body_bold" if sel else "body"])
            b = t["badge"]
            if t["count"]:
                if t["error"]:
                    th.paint(b, bg="error", fg="#FFFFFF" if th.name == "light" else "#000000")
                    b.configure(font=(th.family, -th.px(12), "bold"))
                else:
                    th.paint(b, bg="badge_bg", fg="text2")
                    b.configure(font=th.fonts["caption"])
                b.configure(text=t["count"])
                if not b.winfo_manager():
                    b.pack(side="left", padx=(th.px(8), 0))
            else:
                b.pack_forget()
        self.after_idle(self._place_joint)

    def _place_joint(self) -> None:
        if not self.tabs or not self.winfo_exists():
            return
        t = self.tabs[self.current]["frame"]
        self._joint.place(x=t.winfo_x() + 1, y=self.row.winfo_height() - 1, width=max(1, t.winfo_width() - 2), height=2)
        self._joint.lift()
