"""GUI（tkinter / ttk ＋ sv-ttk）。

画面構成はデザイン資料 docs/claude_design/handoff/IMPLEMENTATION_SPEC.md に従う。
- ヘッダー（手順表示）・本体・フッター
- 画面1：ファイル選択（gui_screen1.py）
- 画面2：出力内容の調整（gui_screen2.py）
- 補助ウィンドウ（gui_dialogs.py）
"""

from __future__ import annotations

import os
import sys
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import messagebox

import config as cfgmod
import design_tokens as T
from gui_dialogs import (
    APP_TITLE,
    ColorPalette,
    DoneDialog,
    GapDetailDialog,
    ProgressDialog,
    open_folder,
)
from gui_screen1 import FileSelectFrame
from gui_screen2 import AdjustFrame, GraphRow, ItemRow
from gui_theme import Theme, resolve_theme
from gui_widgets import ScrollArea, StepIcon, frame, hline, label
from logger_reader import LoggerData
from uart_reader import UartData

# テスト・外部から gui.ColorPalette などで参照できるように公開する
__all__ = ["App", "AdjustFrame", "FileSelectFrame", "ItemRow", "GraphRow", "ColorPalette", "DoneDialog",
           "GapDetailDialog", "ProgressDialog", "open_folder", "messagebox", "tk"]

try:  # ファイルのドラッグ＆ドロップ（なくても動く）
    from tkinterdnd2 import TkinterDnD
except ImportError:  # pragma: no cover
    TkinterDnD = None


def resource_path(*parts: str) -> str:
    """同梱ファイルの場所（PyInstaller の exe では展開先、通常はこのファイルのフォルダ）。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def _win32_set_icons(root: tk.Tk, ico: str) -> int | None:
    """Windows：表示倍率に合ったサイズのアイコンを ico から読み、ウィンドウの大・小アイコンに設定する。

    Tk の iconbitmap は 32px 前後の 1 枚しか設定しないため、タスクバー（大アイコンを使う）で
    拡大・縮小されてぼやける。ここでは ico からちょうどのサイズを Windows に選ばせて設定し直す。
    戻り値は設定に使った DPI（失敗したら None）。
    """
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.LoadImageW.restype = wintypes.HANDLE
    user32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT, ctypes.c_int, ctypes.c_int,
                                  wintypes.UINT]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    hwnd = int(root.wm_frame(), 16)  # タイトルバーを持つ外側のウィンドウ
    try:
        dpi = user32.GetDpiForWindow(hwnd) or 96
    except AttributeError:  # Windows 10 より前
        dpi = 96

    def metric(index: int) -> int:
        try:
            return user32.GetSystemMetricsForDpi(index, dpi)
        except AttributeError:
            return round(user32.GetSystemMetrics(index) * dpi / 96)

    SM_CXICON, SM_CXSMICON = 11, 49
    IMAGE_ICON, LR_LOADFROMFILE = 1, 0x10
    WM_SETICON, ICON_SMALL, ICON_BIG = 0x80, 0, 1
    handles = []
    for which, idx in ((ICON_BIG, SM_CXICON), (ICON_SMALL, SM_CXSMICON)):
        size = metric(idx)
        h = user32.LoadImageW(None, ico, IMAGE_ICON, size, size, LR_LOADFROMFILE)
        if not h:
            return None
        user32.SendMessageW(hwnd, WM_SETICON, which, h)
        handles.append(h)
    root._win_icon_handles = handles  # ウィンドウが使っている間は保持する
    return dpi


def _current_dpi(root: tk.Tk) -> int | None:
    try:
        import ctypes

        return ctypes.windll.user32.GetDpiForWindow(int(root.wm_frame(), 16))
    except Exception:  # noqa: BLE001
        return None


def set_window_icon(root: tk.Tk) -> bool:
    """タイトルバー・タスクバーのアイコンを設定する（補助ウィンドウにも適用）。失敗しても起動は続ける。"""
    try:
        if sys.platform.startswith("win"):
            ico = resource_path("assets", "LogMerger.ico")
            if os.path.exists(ico):
                root.iconbitmap(default=ico)  # 補助ウィンドウにも適用される

                def apply(event=None):
                    # 表示された後に、表示倍率に合ったサイズで設定し直す（倍率が変わったときも）
                    try:
                        dpi = _win32_set_icons(root, ico) if (
                            event is None or getattr(root, "_icon_dpi", None) != _current_dpi(root)) else None
                        if dpi:
                            root._icon_dpi = dpi
                    except Exception:  # noqa: BLE001  失敗しても iconbitmap のアイコンのまま
                        pass

                root.after(50, apply)
                root.bind("<Configure>", lambda e: e.widget is root and apply(e), add="+")
                return True
        pngs = [resource_path("assets", f"LogMerger_{s}.png") for s in (16, 32, 48, 256)]
        pngs = [p for p in pngs if os.path.exists(p)]
        if pngs:
            root._icon_images = [tk.PhotoImage(master=root, file=p) for p in pngs]  # 参照を保持
            root.iconphoto(True, *root._icon_images)
            return True
    except tk.TclError:
        pass
    return False


class App(tk.Tk):
    def __init__(self, config_path: str | None = None):
        super().__init__()
        self.title(APP_TITLE)
        self.icon_set = set_window_icon(self)
        w = T.WINDOW

        self.config_path = config_path or cfgmod.default_config_path()
        res = cfgmod.load_config(self.config_path)
        self.cfg = res.config
        # 「参照」ダイアログの初期フォルダ（出力成功時に config.json へ保存）
        self.dirs = dict(self.cfg.get("last_dirs", {}))

        self.theme = Theme(self, resolve_theme(self.cfg.get("theme")))
        th = self.theme
        self.geometry(f"{th.px(w['init_w'])}x{th.px(w['init_h'])}")
        self.minsize(th.px(w["min_w"]), th.px(w["min_h"]))

        # ドラッグ＆ドロップ（tkdnd）を使えるようにする。読み込めない環境では D&D なしで動く
        self.dnd_enabled = False
        if TkinterDnD is not None:
            try:
                TkinterDnD._require(self)
                self.dnd_enabled = True
            except Exception:  # noqa: BLE001  pragma: no cover
                pass

        self.uart_path = tk.StringVar()
        self.logger_path = tk.StringVar()

        self._build_header()
        self.content = frame(self, th, "bg")
        self.content.pack(fill="both", expand=True)

        self.select_frame = FileSelectFrame(self)
        self.adjust_frame: AdjustFrame | None = None
        self.show_select()
        self._bind_wheel()

        if res.status == "broken":
            self.after(
                100,
                lambda: messagebox.showwarning(
                    APP_TITLE, f"config.json が壊れていたため、既定値で作り直しました。\n{res.message}"
                ),
            )
        elif res.message:
            self.after(100, lambda: messagebox.showwarning(APP_TITLE, res.message))

    # ---- ヘッダー（手順表示） --------------------------------------------------
    def _build_header(self) -> None:
        th = self.theme
        head = frame(self, th, "bg", height=th.size("header_h"))
        head.pack(fill="x")
        head.pack_propagate(False)
        inner = frame(head, th, "bg")
        inner.pack(fill="both", expand=True, padx=th.size("page_pad_x"))
        self._steps = []
        for i, name in enumerate(("ファイル選択", "出力内容の調整"), start=1):
            if i > 1:
                label(inner, th, "›", font="title", fg="#A0A0A0").pack(side="left", padx=th.px(10))
            icon = StepIcon(inner, th, i)
            icon.pack(side="left")
            lb = label(inner, th, name)
            lb.pack(side="left", padx=(th.px(8), 0))
            self._steps.append((icon, lb))
        self.header_note = label(inner, th, "", font="caption", fg="text2")
        self.header_note.pack(side="right")
        hline(self, th, "border").pack(fill="x")

    def set_step(self, step: int, note: str) -> None:
        th = self.theme
        for i, (icon, lb) in enumerate(self._steps, start=1):
            if i == step:
                icon.set_state("current")
                lb.configure(font=th.fonts["title"])
                th.paint(lb, fg="text")
            else:
                icon.set_state("done" if i < step else "todo")
                lb.configure(font=th.fonts["body"])
                th.paint(lb, fg="text2")
        self.header_note.configure(text=note)

    # ---- 画面切替 -------------------------------------------------------------
    def show_select(self) -> None:
        if self.adjust_frame is not None:
            self.adjust_frame.destroy()
            self.adjust_frame = None
        self.select_frame.pack(fill="both", expand=True)
        self.set_step(1, "設定は前回の値を引き継ぎます")

    def show_adjust(self, uart: UartData, logger: LoggerData) -> None:
        self.select_frame.pack_forget()
        self.adjust_frame = AdjustFrame(self, uart, logger)
        self.adjust_frame.pack(fill="both", expand=True)
        self.set_step(2, f"{os.path.basename(uart.path)} ＋ {os.path.basename(logger.path)}")

    def set_theme(self, name: str) -> None:
        """テーマを切り替える（tk 系部品の色もまとめて塗り直す）。"""
        self.theme.apply(name)
        for icon, _ in self._steps:
            icon.draw()

    # ---- マウスホイール：カーソルの下の領域だけ動かす ----------------------------
    def _bind_wheel(self) -> None:
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, self._on_wheel, add="+")
            # プルダウンの上でホイールを回しても選択値を変えず、領域のスクロールだけ行う
            self.bind_class("TCombobox", seq, lambda e: (self._on_wheel(e), "break")[1])

    def _on_wheel(self, e) -> None:
        w = e.widget
        if isinstance(w, str) or "popdown" in str(w):
            return
        try:
            if w.winfo_toplevel() is not self:
                return
            target = self.winfo_containing(e.x_root, e.y_root) or w
        except (tk.TclError, KeyError):
            return
        area = ScrollArea.under(target)
        if area is None:
            return
        if getattr(e, "num", None) == 4:
            step = -1
        elif getattr(e, "num", None) == 5:
            step = 1
        else:
            step = int(-e.delta / 120) or (-1 if e.delta > 0 else 1)
        area.scroll(step * 2)

    # ---- 想定外のエラー -------------------------------------------------------
    def report_callback_exception(self, exc, val, tb) -> None:
        """画面操作中の想定外のエラーを表示し、error.log に記録する（exe ではコンソールが見えないため）。"""
        text = "".join(traceback.format_exception(exc, val, tb))
        log_path = os.path.join(os.path.dirname(os.path.abspath(self.config_path)), "error.log")
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S}\n{text}\n")
            where = f"\n\n詳細は {log_path} に記録しました。"
        except OSError:
            where = ""
        messagebox.showerror(APP_TITLE, f"エラーが発生しました。\n{exc.__name__}: {val}{where}", parent=self)
