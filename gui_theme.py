"""配色・フォント・寸法の適用（design_tokens.py を使う）。

- sv-ttk が塗る ttk 部品はテーマ任せ。tk.Frame / tk.Label / tk.Canvas / tk.Button など
  sv-ttk が塗らない部品は paint() で色の役割（トークン名）を登録し、テーマ切替時に apply() で塗り直す。
"""

from __future__ import annotations

import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

import design_tokens as T

try:  # 見た目を Windows 11 風にするテーマ（なくても動く）
    import sv_ttk
except ImportError:  # pragma: no cover
    sv_ttk = None

# 独自スタイルを用意する背景（トークン名）。この上に置く ttk.Checkbutton / ttk.Label 用
SURFACES = ["bg", "surface", "panel", "tile", "error_row", "warning_row"]


def windows_prefers_dark() -> bool:
    """Windows の「アプリモード」がダークか。Windows 以外・取得できない場合は False。"""
    if not sys.platform.startswith("win"):
        return False
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        ) as k:
            v, _ = winreg.QueryValueEx(k, "AppsUseLightTheme")
        return int(v) == 0
    except OSError:
        return False


def resolve_theme(setting: str | None) -> str:
    """config の theme（auto / light / dark）から実際のテーマ名を決める。"""
    if setting in ("light", "dark"):
        return setting
    return "dark" if windows_prefers_dark() else "light"


class Theme:
    def __init__(self, root: tk.Tk, name: str = "light"):
        self.root = root
        self.name = name
        self.c = dict(T.LIGHT)
        self._painted: dict[str, tuple[tk.Misc, dict[str, str]]] = {}
        self._scale = 1.0
        try:
            self._scale = float(root.tk.call("tk", "scaling")) / (96 / 72)
        except tk.TclError:  # pragma: no cover
            pass
        families = set(tkfont.families(root))
        fam = T.FONT_FAMILY if T.FONT_FAMILY in families else (
            T.FONT_FALLBACK if T.FONT_FALLBACK in families else tkfont.nametofont("TkDefaultFont").actual("family")
        )
        self.family = fam
        self.fonts = {k: (fam, -self.px(abs(size)), *rest) for k, (_, size, *rest) in T.FONTS.items()}
        self.style = ttk.Style(root)
        self.apply(name)

    # ---- 寸法 ------------------------------------------------------------
    def px(self, v: int) -> int:
        """96dpi 基準の px を現在の表示倍率に換算する。"""
        return max(1, round(v * self._scale)) if v else 0

    def size(self, key: str) -> int:
        return self.px(T.SIZE[key])

    # ---- 色の登録・適用 -----------------------------------------------------
    def paint(self, widget: tk.Misc, **roles: str) -> tk.Misc:
        """widget の色オプション（bg / fg / highlightbackground など）にトークン名を割り当てて塗る。"""
        old = self._painted.get(str(widget))
        if old is not None and old[0] is widget:
            roles = {**old[1], **roles}  # 同じ部品を塗り直すときは役割を上書き
        self._painted[str(widget)] = (widget, roles)
        self._paint_one(widget, roles)
        return widget

    def _paint_one(self, widget: tk.Misc, roles: dict[str, str]) -> None:
        opts = {}
        for opt, token in roles.items():
            color = self.c.get(token, token)  # トークン名でなければ色そのもの
            opts[opt] = color
            if opt == "bg" and isinstance(widget, (tk.Button, tk.Label)):
                opts.setdefault("activebackground", color)
        try:
            widget.configure(**opts)
        except tk.TclError:
            pass

    def apply(self, name: str) -> None:
        """テーマを切り替え、登録済みの tk 部品と独自スタイルの色を塗り直す（1 か所にまとめた再適用関数）。"""
        self.name = name
        self.c = dict(T.DARK if name == "dark" else T.LIGHT)
        if sv_ttk is not None:
            sv_ttk.set_theme(name, self.root)
            try:
                self._define_error_combobox()
            except tk.TclError:  # pragma: no cover  テーマの内部が変わった場合
                self.style.map("Error.TCombobox", fieldbackground=[("readonly", self.c["error_bg"])])
        else:  # pragma: no cover
            self.style.map("Error.TCombobox", fieldbackground=[("readonly", self.c["error_bg"])])
        self._configure_styles()
        self.root.configure(bg=self.c["bg"])
        alive = {}
        for key, (w, roles) in self._painted.items():
            try:
                if w.winfo_exists():
                    self._paint_one(w, roles)
                    alive[key] = (w, roles)
            except tk.TclError:
                pass
        self._painted = alive

    def _configure_styles(self) -> None:
        s = self.style
        f = self.fonts
        s.configure(".", font=f["body"])
        s.configure("TButton", font=f["body"])
        s.configure("Accent.TButton", font=f["body_bold"])
        s.configure("Small.TButton", font=f["small_body"], padding=(self.px(8), self.px(1)))
        s.configure("TNotebook.Tab", font=f["body"], padding=(self.px(16), self.px(6)))
        s.map("TNotebook.Tab", font=[("selected", f["body_bold"])])
        s.configure("Treeview", font=f["body"], rowheight=self.px(30))
        s.configure("Treeview.Heading", font=f["heading"])
        for surf in SURFACES:
            bg = self.c[surf]
            s.configure(f"{surf}.TCheckbutton", background=bg)
            s.configure(f"{surf}.TRadiobutton", background=bg)
            s.configure(f"{surf}.TFrame", background=bg)
        s.configure("Error.TCombobox", font=f["body"])

    def _define_error_combobox(self) -> None:
        """エラー時のプルダウン（赤枠）のスタイル。読み取り専用のプルダウンはテーマの invalid 表示が効かないため。"""
        ns = "sv_dark" if self.name == "dark" else "sv_light"
        theme = f"sun-valley-{self.name}"
        self.root.tk.eval(
            f"""
            ttk::style theme settings {theme} {{
              catch {{
                ttk::style element create ErrCombobox.field image [list \\
                  $::ttk::theme::{ns}::I(textbox-error) \\
                  disabled $::ttk::theme::{ns}::I(textbox-dis)] -border 5
              }}
              ttk::style layout Error.TCombobox {{
                ErrCombobox.field -sticky nsew -children {{
                  Combobox.arrow -side right -sticky ns
                  Combobox.padding -sticky nsew -children {{
                    Combobox.textarea -sticky nsew
                  }}
                }}
              }}
              ttk::style configure Error.TCombobox -padding {{6 1 0 2}}
            }}
            """
        )


def mark(widget, ok: bool) -> None:
    """入力欄・プルダウンのエラー表示（赤下線）を切り替える。"""
    widget.state(["!invalid"] if ok else ["invalid"])
    if isinstance(widget, ttk.Combobox):
        widget.configure(style="TCombobox" if ok else "Error.TCombobox")
