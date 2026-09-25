"""ロガーCSV・UART CSV 統合ツールの起動スクリプト。

    python main.py
"""

import sys


def _enable_dpi_awareness() -> None:
    """Windows の表示倍率（125%・150% など）で文字がぼやけないようにする。"""
    if sys.platform.startswith("win"):
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass


def main() -> None:
    _enable_dpi_awareness()
    from gui import App

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
