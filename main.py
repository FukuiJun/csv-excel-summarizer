"""ロガーCSV・UART CSV 統合ツールの起動スクリプト。

    python main.py
"""

from gui import App


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
