"""アプリのバージョン。

exe のビルド時（GitHub Actions・build.bat）に _version.py が作られ、そこからバージョンを読む。
_version.py がない（python main.py で動かしている）ときは「開発版」。
"""

from __future__ import annotations

DEV_VERSION = "開発版"


def get_version() -> str:
    try:
        from _version import VERSION  # type: ignore[import-not-found]
    except ImportError:
        return DEV_VERSION
    return str(VERSION) or DEV_VERSION


def window_title(base: str) -> str:
    """ウィンドウのタイトル（例「ロガーCSV・UART CSV 統合ツール v1.2.0」）。"""
    return f"{base} {get_version()}"
