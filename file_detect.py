"""ドロップされたファイルが UART CSV かロガー CSV かを中身から判別する。"""

from __future__ import annotations

from uart_reader import TIMESTAMP_COLUMN

KIND_UART = "uart"
KIND_LOGGER = "logger"

_HEAD_BYTES = 64 * 1024  # 先頭だけ読む（大きなファイルでもすぐ判別できるように）


def detect_csv_kind(path: str) -> str | None:
    """UART CSV なら "uart"、ロガー CSV なら "logger"、どちらでもなければ None。"""
    try:
        with open(path, "rb") as f:
            head = f.read(_HEAD_BYTES)
    except OSError:
        return None
    if not head:
        return None

    first_line = head.split(b"\n", 1)[0]
    try:
        if TIMESTAMP_COLUMN in first_line.decode("utf-8-sig"):
            return KIND_UART
    except UnicodeDecodeError:
        pass

    text = head.decode("cp932", errors="replace")
    keys = {line.split(",", 1)[0].strip().strip('"') for line in text.splitlines()}
    if "測定値" in keys or ("モデル" in keys and "測定間隔" in keys):
        return KIND_LOGGER
    return None
