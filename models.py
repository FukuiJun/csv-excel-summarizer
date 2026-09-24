"""画面2で調整する設定（出力項目・グラフ）のデータ構造。"""

from __future__ import annotations

from dataclasses import dataclass

SOURCE_UART = "uart"
SOURCE_LOGGER = "logger"


@dataclass
class ItemSetting:
    """出力する項目1つ分の設定（仕様 5.2.3）。"""

    key: str  # UARTは元の列名、ロガーはCH番号（例: voltage_mV, CH1）
    source: str  # SOURCE_UART / SOURCE_LOGGER
    enabled: bool
    label: str
    coef: float = 1.0
    offset: float = 0.0
    unit: str = ""  # ロガーCHの単位

    def convert(self, value: float | None) -> float | None:
        if value is None:
            return None
        return value * self.coef + self.offset


@dataclass
class GraphSetting:
    """グラフ1つ分の設定（仕様 5.2.4）。キーは ItemSetting.key。"""

    primary: str | None
    secondary: str | None = None
