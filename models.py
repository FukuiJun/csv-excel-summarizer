"""画面2で調整する設定（出力項目・グラフ）のデータ構造。"""

from __future__ import annotations

from dataclasses import dataclass

SOURCE_UART = "uart"
SOURCE_LOGGER = "logger"

# グラフ横軸で「経過時間(s)」（pc_timestamp から求めた経過秒数）を表すキー
ELAPSED_KEY = "@elapsed"
# グラフ横軸の初期値
DEFAULT_X_KEY = "elapsed_ms"
# グラフ横軸の時間の単位（ms に対する倍率）と、時間の項目の元の単位
TIME_UNITS = {"ms": 1, "s": 1000, "min": 60_000, "h": 3_600_000}
TIME_KEYS = {"elapsed_ms": "ms", ELAPSED_KEY: "s"}
# グラフ1つの縦軸1本あたりに載せられる系列数
MAX_SERIES_PER_AXIS = 2

# グラフの線の既定色（Excel の標準の系列色の並び）。項目の並び順に割り当てる
DEFAULT_COLORS = [
    "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
    "264478", "9E480E", "636363", "997300", "255E91", "43682B",
]


def default_color(index: int) -> str:
    return DEFAULT_COLORS[index % len(DEFAULT_COLORS)]


def is_color(v) -> bool:
    """"RRGGBB" 形式の色か。"""
    return isinstance(v, str) and len(v) == 6 and all(c in "0123456789abcdefABCDEF" for c in v)


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
    color: str = DEFAULT_COLORS[0]  # グラフの線の色（"RRGGBB"）

    def convert(self, value: float | None) -> float | None:
        if value is None:
            return None
        return value * self.coef + self.offset


@dataclass
class GraphSetting:
    """グラフ1つ分の設定（仕様 5.2.4）。キーは ItemSetting.key（横軸は ELAPSED_KEY も可）。

    primary（第1軸）・secondary（第2軸）はそれぞれ MAX_SERIES_PER_AXIS 個の欄のリスト。
    - x と primary[0] は必須：None は未選択（エラー）
    - それ以外の欄：None は「なし」、"" は未選択（採用が外れた項目など。エラー）
    """

    primary: list[str | None]
    secondary: list[str | None]
    x: str | None = DEFAULT_X_KEY
    x_unit: str | None = None  # 横軸の時間の単位（ms/s/min/h）。None は元の単位のまま

    def effective_x_unit(self) -> str | None:
        """横軸に換算列を使う場合の単位（横軸が時間の項目で、元と違う単位のときだけ）。"""
        native = TIME_KEYS.get(self.x or "")
        if native and self.x_unit in TIME_UNITS and self.x_unit != native:
            return self.x_unit
        return None

    def __post_init__(self) -> None:
        self.primary = _pad(self.primary)
        self.secondary = _pad(self.secondary)

    def series_keys(self) -> list[str]:
        """選ばれている縦軸の項目（第1軸→第2軸の順、「なし」・未選択を除く）。"""
        return [k for k in self.primary + self.secondary if k]


def _pad(v) -> list[str | None]:
    if v is None or isinstance(v, str):
        v = [v]
    v = list(v)[:MAX_SERIES_PER_AXIS]
    return v + [None] * (MAX_SERIES_PER_AXIS - len(v))
