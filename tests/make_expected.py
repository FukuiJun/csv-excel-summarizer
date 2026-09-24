"""期待出力 tests/data/expected_log40.xlsx を作り直すスクリプト。

    python tests/make_expected.py [UART CSV] [ロガー CSV] [出力xlsx]

仕様 9章：期待出力は同時に計測した実データから作り直すこと。
実データを tests/data/ に置き、引数で指定して実行し、出力を目視で確認してからコミットする。
既定値（config.json の既定設定・CSVから求めた間隔）で出力する。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config as cfgmod  # noqa: E402
from excel_writer import write_workbook  # noqa: E402
from helpers import EXPECTED_XLSX, SAMPLE_LOGGER, SAMPLE_UART  # noqa: E402
from logger_reader import read_logger_csv  # noqa: E402
from merger import build_analysis_table, default_uart_interval, merge  # noqa: E402
from uart_reader import read_uart_csv  # noqa: E402


def make(uart_path: str, logger_path: str, out: str) -> None:
    u = read_uart_csv(uart_path)
    lg = read_logger_csv(logger_path)
    cfg = cfgmod.default_config()
    items, _ = cfgmod.build_items(cfg, u, lg)
    graphs = cfgmod.build_graphs(cfg, items)
    r = merge(u, lg, default_uart_interval(u.timestamps), lg.interval_ms)
    table = build_analysis_table(r, u, lg, items, cfgmod.elapsed_label(cfg))
    write_workbook(out, table, u, lg, graphs)


if __name__ == "__main__":
    args = sys.argv[1:] + [None] * 3
    make(args[0] or SAMPLE_UART, args[1] or SAMPLE_LOGGER, args[2] or EXPECTED_XLSX)
    print("written:", args[2] or EXPECTED_XLSX)
