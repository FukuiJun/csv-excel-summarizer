"""Excel 出力と、出力成功時の設定保存をまとめた処理（仕様 2.2・5.2.6）。GUI に依存しない。"""

from __future__ import annotations

from dataclasses import dataclass

import config as cfgmod
from excel_writer import ProgressFn, write_workbook
from logger_reader import LoggerData
from merger import AnalysisTable
from models import GraphSetting, ItemSetting
from uart_reader import UartData


@dataclass
class ExportResult:
    config: dict  # 保存後（保存に失敗した場合も反映済み）の設定
    config_error: str | None  # config.json に保存できなかった場合の理由


def export(
    path: str,
    table: AnalysisTable,
    uart: UartData,
    logger: LoggerData,
    items: list[ItemSetting],
    graphs: list[GraphSetting],
    elapsed_label: str,
    cfg: dict,
    config_path: str | None,
    uart_dir: str | None = None,
    logger_dir: str | None = None,
    progress: ProgressFn | None = None,
) -> ExportResult:
    """Excel を出力し、成功したら設定を config.json に保存する。

    出力に失敗した場合は例外をそのまま送出し、設定は保存しない。
    設定の保存に失敗しても出力は成功扱いとし、理由を config_error に入れて返す。
    """
    write_workbook(path, table, uart, logger, graphs, progress=progress)
    new_cfg = cfgmod.apply_settings(cfg, items, graphs, elapsed_label, uart_dir=uart_dir, logger_dir=logger_dir)
    try:
        cfgmod.save_config(new_cfg, config_path)
        err = None
    except OSError as e:
        err = str(e)
    return ExportResult(new_cfg, err)
