"""テスト用の CSV 生成ヘルパー。"""

from __future__ import annotations

import os
from datetime import datetime, timedelta

UART_HEADER = "pc_timestamp,t_ms,elapsed_ms,voltage_mV,current_mA,cap_mAh,cap_max_mAh,soc_percent,soh_percent,temp_C"
START = datetime(2026, 9, 24, 9, 9, 54, 80339)


def write_uart_csv(path: str, times_ms: list[float], start: datetime = START, extra_lines: dict | None = None) -> str:
    """times_ms の各時刻（1行目からの経過ms）に1行ずつ書く。

    各行の値は行インデックス i から決める（voltage_mV = 3700 + i など）ので、
    どの行が採用されたかを値から確認できる。
    extra_lines: {挿入位置: 行文字列} 不正行などを差し込む。
    """
    lines = [UART_HEADER]
    for i, t in enumerate(times_ms):
        if extra_lines and i in extra_lines:
            lines.append(extra_lines[i])
        ts = start + timedelta(milliseconds=t)
        lines.append(
            f"{ts:%Y-%m-%d %H:%M:%S.%f},{35000 + i},{i * 1000},{3700 + i},{-100 - i},{212},{1261},{17},{94},{23.5}"
        )
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")
    return path


def write_logger_csv(
    path: str,
    n_rows: int,
    interval: str = "200ms",
    channels: list[tuple[str, str]] | None = None,
    model: str = "GL240",
    overrange_rows: set[int] | None = None,
) -> str:
    """GL240 形式のロガー CSV を作る。CHk の値は 行インデックス i × 0.01 + k。"""
    channels = channels or [("CH1", "mV"), ("CH2", "V")]
    ch_names = [c for c, _ in channels]
    lines = [
        'ベンダ,"GRAPHTEC Corporation"',
        f'モデル,"{model}"',
        'バージョン,"Ver1.11  "',
        "最大CH数,10",
        f"測定間隔,{interval}",
        f"測定点数,{n_rows}        ",
        "開始時刻,2026-09-24,09:09:54",
        "アンプ設定",
        "CH,信号名,アンプ,入力,レンジ,温度レンジ,フィルタ,スパン,,単位,色,,,線幅",
    ]
    for k, (ch, unit) in enumerate(channels, start=1):
        lines.append(f'{ch}," CH {k}","M",DC,100mV,,Off,+50.00,-50.00,[{unit}],28,5,6,0')
    lines += [
        "演算設定",
        "CH,演算,演算子,右辺,スケーリング,スパン,,単位",
        "測定値",
        "番号,日付/時間,ms," + ",".join(ch_names) + ",Alarm1-10,AlarmOut",
        "NO.,Time,ms," + ",".join(f'"{u}"' for _, u in channels) + ',"A1234567890","AO1234"',
    ]
    for i in range(n_rows):
        vals = []
        for k in range(1, len(channels) + 1):
            if overrange_rows and i in overrange_rows:
                vals.append("+++++++")
            else:
                vals.append(f"+{i * 0.01 + k:8.2f}")
        lines.append(f"{i + 1:10d},2026/09/24 09:09:54,  0," + ",".join(vals) + ",LLLLLLLLLL,LLLL")
    with open(path, "w", encoding="cp932", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")
    return path


DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
SAMPLE_UART = os.path.join(DATA_DIR, "log40.csv")
SAMPLE_LOGGER = os.path.join(DATA_DIR, "260924-090155.CSV")
EXPECTED_XLSX = os.path.join(DATA_DIR, "expected_log40.xlsx")
