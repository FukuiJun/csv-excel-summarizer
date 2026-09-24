import random

import pytest

from helpers import write_logger_csv, write_uart_csv
from logger_reader import read_logger_csv
from merger import (
    ROWS_ERROR,
    ROWS_OK,
    ROWS_WARN,
    build_analysis_table,
    build_uart_timeline,
    check_row_limit,
    max_sheet_rows,
    merge,
)
from models import SOURCE_LOGGER, SOURCE_UART, ItemSetting
from uart_reader import read_uart_csv


def load(tmp_path, uart_times, logger_rows, logger_interval="200ms"):
    u = read_uart_csv(write_uart_csv(str(tmp_path / "u.csv"), uart_times))
    lg = read_logger_csv(write_logger_csv(str(tmp_path / "l.CSV"), logger_rows, interval=logger_interval))
    return u, lg


def test_uart1000_logger200(tmp_path):
    u, lg = load(tmp_path, [i * 1000 for i in range(10)], 50)
    r = merge(u, lg, 1000, 200)
    assert r.uart_based
    assert [x.logger_idx for x in r.rows] == [0, 5, 10, 15, 20, 25, 30, 35, 40, 45]
    assert [x.uart_idx for x in r.rows] == list(range(10))
    assert [x.t_ms for x in r.rows] == [i * 1000 for i in range(10)]


def test_uart100_logger200(tmp_path):
    u, lg = load(tmp_path, [i * 100 for i in range(20)], 10)
    r = merge(u, lg, 100, 200)
    assert not r.uart_based
    assert [x.logger_idx for x in r.rows] == list(range(10))
    assert [x.uart_idx for x in r.rows] == [0, 2, 4, 6, 8, 10, 12, 14, 16, 18]


def test_uart100_logger200_no_near_uart_row_is_blank(tmp_path):
    # UART の時刻が 60ms ずれていて、許容 50ms を超える行は UART 側を空欄にする
    times = [i * 100 for i in range(20)]
    times[4] = 460  # ロガー t=400 に最も近いのは 460（差60ms > 50ms）… 300 との差は100
    u, lg = load(tmp_path, times, 10)
    r = merge(u, lg, 100, 200)
    row = r.rows[2]  # ロガー t=400
    assert row.uart_idx is None and row.logger_idx == 2
    assert row.t_ms == 400  # 対応行がない場合は tL


def test_uart1000_logger300(tmp_path):
    u, lg = load(tmp_path, [i * 1000 for i in range(10)], 40, "300ms")
    r = merge(u, lg, 1000, 300)
    assert [x.logger_idx for x in r.rows] == [0, 3, 7, 10, 13, 17, 20, 23, 27, 30]


def test_equal_interval(tmp_path):
    u, lg = load(tmp_path, [i * 200 for i in range(10)], 10)
    r = merge(u, lg, 200, 200)
    assert r.uart_based
    assert [x.logger_idx for x in r.rows] == list(range(10))


def test_jitter_1000_200_every_5(tmp_path):
    rnd = random.Random(1)
    times = [0] + [i * 1000 + rnd.uniform(-10, 10) for i in range(1, 30)]
    u, lg = load(tmp_path, times, 200)
    r = merge(u, lg, 1000, 200)
    assert [x.logger_idx for x in r.rows] == [i * 5 for i in range(30)]


def test_stops_when_logger_runs_out(tmp_path):
    u, lg = load(tmp_path, [i * 1000 for i in range(10)], 21)  # ロガーは 0〜4000ms 分
    r = merge(u, lg, 1000, 200)
    assert len(r.rows) == 5


def test_stops_when_uart_runs_out(tmp_path):
    u, lg = load(tmp_path, [i * 100 for i in range(10)], 50)  # UART は 0〜900ms
    r = merge(u, lg, 100, 200)
    assert len(r.rows) == 5


def test_gap_fill(tmp_path):
    times = [i * 1000 for i in range(10)]
    del times[3:5]  # 3000ms, 4000ms の行を欠落させる
    u, lg = load(tmp_path, times, 60)
    tl = build_uart_timeline(u, 1000)
    assert len(tl.gaps) == 1
    g = tl.gaps[0]
    assert g.missing == 2
    assert g.prev_line == 4 and g.next_line == 5  # CSV 上の行番号（ヘッダ=1）
    assert tl.t_ms == [i * 1000 for i in range(10)]
    assert tl.src == [0, 1, 2, None, None, 3, 4, 5, 6, 7]

    r = merge(u, lg, 1000, 200, timeline=tl)
    # 欠落行があってもロガーとの対応はずれない
    assert [x.logger_idx for x in r.rows] == [i * 5 for i in range(10)]
    assert [x.is_gap for x in r.rows] == [False] * 3 + [True] * 2 + [False] * 5

    items = [
        ItemSetting("voltage_mV", SOURCE_UART, True, "V"),
        ItemSetting("CH1", SOURCE_LOGGER, True, "CH1"),
    ]
    t = build_analysis_table(r, u, lg, items, "経過時間(s)")
    assert t.rows[3] == (3.0, [None], [pytest.approx(1.15)], True)
    assert t.rows[5][1] == [3703]  # 欠落後の行は元の4行目以降


def test_gap_fill_with_jitter_and_duplicates(tmp_path):
    times = [0, 1005, 1010, 1998, 5003, 5997]  # 1010 は二重受信、1998→5003 で 2行欠落
    u, lg = load(tmp_path, times, 40)
    tl = build_uart_timeline(u, 1000)
    assert tl.duplicates == 1
    assert [g.missing for g in tl.gaps] == [2]
    assert tl.t_ms[3:5] == [pytest.approx(2998), pytest.approx(3998)]
    r = merge(u, lg, 1000, 200, timeline=tl)
    assert [x.logger_idx for x in r.rows] == [0, 5, 10, 15, 20, 25, 30]


def test_logger_based_gap_rows_marked(tmp_path):
    times = [i * 100 for i in range(20)]
    del times[5:9]  # 500〜800ms 欠落
    u, lg = load(tmp_path, times, 10)
    r = merge(u, lg, 100, 200)
    gap = [x for x in r.rows if x.is_gap]
    assert [x.logger_idx for x in gap] == [3, 4]  # t=600, 800
    assert all(x.uart_idx is None for x in gap)
    assert r.rows[2].uart_idx == 4 and not r.rows[2].is_gap  # t=400 は残っている
    assert r.rows[5].uart_idx == 6  # t=1000（900ms が欠落後の最初の行 idx5）


def test_row_limit():
    assert check_row_limit(100_000, 100_000, 1_000_000) == ROWS_OK
    assert check_row_limit(100_001, 100_000, 1_000_000) == ROWS_WARN
    assert check_row_limit(1_000_000, 100_000, 1_000_000) == ROWS_WARN
    assert check_row_limit(1_000_001, 100_000, 1_000_000) == ROWS_ERROR


def test_max_sheet_rows_uses_largest_sheet(tmp_path):
    u, lg = load(tmp_path, [i * 1000 for i in range(10)], 50)
    r = merge(u, lg, 1000, 200)
    # ロガーの元データシート（ヘッダ部 + データ50行）が最大
    assert max_sheet_rows(len(r.rows), u, lg) == len(lg.raw_rows)
    assert len(lg.raw_rows) > 6 + len(r.rows)


def test_elapsed_ms_estimated_on_rows_without_uart(tmp_path):
    times = [i * 1000 for i in range(10)]
    del times[3:5]
    u, lg = load(tmp_path, times, 60)
    r = merge(u, lg, 1000, 200)
    items = [
        ItemSetting("elapsed_ms", SOURCE_UART, True, "time(ms)"),
        ItemSetting("voltage_mV", SOURCE_UART, True, "V"),
    ]
    t = build_analysis_table(r, u, lg, items, "経過時間(s)")
    # helpers の elapsed_ms は 行インデックス×1000。欠落行（3,4）は直前(2000)＋間隔で推定、他の UART 値は空欄
    assert [row[1][0] for row in t.rows] == [0, 1000, 2000, 3000, 4000, 3000, 4000, 5000, 6000, 7000]
    assert t.rows[3][1][1] is None
    # 係数・オフセットも推定値に適用される
    items[0].coef, items[0].offset = 0.001, 1
    t = build_analysis_table(r, u, lg, items, "経過時間(s)")
    assert t.rows[4][1][0] == pytest.approx(5.0)


def test_elapsed_ms_estimated_logger_based(tmp_path):
    times = [i * 100 for i in range(20)]
    del times[5:9]
    u, lg = load(tmp_path, times, 10)
    r = merge(u, lg, 100, 200)
    items = [ItemSetting("elapsed_ms", SOURCE_UART, True, "time(ms)")]
    t = build_analysis_table(r, u, lg, items, "経過時間(s)")
    # t=600,800 は欠落：直前の実行（t=400, elapsed_ms=4000）＋差分
    assert t.rows[3][1][0] == pytest.approx(4200) and t.rows[4][1][0] == pytest.approx(4400)
