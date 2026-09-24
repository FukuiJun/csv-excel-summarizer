import pytest

from helpers import SAMPLE_LOGGER, SAMPLE_UART, write_logger_csv, write_uart_csv
from logger_reader import LoggerFormatError, parse_interval_ms, parse_number, read_logger_csv
from merger import default_uart_interval
from uart_reader import UartFormatError, read_uart_csv


def test_logger_header_sample():
    lg = read_logger_csv(SAMPLE_LOGGER)
    assert lg.model == "GL240"
    assert lg.interval_ms == 200
    assert [(c.name, c.unit) for c in lg.channels] == [("CH1", "mV"), ("CH2", "V")]
    assert lg.row_count == 36
    assert lg.numbers[0] == 1 and lg.numbers[-1] == 36
    assert lg.values["CH1"][:3] == [0.28, 0.28, 0.29]
    assert lg.values["CH2"][0] == 0.0003
    assert "Alarm1-10" not in lg.values


@pytest.mark.parametrize(
    "text,ms",
    [("200ms", 200), ("1s", 1000), ("1min", 60000), ("500 ms", 500), ("2s", 2000), ("1h", 3600000)],
)
def test_parse_interval(text, ms):
    assert parse_interval_ms(text) == ms


def test_parse_number():
    assert parse_number("+  0.28") == 0.28
    assert parse_number("-  1.50") == -1.5
    assert parse_number("+++++++") is None
    assert parse_number("") is None


def test_logger_dynamic_channels_and_overrange(tmp_path):
    chs = [(f"CH{k}", u) for k, u in zip(range(1, 11), ["mV", "V", "V", "mV", "V", "C", "V", "V", "V", "V"])]
    p = write_logger_csv(str(tmp_path / "lg.CSV"), 5, interval="1s", channels=chs, overrange_rows={2})
    lg = read_logger_csv(p)
    assert lg.interval_ms == 1000
    assert len(lg.channels) == 10
    assert lg.unit_of("CH6") == "C"
    assert lg.values["CH1"][2] is None
    assert lg.values["CH10"][1] == pytest.approx(10.01)


def test_logger_wrong_file(tmp_path):
    p = write_uart_csv(str(tmp_path / "u.csv"), [0, 1000])
    with pytest.raises(LoggerFormatError):
        read_logger_csv(p)


def test_uart_sample():
    u = read_uart_csv(SAMPLE_UART)
    assert u.row_count == 8
    assert u.skipped == 0
    assert u.columns[0] == "t_ms" and "pc_timestamp" not in u.columns
    assert u.values["temp_C"][6] == 23.9


def test_uart_skips_bad_rows(tmp_path):
    bad = {1: "2026-09-24 09:09:55.000000,1,2,3", 2: "2026-09-24 09:09:55.000000,1,2,abc,4,5,6,7,8,9", 3: "xx,1,2,3,4,5,6,7,8,9"}
    p = write_uart_csv(str(tmp_path / "u.csv"), [0, 1000, 2000, 3000], extra_lines=bad)
    u = read_uart_csv(p)
    assert u.row_count == 4
    assert u.skipped == 3


def test_uart_wrong_file():
    with pytest.raises(UartFormatError):
        read_uart_csv(SAMPLE_LOGGER)


def test_default_uart_interval():
    u = read_uart_csv(SAMPLE_UART)
    assert default_uart_interval(u.timestamps) == 1000


def test_default_uart_interval_jitter(tmp_path):
    times = [i * 100 + (7 if i % 2 else -6) for i in range(50)]
    times[0] = 0
    u = read_uart_csv(write_uart_csv(str(tmp_path / "u.csv"), times))
    assert default_uart_interval(u.timestamps) == 100
