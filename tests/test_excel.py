import openpyxl
import pytest

import config as cfgmod
from excel_writer import invalid_filename_chars, normalize_filename, sanitize_sheet_name, validate_graphs, write_workbook
from helpers import EXPECTED_XLSX, SAMPLE_LOGGER, SAMPLE_UART, write_logger_csv, write_uart_csv
from logger_reader import read_logger_csv
from make_expected import make
from merger import build_analysis_table, merge
from models import SOURCE_LOGGER, SOURCE_UART, GraphSetting, ItemSetting
from uart_reader import read_uart_csv


def _values(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def test_matches_expected_output(tmp_path):
    out = str(tmp_path / "out.xlsx")
    make(SAMPLE_UART, SAMPLE_LOGGER, out)
    got = openpyxl.load_workbook(out)
    exp = openpyxl.load_workbook(EXPECTED_XLSX)
    assert got.sheetnames == exp.sheetnames == ["解析", "グラフ", "log40", "260924-090155"]
    for name in ("解析", "log40", "260924-090155"):
        assert _values(got[name]) == _values(exp[name]), name
    assert len(got["グラフ"]._charts) == len(exp["グラフ"]._charts) == 1


def _setup(tmp_path, gaps=False):
    times = [i * 1000 for i in range(10)]
    if gaps:
        del times[3:5]
    u = read_uart_csv(write_uart_csv(str(tmp_path / "u.csv"), times))
    lg = read_logger_csv(write_logger_csv(str(tmp_path / "l.CSV"), 60, overrange_rows={10}))
    r = merge(u, lg, 1000, 200)
    return u, lg, r


def test_analysis_sheet_follows_settings(tmp_path):
    u, lg, r = _setup(tmp_path, gaps=True)
    items = [
        ItemSetting("t_ms", SOURCE_UART, False, "t_ms"),
        ItemSetting("voltage_mV", SOURCE_UART, True, "電圧(V)", 0.001, 0),
        ItemSetting("current_mA", SOURCE_UART, True, "電流(mA)", 1, 5),
        ItemSetting("temp_C", SOURCE_UART, False, "温度"),
        ItemSetting("CH1", SOURCE_LOGGER, False, "CH1(mV)", unit="mV"),
        ItemSetting("CH2", SOURCE_LOGGER, True, "CH2(mV)", 1000, 0, unit="V"),
    ]
    t = build_analysis_table(r, u, lg, items, "時間[s]")
    out = str(tmp_path / "o.xlsx")
    write_workbook(out, t, u, lg, [])
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["解析", "u", "l"]  # グラフ0個ならグラフシートなし
    ws = wb["解析"]
    assert ws["A1"].value == "測定日時"
    assert ws["B1"].value.strftime("%Y/%m/%d %H:%M:%S") == "2026/09/24 09:09:54"
    assert ws["B1"].number_format == "yyyy/mm/dd hh:mm:ss"
    assert (ws["A2"].value, ws["B2"].value) == ("測定機", "GL240")
    assert ws["A5"].value == "マイコン内部データ(UART)"
    assert [c.value for c in ws[6]] == ["時間[s]", "電圧(V)", "電流(mA)", None, "CH2(mV)"]
    assert ws["D5"].value is None and ws["E5"].value == "測定値"
    # 1行目
    assert ws["A7"].value == 0
    assert ws["B7"].value == pytest.approx(3.7)
    assert ws["C7"].value == -95
    assert ws["E7"].value == pytest.approx(2000)
    # 3行目（t=2s）→ ロガー行インデックス10 はレンジオーバー → 空欄
    assert ws["A9"].value == 2
    assert ws["E9"].value is None
    # 欠落行（4, 5行目）は UART 側が空欄で灰色、ロガーは値あり
    for row in (10, 11):
        assert ws[f"B{row}"].value is None and ws[f"C{row}"].value is None
        assert ws[f"E{row}"].value is not None
        assert ws[f"A{row}"].fill.fgColor.rgb.endswith("D9D9D9")
        assert ws[f"E{row}"].fill.fgColor.rgb.endswith("D9D9D9")
    assert ws["A12"].value == 5 and ws["B12"].value == pytest.approx(3.703)
    assert ws["A12"].fill.fill_type is None
    assert isinstance(ws["B7"].value, float)


def test_graph_sheet(tmp_path):
    u, lg, r = _setup(tmp_path)
    cfg = cfgmod.default_config()
    items, _ = cfgmod.build_items(cfg, u, lg)
    t = build_analysis_table(r, u, lg, items, "経過時間(s)")
    graphs = [GraphSetting("voltage_mV", "current_mA"), GraphSetting("CH1", None)]
    out = str(tmp_path / "o.xlsx")
    write_workbook(out, t, u, lg, graphs)

    import zipfile

    with zipfile.ZipFile(out) as z:
        # lxml の有無で "<a />" / "<a/>" の書き方が変わるのでそろえる
        c1 = z.read("xl/charts/chart1.xml").decode().replace(" />", "/>")
        c2 = z.read("xl/charts/chart2.xml").decode().replace(" />", "/>")
    last = 6 + len(r.rows)
    # 横軸 = 解析シートの経過時間(s)列
    assert f"'解析'!$A$7:$A${last}" in c1
    # 第1軸 電圧 = C列、第2軸 電流 = D列（elapsed_ms が B列）
    assert f"'解析'!$C$7:$C${last}" in c1 and f"'解析'!$D$7:$D${last}" in c1
    assert c1.count("<scatterChart>") == 2  # 第2軸あり
    assert '<crosses val="max"/>' in c1
    assert "電圧(mV) / 電流(mA)" in c1
    assert '<dispBlanksAs val="gap"/>' in c1
    assert '<symbol val="none"/>' in c1
    # グラフ2：第2軸なし、CH1 は UART の右に1列空けた位置（A + 6項目 + 空列 → I列）
    assert c2.count("<scatterChart>") == 1
    assert f"'解析'!$I$7:$I${last}" in c2
    assert "経過時間(s)" in c2
    wb = openpyxl.load_workbook(out)
    assert len(wb["グラフ"]._charts) == 2


def test_validate_graphs():
    items = [ItemSetting("a", SOURCE_UART, True, "A"), ItemSetting("b", SOURCE_UART, False, "B")]
    assert validate_graphs([GraphSetting("a", None)], items) == []
    assert validate_graphs([GraphSetting(None, None)], items)
    assert validate_graphs([GraphSetting("b", None)], items)
    assert validate_graphs([GraphSetting("a", "a")], items)
    assert validate_graphs([GraphSetting("a", "")], items)


def test_sheet_and_file_names():
    used = {"解析", "グラフ"}
    assert sanitize_sheet_name("a/b:c*d?[e]", used) == "a_b_c_d__e_"
    assert len(sanitize_sheet_name("x" * 40, used)) == 31
    assert sanitize_sheet_name("解析", used) == "解析(2)"
    assert invalid_filename_chars('a<b>.xlsx') == ["<", ">"]
    assert normalize_filename("abc") == "abc.xlsx"
    assert normalize_filename("abc.XLSX") == "abc.XLSX"
