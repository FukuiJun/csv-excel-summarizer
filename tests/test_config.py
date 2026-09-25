import json
import os

import pytest

import config as cfgmod
import exporter
from helpers import write_logger_csv, write_uart_csv
from logger_reader import read_logger_csv
from merger import build_analysis_table, merge
from models import ELAPSED_KEY, GraphSetting
from uart_reader import read_uart_csv


def _data(tmp_path, channels=None, name="l.CSV"):
    u = read_uart_csv(write_uart_csv(str(tmp_path / "u.csv"), [i * 1000 for i in range(5)]))
    lg = read_logger_csv(write_logger_csv(str(tmp_path / name), 30, channels=channels))
    return u, lg


def _export(tmp_path, cfg, cfg_path, u, lg, items, graphs, out_name="o.xlsx", elapsed="経過時間(s)"):
    r = merge(u, lg, 1000, 200)
    t = build_analysis_table(r, u, lg, items, elapsed)
    return exporter.export(
        str(tmp_path / out_name), t, u, lg, items, graphs, elapsed, cfg, cfg_path,
        uart_dir=str(tmp_path), logger_dir=str(tmp_path),
    )


def test_missing_config_uses_defaults_and_recreates(tmp_path):
    p = str(tmp_path / "config.json")
    res = cfgmod.load_config(p)
    assert res.status == "missing"
    assert res.config == cfgmod.default_config()
    with open(p, encoding="utf-8") as f:
        assert json.load(f) == cfgmod.default_config()


@pytest.mark.parametrize("content", ["{broken", "[1, 2]", ""])
def test_broken_config_uses_defaults_and_recreates(tmp_path, content):
    p = tmp_path / "config.json"
    p.write_text(content, encoding="utf-8")
    res = cfgmod.load_config(str(p))
    assert res.status == "broken"
    assert res.config == cfgmod.default_config()
    assert json.loads(p.read_text(encoding="utf-8")) == cfgmod.default_config()


def test_invalid_fields_fall_back(tmp_path):
    p = tmp_path / "config.json"
    p.write_text(json.dumps({"row_limit": "abc", "graphs": "x", "elapsed_label": "t(s)"}), encoding="utf-8")
    cfg = cfgmod.load_config(str(p)).config
    assert cfg["row_limit"] == 1_000_000
    assert cfg["graphs"] == cfgmod.DEFAULT_GRAPHS
    assert cfg["elapsed_label"] == "t(s)"


def test_defaults_match_spec(tmp_path):
    u, lg = _data(tmp_path)
    items, warnings = cfgmod.build_items(cfgmod.default_config(), u, lg)
    assert warnings == []
    d = {it.key: (it.enabled, it.label, it.coef, it.offset) for it in items}
    assert d["t_ms"] == (False, "t_ms", 1, 0)
    assert d["elapsed_ms"] == (True, "time(ms)", 1, 0)
    assert d["soc_percent"] == (True, "SOC(％)", 1, 0)
    assert d["soh_percent"][0] is False
    assert d["CH1"] == (True, "CH1(mV)", 1, 0)
    assert d["CH2"] == (True, "CH2(V)", 1, 0)
    graphs = cfgmod.build_graphs(cfgmod.default_config(), items)
    assert graphs == [GraphSetting("voltage_mV", "current_mA")]


def test_save_on_success_and_restore(tmp_path):
    cfg_path = str(tmp_path / "config.json")
    cfg = cfgmod.load_config(cfg_path).config
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg)
    by = {it.key: it for it in items}
    by["voltage_mV"].label = "電圧(V)"
    by["voltage_mV"].coef = 0.001
    by["t_ms"].enabled = True
    by["CH1"].coef = 1000
    graphs = [GraphSetting("voltage_mV", "CH1"), GraphSetting("soc_percent", None)]
    res = _export(tmp_path, cfg, cfg_path, u, lg, items, graphs, elapsed="秒")
    assert res.config_error is None

    cfg2 = cfgmod.load_config(cfg_path).config
    items2, _ = cfgmod.build_items(cfg2, u, lg)
    by2 = {it.key: it for it in items2}
    assert (by2["voltage_mV"].label, by2["voltage_mV"].coef) == ("電圧(V)", 0.001)
    assert by2["t_ms"].enabled is True
    assert by2["CH1"].coef == 1000
    # 項目ごとのオフセットは廃止（保存しない）
    assert "offset" not in json.load(open(cfg_path, encoding="utf-8"))["logger_channels"]["CH1"]
    assert cfgmod.elapsed_label(cfg2) == "秒"
    # ラベルを変更してもグラフ設定は元の列（キー）に対応したまま
    assert cfgmod.build_graphs(cfg2, items2) == graphs
    assert cfg2["graphs"][0] == {"x": "elapsed_ms", "x_unit": None, "primary": ["voltage_mV", None], "secondary": ["CH1", None]}
    assert cfg2["last_dirs"] == {"uart": str(tmp_path), "logger": str(tmp_path)}


def test_not_saved_on_failure(tmp_path):
    cfg_path = str(tmp_path / "config.json")
    cfg = cfgmod.load_config(cfg_path).config
    before = open(cfg_path, encoding="utf-8").read()
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg)
    items[0].label = "changed"
    with pytest.raises(OSError):
        _export(tmp_path, cfg, cfg_path, u, lg, items, [], out_name=os.path.join("no_such_dir", "o.xlsx"))
    assert open(cfg_path, encoding="utf-8").read() == before


def test_config_write_failure_is_reported_not_fatal(tmp_path):
    cfg = cfgmod.default_config()
    bad_path = str(tmp_path / "no_such_dir" / "config.json")
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg)
    res = _export(tmp_path, cfg, bad_path, u, lg, items, [])
    assert os.path.exists(tmp_path / "o.xlsx")
    assert res.config_error


def test_intervals_folder_filename_not_saved(tmp_path):
    cfg_path = str(tmp_path / "config.json")
    cfg = cfgmod.load_config(cfg_path).config
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg)
    _export(tmp_path, cfg, cfg_path, u, lg, items, [])
    saved = json.load(open(cfg_path, encoding="utf-8"))
    # 間隔・出力先フォルダ・ファイル名の項目は保存されない（キーは既定と同じ）
    assert set(saved) == set(cfgmod.DEFAULT_CONFIG)
    assert saved["output_filename_format"] == cfgmod.DEFAULT_CONFIG["output_filename_format"]
    # ファイル名は毎回 UART 1行目の時刻から生成される
    assert cfgmod.make_output_filename(saved, u.timestamps[0]) == "解析_260924-090954.xlsx"


def test_missing_channels_are_kept(tmp_path):
    cfg_path = str(tmp_path / "config.json")
    cfg = cfgmod.load_config(cfg_path).config
    chs3 = [("CH1", "mV"), ("CH2", "V"), ("CH3", "V")]
    u, lg3 = _data(tmp_path, chs3, "l3.CSV")
    items, _ = cfgmod.build_items(cfg, u, lg3)
    {it.key: it for it in items}["CH3"].label = "温度センサ"
    res = _export(tmp_path, cfg, cfg_path, u, lg3, items, [])

    _, lg1 = _data(tmp_path, [("CH1", "mV")], "l1.CSV")
    items1, _ = cfgmod.build_items(res.config, u, lg1)
    assert [it.key for it in items1 if it.source == "logger"] == ["CH1"]  # 画面には出さない
    _export(tmp_path, res.config, cfg_path, u, lg1, items1, [], out_name="o2.xlsx")
    saved = json.load(open(cfg_path, encoding="utf-8"))
    assert saved["logger_channels"]["CH3"]["label"] == "温度センサ"


def test_graph_item_missing_or_disabled_becomes_unselected(tmp_path):
    u, lg = _data(tmp_path, [("CH1", "mV")])
    cfg = cfgmod.default_config()
    cfg["graphs"] = [{"primary": "CH5", "secondary": "voltage_mV"}, {"primary": "voltage_mV", "secondary": "t_ms"}]
    items, _ = cfgmod.build_items(cfg, u, lg)  # t_ms は既定で採用なし
    graphs = cfgmod.build_graphs(cfg, items)
    assert graphs == [GraphSetting(None, "voltage_mV"), GraphSetting("voltage_mV", "")]


def test_unit_change_warning(tmp_path):
    cfg = cfgmod.default_config()
    cfg["logger_channels"] = {"CH1": {"enabled": True, "label": "CH1(mV)", "coef": 1, "offset": 0, "unit": "mV"}}
    u, lg = _data(tmp_path, [("CH1", "V")])
    _, warnings = cfgmod.build_items(cfg, u, lg)
    assert len(warnings) == 1 and "CH1" in warnings[0]
    u, lg = _data(tmp_path, [("CH1", "mV")], "same.CSV")
    assert cfgmod.build_items(cfg, u, lg)[1] == []


def test_reset_to_defaults(tmp_path):
    cfg = cfgmod.default_config()
    cfg["uart_columns"]["voltage_mV"] = {"enabled": False, "label": "X", "coef": 2, "offset": 3}
    cfg["logger_channels"] = {"CH1": {"enabled": False, "label": "Y", "coef": 2, "offset": 3, "unit": "mV"}}
    cfg["graphs"] = []
    cfg["elapsed_label"] = "t"
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg, use_saved=False)
    by = {it.key: it for it in items}
    assert (by["voltage_mV"].enabled, by["voltage_mV"].label, by["voltage_mV"].coef) == (True, "電圧(mV)", 1)
    assert (by["CH1"].enabled, by["CH1"].label) == (True, "CH1(mV)")
    assert cfgmod.build_graphs(cfg, items, use_saved=False) == [GraphSetting("voltage_mV", "current_mA")]
    assert cfgmod.elapsed_label(cfg, use_saved=False) == "経過時間(s)"


def test_graph_x_axis_default_and_restore(tmp_path):
    u, lg = _data(tmp_path)
    cfg = cfgmod.default_config()
    assert cfg["graphs"][0]["x"] == "elapsed_ms"
    items, _ = cfgmod.build_items(cfg, u, lg)
    assert cfgmod.build_graphs(cfg, items)[0].x == "elapsed_ms"

    # 横軸の設定がない古い config.json は初期値（elapsed_ms）になる
    p = tmp_path / "config.json"
    p.write_text(json.dumps({"graphs": [{"primary": "voltage_mV", "secondary": None}]}), encoding="utf-8")
    old = cfgmod.load_config(str(p)).config
    assert cfgmod.build_graphs(old, items) == [GraphSetting("voltage_mV", None, "elapsed_ms")]

    # 経過時間(s) を選んだ設定は保存・復元される
    new = cfgmod.apply_settings(cfg, items, [GraphSetting("voltage_mV", None, ELAPSED_KEY)], "経過時間(s)")
    assert cfgmod.build_graphs(new, items) == [GraphSetting("voltage_mV", None, ELAPSED_KEY)]

    # 横軸の項目の採用を外すと未選択
    {it.key: it for it in items}["elapsed_ms"].enabled = False
    assert cfgmod.build_graphs(cfg, items)[0].x is None


def test_old_item_offset_is_ignored(tmp_path):
    cfg = cfgmod.default_config()
    cfg["uart_columns"]["voltage_mV"] = {"enabled": True, "label": "V", "coef": 1, "offset": 5}
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg)
    assert {it.key: it for it in items}["voltage_mV"].offset == 0


def test_item_colors_default_save_restore(tmp_path):
    from models import DEFAULT_COLORS

    cfg_path = str(tmp_path / "config.json")
    cfg = cfgmod.load_config(cfg_path).config
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg)
    # 既定色は項目の並び順に割り当て
    defaults = [it.color for it in items]
    assert defaults[:3] == DEFAULT_COLORS[:3]
    {it.key: it for it in items}["CH1"].color = "112233"
    _export(tmp_path, cfg, cfg_path, u, lg, items, [])

    cfg2 = cfgmod.load_config(cfg_path).config
    items2, _ = cfgmod.build_items(cfg2, u, lg)
    assert {it.key: it for it in items2}["CH1"].color == "112233"
    # 初期値に戻すと既定色
    items3, _ = cfgmod.build_items(cfg2, u, lg, use_saved=False)
    assert [it.color for it in items3] == defaults
    # 不正な色は既定色
    raw = json.loads(open(cfg_path, encoding="utf-8").read())
    raw["logger_channels"]["CH1"]["color"] = "zzz"
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(raw, f)
    items4, _ = cfgmod.build_items(cfgmod.load_config(cfg_path).config, u, lg)
    assert [it.color for it in items4] == defaults


def test_old_graph_format_is_read(tmp_path):
    p = tmp_path / "config.json"
    p.write_text(
        json.dumps({"graphs": [{"x": "elapsed_ms", "primary": "voltage_mV", "secondary": "current_mA"}]}),
        encoding="utf-8",
    )
    cfg = cfgmod.load_config(str(p)).config
    u, lg = _data(tmp_path)
    items, _ = cfgmod.build_items(cfg, u, lg)
    assert cfgmod.build_graphs(cfg, items) == [GraphSetting(["voltage_mV", None], ["current_mA", None])]


def test_graph_x_unit_save_restore(tmp_path):
    u, lg = _data(tmp_path)
    cfg = cfgmod.default_config()
    items, _ = cfgmod.build_items(cfg, u, lg)
    assert cfgmod.build_graphs(cfg, items)[0].x_unit is None  # 既定は元の単位
    new = cfgmod.apply_settings(cfg, items, [GraphSetting(["voltage_mV"], [None], "elapsed_ms", "min")], "経過時間(s)")
    assert cfgmod.build_graphs(new, items)[0].x_unit == "min"
    p = tmp_path / "config.json"
    p.write_text(json.dumps({"graphs": [{"x": "elapsed_ms", "x_unit": "day", "primary": ["voltage_mV"]}]}), encoding="utf-8")
    assert cfgmod.build_graphs(cfgmod.load_config(str(p)).config, items)[0].x_unit is None  # 不正な単位
