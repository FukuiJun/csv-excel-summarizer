"""GUI の簡易テスト。tkinter とディスプレイがない環境ではスキップする。"""

import os

import pytest

tk = pytest.importorskip("tkinter")

from helpers import write_logger_csv, write_uart_csv  # noqa: E402


@pytest.fixture
def app(tmp_path):
    try:
        import gui

        a = gui.App(config_path=str(tmp_path / "config.json"))
    except tk.TclError:
        pytest.skip("ディスプレイがありません")
    a.withdraw()
    times = [i * 1000 for i in range(10)]
    del times[3:5]
    a.uart_path.set(write_uart_csv(str(tmp_path / "u.csv"), times))
    a.logger_path.set(write_logger_csv(str(tmp_path / "l.CSV"), 60))
    yield a
    a.destroy()


def _load(app):
    app.select_frame.load()
    app.update()
    return app.adjust_frame


def test_load_button_disabled_until_both_selected(app):
    app.logger_path.set("")
    app.update()
    assert app.select_frame.load_btn.instate(["disabled"])
    app.logger_path.set("x.CSV")
    assert not app.select_frame.load_btn.instate(["disabled"])


def test_screen2_flow(app, tmp_path):
    f = _load(app)
    assert f is not None
    assert f.uart_int.get() == "1000" and f.logger_int.get() == "200"
    assert "欠落：1箇所（2行" in f.gap_label.cget("text")
    assert "出力予定 10行" in f.mode_label.cget("text")
    assert f.validate() == []
    assert f.filename.get() == "解析_260924-090954.xlsx"
    assert f.folder.get() == str(tmp_path)

    # 間隔の入力エラーで出力ボタンが無効になる
    f.uart_int.set("0")
    f.recalc()
    assert f.export_btn.instate(["disabled"])
    f.uart_int.set("500")
    f.recalc()
    assert "ロガー基準" not in f.mode_label.cget("text")
    f.uart_int.set("1000")
    f.recalc()

    rows = {r.key: r for r in f.item_rows}
    # グラフで選択中の項目の採用を外すと未選択（エラー）になる
    rows["current_mA"].enabled.set(False)
    app.update()
    assert f.graph_rows[0].secondary == ""
    assert f.export_btn.instate(["disabled"])
    assert str(rows["current_mA"].label_e.cget("state")) == "disabled"
    rows["current_mA"].enabled.set(True)
    f.graph_rows[0].secondary = None
    f.on_items_changed()
    assert f.validate() == []

    # ラベルの変更はグラフの選択肢に即時反映
    rows["voltage_mV"].label.set("電圧(V)")
    app.update()
    assert f.graph_rows[0].p_cb.get() == "電圧(V)"

    # ラベル重複・係数の数値エラー・ファイル名エラー
    rows["cap_mAh"].label.set("電圧(V)")
    assert any("重複" in e for e in f.validate())
    rows["cap_mAh"].label.set("CAP")
    rows["CH1"].coef.set("abc")
    assert any("係数" in e for e in f.validate())
    rows["CH1"].coef.set("0.5")
    f.filename.set("a?b")
    assert any("ファイル名" in e for e in f.validate())
    f.filename.set("out")
    assert f.validate() == []

    # 初期値に戻す
    import gui

    orig = gui.messagebox.askokcancel
    gui.messagebox.askokcancel = lambda *a, **k: True
    try:
        f.reset_defaults()
    finally:
        gui.messagebox.askokcancel = orig
    rows = {r.key: r for r in f.item_rows}
    assert rows["voltage_mV"].label.get() == "電圧(mV)"
    assert rows["CH1"].coef.get() == "1"
    assert len(f.graph_rows) == 1

    # 戻る → ファイルパスは保持
    path = app.uart_path.get()
    app.show_select()
    assert app.adjust_frame is None and app.uart_path.get() == path
    assert os.path.exists(path)
