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
    assert "欠落 1箇所（2行" in f.gap_label_text
    assert "出力予定 10 行" in f.mode_label.cget("text")
    # 画面2のヘッダー：現在のステップと読み込んだファイル名
    assert "a" in app.header_note.cget("text") or ".csv" in app.header_note.cget("text")
    assert f.validate() == []
    assert f.filename.get() == "解析_260924-090954.xlsx"
    assert f.folder.get() == str(tmp_path)
    # グラフの横軸は初期値 elapsed_ms（ラベル time(ms)）、選択肢に経過時間(s)もある
    assert f.graph_rows[0].x == "elapsed_ms"
    assert f.graph_rows[0].x_cb.get() == "time(ms)"
    assert "経過時間(s)" in f.graph_rows[0].x_cb.cget("values")

    # 時間オフセット：ロガーを 1000ms 遅らせると、1行目はロガーなし・2行目にロガー1行目が対応する
    f.logger_off.set("1000")
    f.recalc()
    assert f.result.rows[0].logger_idx is None and f.result.rows[1].logger_idx == 0
    f.logger_off.set("abc")
    f.recalc()
    assert f.export_btn.instate(["disabled"])
    f.logger_off.set("0")
    f.recalc()
    assert f.validate() == []

    # 間隔の入力エラーで出力ボタンが無効になる
    f.uart_int.set("0")
    f.recalc()
    assert f.export_btn.instate(["disabled"])
    f.uart_int.set("500")
    f.recalc()
    assert "ロガー基準" not in f.mode_line1.cget("text")
    f.uart_int.set("1000")
    f.recalc()

    rows = {r.key: r for r in f.item_rows}
    # グラフで選択中の項目の採用を外すと未選択（エラー）になる
    rows["current_mA"].enabled.set(False)
    app.update()
    assert f.graph_rows[0].secondary == ["", None]
    assert f.graph_rows[0].cbs[3].instate(["invalid"])  # 第2軸の1つ目が赤枠
    assert f.export_btn.instate(["disabled"])
    assert str(rows["current_mA"].label_e.cget("state")) == "disabled"
    rows["current_mA"].enabled.set(True)
    f.graph_rows[0].values[3] = None  # 「なし」にする
    f.on_items_changed()
    assert f.validate() == []

    # ラベルの変更はグラフの選択肢に即時反映
    rows["voltage_mV"].label.set("電圧(V)")
    app.update()
    assert f.graph_rows[0].cbs[1].get() == "電圧(V)"

    # 第1軸に2つ目の項目を追加できる／同じ項目を2回選ぶとエラー
    g = f.graph_rows[0]
    g.cbs[2].set("CH1(mV)")
    g._on_select()
    assert g.primary == ["voltage_mV", "CH1"] and f.validate() == []
    g.cbs[2].set("電圧(V)")
    g._on_select()
    assert any("2回以上" in e for e in f.validate())
    g.cbs[2].set("なし")
    g._on_select()
    assert f.validate() == []

    # 色：パレットで選ぶと項目の色が変わる
    import gui

    r = rows["voltage_mV"]
    r.color_btn.invoke()  # 色見本ボタンを押す
    app.update()
    pal = r._palette
    assert isinstance(pal, gui.ColorPalette) and pal.winfo_exists() and pal.winfo_viewable()
    r.color_btn.invoke()  # 開いているときにもう一度押すと閉じる
    app.update()
    assert not pal.winfo_exists() and r._palette is None
    before = r.color
    for close in ("close_btn", "cancel_btn"):  # ［✕ 閉じる］［キャンセル］で閉じる（色は変わらない）
        r.color_btn.invoke()
        app.update()
        p = r._palette
        getattr(p, close).invoke()
        app.update()
        assert not p.winfo_exists() and r.color == before
    r.color_btn.invoke()
    app.update()
    pal = r._palette
    pal._pick("ff0000")  # パレットの色をクリック
    assert not pal.winfo_exists()
    assert r.to_setting().color == "FF0000"
    assert str(r.color_btn.cget("bg")).lower() == "#ff0000"
    # 採用を外した項目の色見本は押しても開かない
    rows["t_ms"].color_btn.invoke()
    assert rows["t_ms"]._palette is None

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


def test_layout_and_separate_scroll(app):
    app.deiconify()
    app.geometry("960x560")  # 最小サイズ
    f = _load(app)
    app.update()
    # ラベル欄は最大 320px（ウィンドウ幅に合わせて無制限には伸びない）
    assert f.item_rows[0].label_e.winfo_width() <= 330
    # フッターのボタンは常に見えている
    assert f.export_btn.winfo_ismapped() and f.back_btn.winfo_ismapped()

    # サイドバーと項目の表は別々にスクロールする（ホイールはカーソルの下の領域だけ）
    ev = type("E", (), {"widget": f.item_rows[0].label_e, "num": 5, "delta": 0,
                        "x_root": f.item_rows[0].label_e.winfo_rootx() + 2,
                        "y_root": f.item_rows[0].label_e.winfo_rooty() + 2})()
    side_before = f.side.canvas.yview()[0]
    items_before = f.items_area.canvas.yview()[0]
    for _ in range(10):
        app._on_wheel(ev)
    app.update()
    assert f.items_area.canvas.yview()[0] > items_before
    assert f.side.canvas.yview()[0] == side_before


def test_offset_section_toggle_and_error_navigation(app):
    app.deiconify()
    f = _load(app)
    # 折りたたむと要約 1 行になり、開閉状態は設定に入る（出力成功時に保存）
    f.toggle_offset(False)
    app.update()
    assert not f.offset_body.winfo_manager() and f.offset_summary.winfo_manager()
    assert "間隔 1000 / 200 ms" in f.offset_summary.cget("text")
    assert app.cfg["offset_section_open"] is False
    # エラーがあると自動で開き、見出し右に件数
    f.uart_int.set("0")
    f.recalc()
    app.update()
    assert f.offset_open and "1件" in f.offset_status.cget("text")
    assert "計算できません" in f.mode_label.cget("text")
    f.uart_int.set("1000")
    f.recalc()

    # 項目のエラー：行の背景・状態列・タブのバッジ・フッター一覧
    rows = {r.key: r for r in f.item_rows}
    rows["CH1"].coef.set("abc")
    app.update()
    assert rows["CH1"].state_ == "error" and "係数" in rows["CH1"].status.cget("text")
    assert f.nb.tabs[0]["count"] == "エラー 1" and f.nb.tabs[0]["error"]
    assert f.err_panel.winfo_manager()
    err = f.errors[0]
    assert err.place == "項目 › CH1"
    # 場所チップをクリック：グラフタブを開いていても項目タブに切り替え、入力欄にフォーカス
    f.nb.select(1)
    f.goto_error(err)
    app.update()
    assert f.nb.select() == 0
    # ウィンドウマネージャーのない環境でも確認できるよう「このウィンドウで最後にフォーカスした部品」を見る
    assert str(app.tk.call("focus", "-lastfor", str(app))) == str(rows["CH1"].coef_e)
    rows["CH1"].coef.set("1")
    app.update()
    assert not f.err_panel.winfo_manager()
    assert "警告 1件" in f.status_lb.cget("text")  # fixture のデータには欠落がある（警告のみ）

    # 採用チェックでタブのバッジとグループ見出しの件数が即更新
    before = f.nb.tabs[0]["count"]
    rows["CH2"].enabled.set(False)
    app.update()
    assert f.nb.tabs[0]["count"] != before
    assert "2CH中 1CHを出力" in f.group_labels["logger"].cget("text")


def test_warning_only_allows_export(app, tmp_path):
    # 欠落があるデータ（fixture）は警告のみ：出力ボタンは有効
    f = _load(app)
    assert f.warnings and not f.errors
    assert "警告" in f.status_lb.cget("text")
    assert not f.export_btn.instate(["disabled"])


def test_dark_theme(tmp_path):
    import json

    import gui

    (tmp_path / "config.json").write_text(json.dumps({"theme": "dark"}), encoding="utf-8")
    try:
        a = gui.App(config_path=str(tmp_path / "config.json"))
    except tk.TclError:
        pytest.skip("ディスプレイがありません")
    try:
        assert a.theme.name == "dark"
        assert a.cget("bg").upper() == "#1C1C1C"
        a.set_theme("light")
        assert a.cget("bg").upper() == "#FAFAFA"
        assert a.select_frame.cget("bg").upper() == "#FAFAFA"
    finally:
        a.destroy()


def test_drop_files(app, tmp_path):
    fs = app.select_frame
    app.uart_path.set("")
    app.logger_path.set("")
    u = write_uart_csv(str(tmp_path / "a.csv"), [0, 1000])
    lg = write_logger_csv(str(tmp_path / "b.CSV"), 3)
    # 2つ同時に枠へドロップ → 中身で判別（順番が逆でもよい）
    fs.set_dropped_files([lg, u])
    assert app.uart_path.get() == os.path.normpath(u)
    assert app.logger_path.get() == os.path.normpath(lg)
    assert not fs.load_btn.instate(["disabled"])
    # 欄に直接ドロップした場合はその欄に入る
    fs.set_dropped_files([u], "logger")
    assert app.logger_path.get() == os.path.normpath(u)
    # 判別できないファイルは警告
    import gui

    other = tmp_path / "x.csv"
    other.write_text("a,b\n", encoding="utf-8")
    shown = []
    orig = gui.messagebox.showwarning
    gui.messagebox.showwarning = lambda *a, **k: shown.append(a)
    try:
        fs.set_dropped_files([str(other)])
    finally:
        gui.messagebox.showwarning = orig
    assert shown and "x.csv" in shown[0][1]


def test_window_icon(app):
    import gui

    assert os.path.exists(gui.resource_path("assets", "LogMerger.ico"))
    assert app.icon_set  # アイコンを設定できた


def test_color_palette_opens_next_to_swatch(app):
    app.deiconify()
    app.geometry("1120x720+0+0")
    f = _load(app)
    app.update()
    for r in (f.item_rows[1], f.item_rows[-1]):  # 上の方の行・一番下の行
        f.items_area.see(r.color_btn)
        app.update()
        r.color_btn.invoke()
        app.update()
        pal = r._palette
        btn = r.color_btn
        bx, by, bh = btn.winfo_rootx(), btn.winfo_rooty(), btn.winfo_height()
        px, py = pal.winfo_rootx(), pal.winfo_rooty()
        ph, pw = pal.winfo_height(), pal.winfo_width()
        # 色見本のすぐ下か、すぐ上に出る
        assert abs(py - (by + bh + 2)) <= 4 or abs((py + ph) - (by - 2)) <= 4, (py, ph, by)
        # 横方向は色見本と重なる位置（離れた場所に出ない）
        assert px <= bx + btn.winfo_width() and px + pw >= bx
        # アプリのウィンドウの中に収まる
        assert px >= app.winfo_rootx() - 1 and px + pw <= app.winfo_rootx() + app.winfo_width() + 1
        pal.close()
        app.update()


def test_graph_x_unit_radio(app):
    f = _load(app)
    g = f.graph_rows[0]
    assert g.x == "elapsed_ms" and g.x_unit.get() == "ms"
    assert not g.unit_radios[0].instate(["disabled"])
    g.x_unit.set("min")
    assert g.to_setting().x_unit == "min" and g.to_setting().effective_x_unit() == "min"
    # 横軸を時間以外（電圧）にすると単位は選べない
    labels = list(g.x_cb.cget("values"))
    g.x_cb.current(labels.index("電圧(mV)"))
    g.cbs[1].current(list(g.cbs[1].cget("values")).index("電流(mA)"))
    g._on_select()
    assert all(rb.instate(["disabled"]) for rb in g.unit_radios)
    assert g.to_setting().x_unit is None
    # 経過時間(s) にすると選べるようになり、元の単位 s から
    g.x_cb.current(labels.index("経過時間(s)"))
    g._on_select()
    assert not g.unit_radios[0].instate(["disabled"]) and g.x_unit.get() == "s"
