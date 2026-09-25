"""設定ファイル config.json の読み書きと、前回設定の復元（仕様 2.1・2.2）。"""

from __future__ import annotations

import copy
import json
import os
import sys
from dataclasses import dataclass
from datetime import datetime

from logger_reader import LoggerData
from models import (
    DEFAULT_X_KEY,
    ELAPSED_KEY,
    SOURCE_LOGGER,
    SOURCE_UART,
    TIME_UNITS,
    GraphSetting,
    ItemSetting,
    default_color,
    is_color,
)
from uart_reader import UartData

CONFIG_FILENAME = "config.json"

DEFAULT_UART_COLUMNS: dict[str, dict] = {
    "t_ms": {"enabled": False, "label": "t_ms", "coef": 1},
    "elapsed_ms": {"enabled": True, "label": "time(ms)", "coef": 1},
    "voltage_mV": {"enabled": True, "label": "電圧(mV)", "coef": 1},
    "current_mA": {"enabled": True, "label": "電流(mA)", "coef": 1},
    "cap_mAh": {"enabled": True, "label": "CAP(mAh)", "coef": 1},
    "cap_max_mAh": {"enabled": False, "label": "CAP_MAX(mAh)", "coef": 1},
    "soc_percent": {"enabled": True, "label": "SOC(％)", "coef": 1},
    "soh_percent": {"enabled": False, "label": "SOH(％)", "coef": 1},
    "temp_C": {"enabled": True, "label": "温度(℃)", "coef": 1},
}

DEFAULT_GRAPHS: list[dict] = [{"x": DEFAULT_X_KEY, "primary": ["voltage_mV", None], "secondary": ["current_mA", None]}]

DEFAULT_CONFIG: dict = {
    "uart_columns": DEFAULT_UART_COLUMNS,
    "logger_channels": {},
    "logger_label_format": "{ch}({unit})",
    "elapsed_label": "経過時間(s)",
    "graphs": DEFAULT_GRAPHS,
    "last_dirs": {"uart": "", "logger": ""},
    "output_filename_format": "解析_{start:%y%m%d-%H%M%S}.xlsx",
    "row_warn_threshold": 100_000,
    "row_limit": 1_000_000,
    "theme": "auto",  # auto（Windows の設定に追従）/ light / dark
    "offset_section_open": True,  # 画面2「間隔・時間オフセット」の開閉状態
}


def default_config() -> dict:
    return copy.deepcopy(DEFAULT_CONFIG)


def app_dir() -> str:
    """exe（または main.py）のあるフォルダ。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def default_config_path() -> str:
    return os.path.join(app_dir(), CONFIG_FILENAME)


# ---------------------------------------------------------------------------
# 読み込み・保存
# ---------------------------------------------------------------------------


def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _clean_keys(v) -> list[str | None]:
    """グラフの軸の設定を「キーのリスト」にそろえる（旧形式の文字列・null も受け付ける）。"""
    if v is None or isinstance(v, str):
        v = [v]
    if not isinstance(v, list):
        return []
    return [k if isinstance(k, str) and k else None for k in v]


def _clean_item(v, with_unit: bool) -> dict | None:
    if not isinstance(v, dict):
        return None
    out = {
        "enabled": bool(v.get("enabled", True)),
        "label": str(v.get("label", "")),
        "coef": v.get("coef", 1) if _is_num(v.get("coef", 1)) else 1,
    }
    if is_color(v.get("color")):
        out["color"] = v["color"].upper()
    if with_unit:
        out["unit"] = str(v.get("unit", ""))
    return out


def _normalize(raw: dict) -> dict:
    """読み込んだ内容を検査し、不正な項目は既定値に置き換える。"""
    cfg = default_config()
    if isinstance(raw.get("uart_columns"), dict):
        cols = {}
        for k, v in raw["uart_columns"].items():
            c = _clean_item(v, with_unit=False)
            if c is not None:
                cols[str(k)] = c
        cfg["uart_columns"] = cols
    if isinstance(raw.get("logger_channels"), dict):
        chs = {}
        for k, v in raw["logger_channels"].items():
            c = _clean_item(v, with_unit=True)
            if c is not None:
                chs[str(k)] = c
        cfg["logger_channels"] = chs
    for key in ("logger_label_format", "elapsed_label", "output_filename_format"):
        if isinstance(raw.get(key), str) and raw[key]:
            cfg[key] = raw[key]
    if isinstance(raw.get("graphs"), list):
        graphs = []
        for g in raw["graphs"]:
            if isinstance(g, dict):
                x = g.get("x", DEFAULT_X_KEY)  # 横軸の設定がない古い config.json は初期値
                xu = g.get("x_unit")
                if xu == "h":  # h は廃止。以前の設定は一番近い min に読み替える
                    xu = "min"
                graphs.append(
                    {
                        "x": x if isinstance(x, str) else None,
                        "x_unit": xu if xu in TIME_UNITS else None,
                        "primary": _clean_keys(g.get("primary")),
                        "secondary": _clean_keys(g.get("secondary")),
                    }
                )
        cfg["graphs"] = graphs
    if isinstance(raw.get("last_dirs"), dict):
        for k in ("uart", "logger"):
            if isinstance(raw["last_dirs"].get(k), str):
                cfg["last_dirs"][k] = raw["last_dirs"][k]
    if raw.get("theme") in ("auto", "light", "dark"):
        cfg["theme"] = raw["theme"]
    if isinstance(raw.get("offset_section_open"), bool):
        cfg["offset_section_open"] = raw["offset_section_open"]
    for key in ("row_warn_threshold", "row_limit"):
        if _is_num(raw.get(key)) and raw[key] > 0:
            cfg[key] = int(raw[key])
    return cfg


@dataclass
class LoadResult:
    config: dict
    status: str  # "ok" / "missing" / "broken"
    message: str = ""


def load_config(path: str | None = None) -> LoadResult:
    """config.json を読み込む。ない・壊れている場合は既定値で作り直す。"""
    path = path or default_config_path()
    status = "ok"
    message = ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        if not isinstance(raw, dict):
            raise ValueError("JSON の最上位がオブジェクトではありません")
        return LoadResult(_normalize(raw), "ok")
    except FileNotFoundError:
        status = "missing"
    except (OSError, ValueError) as e:  # json.JSONDecodeError は ValueError のサブクラス
        status = "broken"
        message = str(e)
    cfg = default_config()
    try:
        save_config(cfg, path)
    except OSError as e:
        message = (message + " / " if message else "") + f"既定の設定ファイルを作成できませんでした: {e}"
    return LoadResult(cfg, status, message)


def save_config(cfg: dict, path: str | None = None) -> None:
    path = path or default_config_path()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


# ---------------------------------------------------------------------------
# 前回設定の復元
# ---------------------------------------------------------------------------


def default_uart_item(col: str) -> ItemSetting:
    d = DEFAULT_UART_COLUMNS.get(col, {"enabled": True, "label": col, "coef": 1})
    return ItemSetting(col, SOURCE_UART, d["enabled"], d["label"], float(d["coef"]))


def logger_default_label(fmt: str, ch: str, unit: str) -> str:
    try:
        return fmt.format(ch=ch, unit=unit)
    except (KeyError, IndexError, ValueError):
        return f"{ch}({unit})"


def default_logger_item(ch: str, unit: str, fmt: str) -> ItemSetting:
    return ItemSetting(ch, SOURCE_LOGGER, True, logger_default_label(fmt, ch, unit), 1.0, unit=unit)


def build_items(cfg: dict, uart: UartData | None, logger: LoggerData | None,
                use_saved: bool = True) -> tuple[list[ItemSetting], list[str]]:
    """画面2の出力項目の初期値を作る。戻り値: (項目一覧, 警告メッセージ)

    uart・logger のどちらかが None（片方だけ読み込んだ）ならその分の項目は作らない。
    """
    items: list[ItemSetting] = []
    warnings: list[str] = []
    saved_u = cfg.get("uart_columns", {}) if use_saved else {}
    saved_l = cfg.get("logger_channels", {}) if use_saved else {}
    fmt = cfg.get("logger_label_format") or DEFAULT_CONFIG["logger_label_format"]
    if not use_saved:
        fmt = DEFAULT_CONFIG["logger_label_format"]

    for col in (uart.columns if uart is not None else []):
        s = saved_u.get(col)
        if s:
            items.append(ItemSetting(col, SOURCE_UART, s["enabled"], s["label"], float(s["coef"])))
        else:
            items.append(default_uart_item(col))
        items[-1].color = s["color"] if s and is_color(s.get("color")) else default_color(len(items) - 1)

    for ch in (logger.channels if logger is not None else []):
        s = saved_l.get(ch.name)
        if s:
            items.append(
                ItemSetting(ch.name, SOURCE_LOGGER, s["enabled"], s["label"], float(s["coef"]), unit=ch.unit)
            )
            if s.get("unit", "") != ch.unit:
                warnings.append(
                    f"{ch.name} の単位が前回と異なります（前回 [{s.get('unit', '')}] → 今回 [{ch.unit}]）。"
                    "ラベル・係数を確認してください。"
                )
        else:
            items.append(default_logger_item(ch.name, ch.unit, fmt))
        items[-1].color = s["color"] if s and is_color(s.get("color")) else default_color(len(items) - 1)
    return items, warnings


def build_graphs(cfg: dict, items: list[ItemSetting], use_saved: bool = True,
                 single_source: bool = False) -> list[GraphSetting]:
    """グラフ設定の初期値。今回ない・採用されていない項目は未選択にする。

    single_source=True（UART かロガーの片方だけ読み込んだ）のときは、読み込んでいない方の項目を
    エラーにせず外す：縦軸の欄は「なし」、横軸は経過時間(s)。縦軸が 1 つも残らないグラフは作らない。
    """
    src = cfg.get("graphs", DEFAULT_GRAPHS) if use_saved else DEFAULT_GRAPHS
    enabled = {it.key for it in items if it.enabled}
    present = {it.key for it in items}
    graphs = []
    for g in src:
        x = g.get("x", DEFAULT_X_KEY)
        if single_source:
            keys = [k for k in _clean_keys(g.get("primary")) + _clean_keys(g.get("secondary")) if k in present]
            if not keys:
                continue  # このデータでは描けるものがないグラフ
            g = dict(g)
            prim = [k for k in _clean_keys(g.get("primary")) if k in present]
            sec = [k for k in _clean_keys(g.get("secondary")) if k in present]
            if not prim:  # 第1軸が全部ない → 第2軸の項目を第1軸に
                prim, sec = sec[:1], sec[1:]
            g["primary"], g["secondary"] = prim, sec
            if x != ELAPSED_KEY and x not in present:
                x = ELAPSED_KEY
        # 「なし」(None) と未選択を区別するため、今回使えないキーは "" にする
        prim = [None if k is None else (k if k in enabled else "") for k in _clean_keys(g.get("primary"))]
        sec = [None if k is None else (k if k in enabled else "") for k in _clean_keys(g.get("secondary"))]
        xu = g.get("x_unit")
        g2 = GraphSetting(prim, sec, x if (x == ELAPSED_KEY or x in enabled) else None,
                          xu if xu in TIME_UNITS else None)
        if not g2.primary[0]:
            g2.primary[0] = None  # 第1軸の1つ目は必須：未選択は None
        graphs.append(g2)
    if single_source and not graphs:
        # 片方だけのとき、描けるグラフが残らなければ採用中の項目で 1 つ用意する
        ys = [it.key for it in items if it.enabled and it.key != DEFAULT_X_KEY]
        if ys:
            x = DEFAULT_X_KEY if DEFAULT_X_KEY in enabled else ELAPSED_KEY
            graphs.append(GraphSetting(ys[:1], ys[1:2], x))
    return graphs


def elapsed_label(cfg: dict, use_saved: bool = True) -> str:
    if not use_saved:
        return DEFAULT_CONFIG["elapsed_label"]
    return cfg.get("elapsed_label") or DEFAULT_CONFIG["elapsed_label"]


def apply_settings(
    cfg: dict,
    items: list[ItemSetting],
    graphs: list[GraphSetting],
    elapsed: str,
    uart_dir: str | None = None,
    logger_dir: str | None = None,
    save_graphs: bool = True,
) -> dict:
    """出力成功時に保存する内容を cfg に反映した新しい dict を返す。

    今回のCSVにない列・CHの設定は残す。save_graphs=False（片方だけの出力）のときは
    両方そろったときのグラフ設定を壊さないよう、グラフ設定は保存しない。
    """
    new = copy.deepcopy(cfg)
    new.setdefault("uart_columns", {})
    new.setdefault("logger_channels", {})
    for it in items:
        d = {"enabled": it.enabled, "label": it.label, "coef": it.coef, "color": it.color}
        if it.source == SOURCE_UART:
            new["uart_columns"][it.key] = d
        else:
            d["unit"] = it.unit
            new["logger_channels"][it.key] = d
    if save_graphs:
        new["graphs"] = [
            {"x": g.x, "x_unit": g.x_unit, "primary": [k or None for k in g.primary],
             "secondary": [k or None for k in g.secondary]}
            for g in graphs
        ]
    new["elapsed_label"] = elapsed
    new.setdefault("last_dirs", {"uart": "", "logger": ""})
    if uart_dir:
        new["last_dirs"]["uart"] = uart_dir
    if logger_dir:
        new["last_dirs"]["logger"] = logger_dir
    return new


def make_output_filename(cfg: dict, start: datetime | None) -> str:
    if start is None:  # 開始時刻が分からない（ロガーのみで日時なし等）→ 今の時刻
        start = datetime.now()
    fmt = cfg.get("output_filename_format") or DEFAULT_CONFIG["output_filename_format"]
    try:
        name = fmt.format(start=start)
    except (KeyError, IndexError, ValueError, AttributeError):
        name = DEFAULT_CONFIG["output_filename_format"].format(start=start)
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name
