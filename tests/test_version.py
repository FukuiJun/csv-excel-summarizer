import os
import sys

import version

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import write_version  # noqa: E402


def test_dev_version_without_version_file(monkeypatch):
    monkeypatch.setitem(sys.modules, "_version", None)  # _version.py がない状態
    assert version.get_version() == version.DEV_VERSION
    assert version.window_title("ツール") == "ツール 開発版"


def test_version_from_build_file(monkeypatch):
    mod = type(sys)("_version")
    mod.VERSION = "v1.2.0"
    monkeypatch.setitem(sys.modules, "_version", mod)
    assert version.window_title("ツール") == "ツール v1.2.0"


def test_version_text():
    from datetime import datetime

    t = write_version.version_text("v1.2.0", "abc1234", datetime(2026, 10, 2, 15, 30))
    assert "バージョン : v1.2.0" in t and "2026-10-02 15:30" in t and "abc1234" in t
    assert "releases/tag/v1.2.0" in t and t.endswith("\r\n")
    assert "開発版" in write_version.version_text("dev-abc1234", "abc1234", datetime(2026, 10, 2))


def test_write_files(tmp_path, monkeypatch):
    monkeypatch.setattr(write_version, "ROOT", str(tmp_path))
    assert write_version.main(["v9.9.9"]) == 0
    first = (tmp_path / "_version.py").read_text(encoding="utf-8")
    assert "VERSION = 'v9.9.9'" in first
    out = tmp_path / "dist"
    out.mkdir()
    assert write_version.main(["v9.9.9", str(out)]) == 0
    assert (tmp_path / "_version.py").read_text(encoding="utf-8") == first  # ビルド前の日時のまま
    raw = (out / "VERSION.txt").read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf") and "v9.9.9" in raw.decode("utf-8-sig")
