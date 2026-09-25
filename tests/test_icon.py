"""アプリアイコン（assets/LogMerger.ico）の形式の確認。"""

import os
import struct
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))


def _entries(path):
    d = open(path, "rb").read()
    _, typ, n = struct.unpack("<HHH", d[:6])
    assert typ == 1
    out = []
    for i in range(n):
        w, h, _, _, _, bpp, size, off = struct.unpack("<BBBBHHII", d[6 + 16 * i : 22 + 16 * i])
        blob = d[off : off + size]
        out.append((w or 256, bpp, "PNG" if blob[:8] == b"\x89PNG\r\n\x1a\n" else "BMP"))
    return out


def test_ico_has_all_sizes_and_standard_format():
    entries = _entries(os.path.join(ROOT, "assets", "LogMerger.ico"))
    assert [e[0] for e in entries] == [16, 20, 24, 32, 40, 48, 64, 128, 256]
    assert all(bpp == 32 for _, bpp, _ in entries)
    # 48px 以下は BMP（Windows の多くの API が確実に使える）、64px 以上は PNG
    assert all(kind == ("BMP" if s <= 48 else "PNG") for s, _, kind in entries)


def test_ico_is_reproducible(tmp_path):
    import make_ico

    out = tmp_path / "x.ico"
    make_ico.build(out=str(out))
    assert out.read_bytes() == open(os.path.join(ROOT, "assets", "LogMerger.ico"), "rb").read()
