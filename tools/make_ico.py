"""アプリアイコン assets/LogMerger.ico を、資料の PNG から標準的な形式で作り直す。

    python tools/make_ico.py

- 16〜48px：BMP（32bit BGRA ＋ 透過マスク）で格納（Windows の多くの API・Tk が確実に使える形式）
- 64px 以上：PNG のまま格納（Windows Vista 以降の標準）
資料の ico は全サイズが PNG 格納だったため、小さいサイズが使われずに拡大縮小されてぼやけることがあった。
Pillow などの外部ライブラリは不要（標準ライブラリの zlib で PNG を読む）。
"""

from __future__ import annotations

import os
import struct
import sys
import zlib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "docs", "claude_design", "handoff", "assets", "icon")
OUT = os.path.join(ROOT, "assets", "LogMerger.ico")
SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]
BMP_MAX = 48  # これ以下のサイズは BMP で格納


def read_png_rgba(path: str) -> tuple[int, int, list[bytes]]:
    """8bit RGBA・インターレースなしの PNG を読み、(幅, 高さ, 各行の RGBA バイト列) を返す。"""
    data = open(path, "rb").read()
    assert data[:8] == b"\x89PNG\r\n\x1a\n", path
    pos, idat = 8, b""
    while pos < len(data):
        length, ctype = struct.unpack(">I4s", data[pos : pos + 8])
        chunk = data[pos + 8 : pos + 8 + length]
        if ctype == b"IHDR":
            w, h, depth, color, _, _, interlace = struct.unpack(">IIBBBBB", chunk)
            if (depth, color, interlace) != (8, 6, 0):
                raise ValueError(f"{path}: 8bit RGBA・インターレースなしの PNG のみ対応です")
        elif ctype == b"IDAT":
            idat += chunk
        pos += 12 + length
    raw = zlib.decompress(idat)
    bpp, stride = 4, w * 4
    rows, prev = [], bytearray(stride)
    for y in range(h):
        f = raw[y * (stride + 1)]
        line = bytearray(raw[y * (stride + 1) + 1 : (y + 1) * (stride + 1)])
        for x in range(stride):
            a = line[x - bpp] if x >= bpp else 0
            b = prev[x]
            c = prev[x - bpp] if x >= bpp else 0
            if f == 1:
                line[x] = (line[x] + a) & 0xFF
            elif f == 2:
                line[x] = (line[x] + b) & 0xFF
            elif f == 3:
                line[x] = (line[x] + ((a + b) >> 1)) & 0xFF
            elif f == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                pr = a if pa <= pb and pa <= pc else (b if pb <= pc else c)
                line[x] = (line[x] + pr) & 0xFF
        rows.append(bytes(line))
        prev = line
    return w, h, rows


def bmp_entry(w: int, h: int, rows: list[bytes]) -> bytes:
    """ico 用の BMP（BITMAPINFOHEADER ＋ 下から上の BGRA ＋ AND マスク）。"""
    header = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 32, 0, 0, 0, 0, 0, 0)
    pixels = bytearray()
    mask = bytearray()
    mask_stride = ((w + 31) // 32) * 4
    for row in reversed(rows):
        mrow = bytearray(mask_stride)
        for x in range(w):
            r, g, b, a = row[x * 4 : x * 4 + 4]
            pixels += bytes((b, g, r, a))
            if a == 0:
                mrow[x // 8] |= 0x80 >> (x % 8)
        mask += mrow
    return header + bytes(pixels) + bytes(mask)


def build(src_dir: str = SRC, out: str = OUT) -> list[tuple[int, str]]:
    entries = []
    for s in SIZES:
        path = os.path.join(src_dir, f"LogMerger_{s}.png")
        if s <= BMP_MAX:
            w, h, rows = read_png_rgba(path)
            assert (w, h) == (s, s), path
            entries.append((s, "BMP", bmp_entry(w, h, rows)))
        else:
            entries.append((s, "PNG", open(path, "rb").read()))
    head = struct.pack("<HHH", 0, 1, len(entries))
    offset = 6 + 16 * len(entries)
    dirs, blobs = b"", b""
    for s, _, blob in entries:
        dim = 0 if s >= 256 else s
        dirs += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(blob), offset)
        offset += len(blob)
        blobs += blob
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "wb") as f:
        f.write(head + dirs + blobs)
    return [(s, kind) for s, kind, _ in entries]


if __name__ == "__main__":
    result = build(*sys.argv[1:3])
    print("written:", OUT if len(sys.argv) < 3 else sys.argv[2])
    for s, kind in result:
        print(f"  {s:>3}px  {kind}")
