"""ビルド用：_version.py と配布フォルダの VERSION.txt を作る。

    python tools/write_version.py <バージョン> [<配布フォルダ>]

<バージョン> はタグ名（例 v1.2.0）か dev-<コミット>。ビルド前に配布フォルダなしで実行して
_version.py を作り（exe に含まれる）、ビルド後に配布フォルダを指定して VERSION.txt
（UTF-8 BOM 付き。メモ帳で文字化けしない）を書く。
"""

from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime, timedelta, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REPO_URL = "https://github.com/FukuiJun/csv-excel-summarizer"
JST = timezone(timedelta(hours=9))


def _commit() -> str:
    sha = os.environ.get("GITHUB_SHA", "")
    if sha:
        return sha[:7]
    try:
        return subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=ROOT, capture_output=True,
                              text=True, check=True).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return "不明"


def version_text(version: str, commit: str, built: datetime) -> str:
    lines = [
        "ロガーCSV・UART CSV 統合ツール（LogMerger）",
        "",
        f"バージョン : {version}",
        f"ビルド日時 : {built:%Y-%m-%d %H:%M}（日本時間）",
        f"コミット   : {commit}",
    ]
    if version.startswith("v"):
        lines.append(f"リリース   : {REPO_URL}/releases/tag/{version}")
    else:
        lines.append("※ リリース前の開発版です")
    return "\r\n".join(lines) + "\r\n"


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 2
    version = argv[0]
    commit = _commit()
    built = datetime.now(JST)
    path = os.path.join(ROOT, "_version.py")
    prev: dict = {}
    if os.path.exists(path):  # ビルド前に作った _version.py と同じ日時にそろえる
        with open(path, encoding="utf-8") as f:
            exec(f.read(), prev)
    if prev.get("VERSION") == version and prev.get("BUILD_DATE"):
        built = datetime.strptime(prev["BUILD_DATE"], "%Y-%m-%d %H:%M")
    else:
        with open(path, "w", encoding="utf-8") as f:
            f.write(f'VERSION = {version!r}\nCOMMIT = {commit!r}\nBUILD_DATE = "{built:%Y-%m-%d %H:%M}"\n')
    if len(argv) > 1:
        with open(os.path.join(argv[1], "VERSION.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write(version_text(version, commit, built))
    print(f"version: {version} ({commit})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
