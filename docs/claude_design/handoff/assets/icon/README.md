# LogMerger アプリアイコン（採用案：BE-1 P1）

青（UART）とオレンジ（ロガー）の2枚が重なった部分に、影付きの折れ線グラフ。

## ファイル

| ファイル | 用途 |
|---|---|
| `LogMerger.ico` | exe・ウィンドウ用。16/20/24/32/40/48/64/128/256px を1つに収録 |
| `LogMerger_16.png` 〜 `LogMerger_256.png` | tkinter の `iconphoto` 用、ドキュメント用 |
| `LogMerger_512.png` | 資料・スライドなど大きく載せる用 |
| `LogMerger.svg` | 元データ（32px 以上の絵柄） |
| `LogMerger_small.svg` | 元データ（24px 以下の絵柄。線の頂点を減らし、太くしたもの） |

色：青 #005FB8、オレンジ #ED7D31、白 #FFFFFF（背景は透明）

## 使い方

### exe のアイコン（PyInstaller）

```
pyinstaller --onefile --windowed --icon=LogMerger.ico --name LogMerger main.py
```

### ウィンドウ（タイトルバー・タスクバー）のアイコン（tkinter）

```python
import sys, tkinter as tk
from pathlib import Path

def resource(name):  # PyInstaller の onefile でも読めるようにする
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return str(base / "assets" / name)

root = tk.Tk()
if sys.platform == "win32":
    root.iconbitmap(default=resource("LogMerger.ico"))  # 補助ウィンドウにも適用される
else:
    icons = [tk.PhotoImage(file=resource(f"LogMerger_{s}.png")) for s in (16, 32, 48, 256)]
    root.iconphoto(True, *icons)
```

- PyInstaller では `--add-data "assets/LogMerger.ico;assets"` で ico を同梱する
- タスクバーで Python のアイコンになる場合は、起動直後に
  `ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("LogMerger")` を呼ぶ
