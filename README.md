# ロガーCSV・UART CSV 統合ツール

GRAPHTEC GL240（データロガー）の CSV と、マイコンから UART 経由で受信した CSV を読み込み、
時刻をそろえて 1 つの Excel ファイル（.xlsx）にまとめるツールです。
出力前に GUI で、間隔・出力する項目・ラベル・値の換算（係数・オフセット）・グラフ・出力先を調整できます。

仕様：[spec_log_merger.md（v4）](docs/spec_log_merger.md)

## 動作環境

- Windows（Python 3.10 以上。exe 版は Python 不要）
- ライブラリ：openpyxl、tkinter（Python 標準）

## 使い方

### 起動

- exe 版：`LogMerger\LogMerger.exe` をダブルクリック（exe の入手方法は下記）
- Python 版：
  ```
  pip install -r requirements.txt
  python main.py
  ```

### exe の入手方法

exe はリポジトリには含まれていません（GitHub の「Code → Download ZIP」で落としてもソースのみで、`LogMerger` フォルダはありません）。
次のいずれかの方法で入手してください。

1. **GitHub Actions でビルドされたものをダウンロード**（おすすめ）
   1. GitHub のリポジトリページで **Actions** タブを開く
   2. 左の一覧から **Build exe** を選び、緑のチェック（成功）が付いた最新の実行を開く
   3. ページ下部の **Artifacts** にある **LogMerger** をクリックしてダウンロード（GitHub へのログインが必要。保存期間は 90 日）
   4. zip を展開し、`LogMerger\LogMerger.exe` を実行
2. **Releases からダウンロード**：`v1.0.0` のようなタグを push すると、Releases に `LogMerger.zip` が添付されます
3. **自分の PC でビルド**：Python をインストールした PC で `build.bat` を実行すると `dist\LogMerger\LogMerger.exe` ができます（下記「exe のビルド手順」）

※ 署名のない exe のため、初回起動時に Windows の SmartScreen 警告が出ることがあります。その場合は「詳細情報」→「実行」を押してください。

### 画面1：ファイル選択

1. 「UART CSV」「ロガー CSV」の［参照］でファイルを選ぶ（パスを直接入力しても可）
2. ［読み込み］を押す。形式が違う場合は原因を表示して画面1にとどまります

### 画面2：出力内容の調整

| 欄 | 内容 |
|---|---|
| 読み込み結果 | 行数・時刻・スキップ行数・重複行数・欠落の警告。［詳細］で欠落箇所の一覧 |
| 間隔 | UART間隔（初期値：`pc_timestamp` の差の中央値を100ms単位に丸めた値）とロガー間隔（初期値：ヘッダの測定間隔）。変更するとその場で再計算します |
| 出力する項目 | 採用・ラベル・係数・オフセット。出力値は `元の値 × 係数 ＋ オフセット` |
| グラフ | 1行1グラフ。横軸（初期値 elapsed_ms、経過時間(s) も選べる）・第1軸（左）は必須、第2軸（右）は「なし」も可。0個ならグラフシートを作りません |
| 出力先 | 初期値は UART CSV のフォルダと `解析_<UART 1行目の時刻>.xlsx` |

入力にエラーがある欄は赤くなり、画面下にエラー内容が表示され、［Excelに出力］は押せなくなります。

- ［初期値に戻す］：出力する項目とグラフをプログラム内の既定値に戻します
- ［戻る］：画面1に戻ります（選んだファイルは保持）
- ［Excelに出力］：出力します。成功すると設定を `config.json` に保存し、次回の初期値になります

### 時刻合わせのしくみ

- UART の 1 行目の `pc_timestamp` とロガーの番号1 を経過時間 0 とします（両方の記録を同時に開始する運用が前提）
- ロガーの内部時刻は使わず、ロガー各行の経過時間は `(番号 − 1) × ロガー間隔` とみなします
- UART の時刻の差が間隔の 1.5 倍を超える箇所は受信の欠落とみなし、空欄の行を補完します（Excel では灰色の行）。グラフの横軸に使えるよう、補完行の elapsed_ms だけは「直前の値＋経過時間」の推定値を入れます。0.5 倍未満の行は二重受信として捨てます
- 間隔の長い方を基準に、短い方から時刻の最も近い行を採用します（例：UART 1000ms・ロガー 200ms ならロガーを約5行ごとに採用）

### 出力される Excel

| シート | 内容 |
|---|---|
| 解析 | 測定日時・測定機、UART の採用列（経過時間(s)が先頭）、1列空けてロガーの採用CH |
| グラフ | 散布図（直線・マーカーなし）。横軸はグラフごとに選択（初期値は elapsed_ms）、空欄は線を途切れさせて表示 |
| UART CSV のファイル名 | UART CSV の全データ（換算・補完なし） |
| ロガー CSV のファイル名 | ロガー CSV の全内容（換算なし） |

出力する行数が最も多いシートで 10 万行を超えると確認ダイアログ、100 万行を超えるとエラーになります（`config.json` で変更可）。

## config.json の書き方

exe（または `main.py`）と同じフォルダに置きます。ない場合・壊れている場合は既定値で作り直されます。
画面2の設定は［Excelに出力］が成功したときに自動で保存されるので、通常は手で編集する必要はありません。
UTF-8 で保存してください。

```json
{
  "uart_columns": {
    "voltage_mV": {"enabled": true, "label": "電圧(V)", "coef": 0.001, "offset": 0},
    "t_ms":       {"enabled": false, "label": "t_ms", "coef": 1, "offset": 0}
  },
  "logger_channels": {
    "CH1": {"enabled": true, "label": "CH1(mV)", "coef": 1, "offset": 0, "unit": "mV"}
  },
  "logger_label_format": "{ch}({unit})",
  "elapsed_label": "経過時間(s)",
  "graphs": [
    {"x": "elapsed_ms", "primary": "voltage_mV", "secondary": "current_mA"},
    {"x": "@elapsed", "primary": "soc_percent", "secondary": null}
  ],
  "last_dirs": {"uart": "C:\\data\\uart", "logger": "C:\\data\\logger"},
  "output_filename_format": "解析_{start:%y%m%d-%H%M%S}.xlsx",
  "row_warn_threshold": 100000,
  "row_limit": 1000000
}
```

| キー | 内容 |
|---|---|
| `uart_columns` | UART の列ごと（キーは CSV の元の列名）の 採用・ラベル・係数・オフセット |
| `logger_channels` | ロガーの CH ごと（キーは `CH1` など）の設定と、前回の単位 `unit`。単位が前回と違うと画面2に警告が出ます |
| `logger_label_format` | 設定がない CH のラベルの書式。`{ch}` が CH 番号、`{unit}` が単位に置き換わります |
| `elapsed_label` | 経過時間(s)列のラベル |
| `graphs` | グラフの一覧。横軸 `x`・第1軸 `primary`・第2軸 `secondary` は列名・CH 番号で指定（`secondary` の `null` は「なし」、`x` の `"@elapsed"` は経過時間(s)。`x` を省略すると `elapsed_ms`） |
| `last_dirs` | ［参照］ダイアログで最初に開くフォルダ |
| `output_filename_format` | 出力ファイル名の書式。`{start:...}` に UART 1 行目の時刻が入ります（`%y%m%d` などは Python の strftime 書式） |
| `row_warn_threshold` / `row_limit` | 行数の警告しきい値 / 上限 |

- 今回の CSV にない列・CH の設定は削除されずに残ります（CH 数を一時的に減らした場合など）
- 間隔・出力先フォルダ・ファイル名は保存されません（毎回 CSV から決まります）

## exe のビルド手順（Windows）

`--onedir` 形式（フォルダ配布）でビルドします。`--onefile` は起動が遅くなるため使いません。

```
pip install -r requirements-dev.txt
pyinstaller --noconfirm --clean --onedir --windowed --name LogMerger main.py
```

（`build.bat` を実行しても同じです）

GitHub に push すると、GitHub Actions（`.github/workflows/build-exe.yml`）が Windows 上でテストと exe のビルドを自動で行い、
結果を Artifacts（`LogMerger`）としてアップロードします。タグ `v*` を push したときは Releases にも `LogMerger.zip` を添付します。

`dist\LogMerger\` フォルダができるので、フォルダごと配布してください。
`config.json` は初回起動時に `LogMerger.exe` と同じフォルダに作られます。
あらかじめ調整した `config.json` を同じフォルダに入れて配布することもできます。

## 開発

### ファイル構成

| ファイル | 内容 |
|---|---|
| `main.py` | 起動スクリプト |
| `gui.py` | GUI（画面1・画面2） |
| `uart_reader.py` | UART CSV の読み込み |
| `logger_reader.py` | ロガー CSV の読み込み（ヘッダ解析） |
| `merger.py` | 欠落補完・時刻合わせ・換算・行数チェック |
| `excel_writer.py` | Excel 出力 |
| `exporter.py` | Excel 出力と、成功時の設定保存 |
| `config.py` | `config.json` の読み書きと前回設定の復元 |
| `models.py` | 設定のデータ構造 |

GUI 以外のモジュールは tkinter に依存しないので、GUI なしで実行・テストできます。

### テスト

```
pip install -r requirements-dev.txt
python -m pytest tests
```

- `tests/data/` に入力サンプルと期待出力 `expected_log40.xlsx` があります
- 期待出力の作り直し：`python tests/make_expected.py <UART CSV> <ロガー CSV> tests/data/expected_log40.xlsx`
  - **注意**：現在のサンプル（`log40.csv` と `260924-090155.CSV`）は計測時間帯が重なっていません。
    同時に計測した実データが用意できたら、そのデータで期待出力を作り直してください（仕様 9章）
- GUI のテスト（`test_gui_smoke.py`）はディスプレイがない環境では自動でスキップされます
