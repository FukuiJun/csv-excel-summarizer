# ロガーCSV・UART CSV 統合ツール

GRAPHTEC GL240（データロガー）の CSV と、マイコンから UART 経由で受信した CSV を読み込み、
時刻をそろえて 1 つの Excel ファイル（.xlsx）にまとめるツールです。
出力前に GUI で、間隔・出力する項目・ラベル・値の換算（係数）・時間オフセット・グラフ・出力先を調整できます。

仕様：[spec_log_merger.md（v4）](docs/spec_log_merger.md)

## 動作環境

- Windows（Python 3.10 以上。exe 版は Python 不要）
- ライブラリ：openpyxl、tkinter（Python 標準）、sv-ttk（画面のテーマ）、tkinterdnd2（ドラッグ＆ドロップ）

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

画面の上部に手順（①ファイル選択 › ②出力内容の調整）、下部にボタンが並びます。画面のデザインは
`docs/claude_design/handoff/`（Claude Design の引き継ぎ資料）に基づいています。改修後の画面は `docs/claude_design/implemented/` にあります。

### 画面1：ファイル選択

1. CSV ファイルを破線の枠にドラッグ＆ドロップする（2つ同時可。中身から UART／ロガーを自動判別）か、
   「UART CSV」「ロガー CSV」の［参照…］で選ぶ（パスを直接入力しても可）
   - UART CSV 欄・ロガー CSV 欄の上に落とすとその欄に入ります。入力済みの欄には緑のチェックが付きます
2. ［読み込み →］を押す（2つそろうまで押せません）。形式が違う場合は原因を表示して画面1にとどまります

### 画面2：出力内容の調整

左のサイドバーに「読み込み結果」「間隔・時間オフセット」「出力先」、右のタブに「出力する項目」「グラフ」があります。
サイドバーと表はそれぞれ別にスクロールします（マウスホイールはカーソルの下の領域だけ動きます）。

| 欄 | 内容 |
|---|---|
| 読み込み結果 | UART・ロガーの行数・時刻などのタイルと、状態の帯（欠落なし／欠落の警告＋［詳細…］／単位変更の警告） |
| 間隔・時間オフセット | UART間隔（初期値：`pc_timestamp` の差の中央値を100ms単位に丸めた値）とロガー間隔（初期値：ヘッダの測定間隔）。時間オフセットは UART 全体・ロガー全体の時刻をずらす量（ms、＋で遅らせる、初期値 0、保存しない）。例：ロガーの変化が UART より1秒早く出るときはロガーに 1000。変更するとその場で再計算します。見出しをクリックで折りたたみ（要約 1 行表示）、エラーがあると自動で開きます |
| 出力する項目 | 採用・ラベル・係数・グラフの色。出力値は `元の値 × 係数`。色の見本をクリックするとカラーパレットが開き、色を選べます（「その他の色…」で任意の色） |
| グラフ | 1行1グラフ。横軸（初期値 elapsed_ms、経過時間(s) も選べる）、第1軸（左）・第2軸（右）にそれぞれ2つまで項目を載せられる。第1軸の1つ目は必須、それ以外は「なし」も可。同じ項目は1つのグラフで1回まで。横軸が時間（elapsed_ms・経過時間(s)）のときは「横軸の単位」で ms / s / min / h を選べ、元と違う単位を選ぶと解析シートの元の列の右に換算した列（例 time(min)）が追加されて、グラフの横軸になります。0個ならグラフシートを作りません |
| 出力先 | 初期値は UART CSV のフォルダと `解析_<UART 1行目の時刻>.xlsx` |

入力にエラーがあると、その欄が赤い下線、項目の行が赤い背景になり、タブに「エラー N」のバッジが付きます。
画面下にはエラーの一覧（`[場所] メッセージ`）が出て、［Excelに出力］は押せなくなります。場所の見出しをクリックすると該当の欄へ移動します。
欠落や単位変更などの警告だけのときは「⚠ 警告 N件（出力はできます）」と表示され、出力できます。

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
| 解析 | 測定日時・測定機、UART の採用列（経過時間(s)が先頭）、1列空けてロガーの採用CH。見出し行は薄い緑で罫線付き、UART・ロガーのブロックごとに外枠、欠落補完行は灰色。シート見出しは緑 |
| グラフ | 散布図（直線・マーカーなし、シート見出しはオレンジ）。横軸はグラフごとに選択（初期値は elapsed_ms）、空欄は線を途切れさせて表示 |
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
    "voltage_mV": {"enabled": true, "label": "電圧(V)", "coef": 0.001, "color": "4472C4"},
    "t_ms":       {"enabled": false, "label": "t_ms", "coef": 1}
  },
  "logger_channels": {
    "CH1": {"enabled": true, "label": "CH1(mV)", "coef": 1, "color": "A5A5A5", "unit": "mV"}
  },
  "logger_label_format": "{ch}({unit})",
  "elapsed_label": "経過時間(s)",
  "graphs": [
    {"x": "elapsed_ms", "x_unit": "min", "primary": ["voltage_mV", "CH2"], "secondary": ["current_mA", null]},
    {"x": "@elapsed", "primary": ["soc_percent", null], "secondary": [null, null]}
  ],
  "last_dirs": {"uart": "C:\\data\\uart", "logger": "C:\\data\\logger"},
  "output_filename_format": "解析_{start:%y%m%d-%H%M%S}.xlsx",
  "row_warn_threshold": 100000,
  "row_limit": 1000000,
  "theme": "auto",
  "offset_section_open": true
}
```

| キー | 内容 |
|---|---|
| `uart_columns` | UART の列ごと（キーは CSV の元の列名）の 採用・ラベル・係数・グラフの色（`RRGGBB`） |
| `logger_channels` | ロガーの CH ごと（キーは `CH1` など）の設定と、前回の単位 `unit`。単位が前回と違うと画面2に警告が出ます |
| `logger_label_format` | 設定がない CH のラベルの書式。`{ch}` が CH 番号、`{unit}` が単位に置き換わります |
| `elapsed_label` | 経過時間(s)列のラベル |
| `graphs` | グラフの一覧。横軸 `x`・第1軸 `primary`・第2軸 `secondary` は列名・CH 番号で指定（縦軸は2つまでのリストで `null` は「なし」、`x` の `"@elapsed"` は経過時間(s)。`x` を省略すると `elapsed_ms`。`x_unit` は横軸の時間の単位 `ms`/`s`/`min`/`h`、省略すると元の単位） |
| `last_dirs` | ［参照］ダイアログで最初に開くフォルダ |
| `output_filename_format` | 出力ファイル名の書式。`{start:...}` に UART 1 行目の時刻が入ります（`%y%m%d` などは Python の strftime 書式） |
| `row_warn_threshold` / `row_limit` | 行数の警告しきい値 / 上限 |
| `theme` | 画面のテーマ。`auto`（Windows の「アプリモード」のライト／ダークに追従）、`light`、`dark` |
| `offset_section_open` | 画面2の「間隔・時間オフセット」を開いておくか（出力成功時に保存） |

- 今回の CSV にない列・CH の設定は削除されずに残ります（CH 数を一時的に減らした場合など）
- 間隔・出力先フォルダ・ファイル名は保存されません（毎回 CSV から決まります）

## アイコン

アプリのアイコン（Claude Design で作成）は `assets/` にあります。exe のアイコン（ビルド時の `--icon`）と、
ウィンドウのタイトルバー・タスクバーのアイコンに使われます。元データ（SVG・各サイズの PNG）は
`docs/claude_design/handoff/assets/icon/` にあります。

## 困ったとき

画面の操作中に想定外のエラーが起きると、メッセージを表示し、`config.json` と同じフォルダの `error.log` に詳細を記録します。
不具合の報告の際は `error.log` の内容を添えてください。

## exe のビルド手順（Windows）

`--onedir` 形式（フォルダ配布）でビルドします。`--onefile` は起動が遅くなるため使いません。

```
pip install -r requirements-dev.txt
pyinstaller --noconfirm --clean --onedir --windowed --collect-data sv_ttk --collect-all tkinterdnd2 --icon assets/LogMerger.ico --add-data "assets;assets" --name LogMerger main.py
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
| `gui.py` | GUI 本体（ウィンドウ・ヘッダー・画面切替・テーマ） |
| `gui_screen1.py` / `gui_screen2.py` | 画面1（ファイル選択）／画面2（出力内容の調整） |
| `gui_dialogs.py` | カラーパレット・欠落の詳細・出力中・完了 |
| `gui_widgets.py` / `gui_theme.py` | 共通部品（スクロール領域・タブ・アイコン）／配色・フォント・寸法の適用 |
| `design_tokens.py` | 色（ライト／ダーク）・フォント・寸法の定数（Claude Design の引き継ぎ資料） |
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
