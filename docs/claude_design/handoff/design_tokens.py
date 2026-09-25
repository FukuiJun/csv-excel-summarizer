"""LogMerger GUI のデザイン定数（IMPLEMENTATION_SPEC.md と対応）。

- 寸法はすべて 96dpi（表示倍率 100%）時の px。
  高 DPI では tk.call("tk", "scaling") に合わせて scale() で換算する。
- フォントサイズは tkinter の負の値（= ピクセル指定）。
- sv-ttk が塗る ttk 部品はテーマ任せ。ここの色は主に tk.Frame / tk.Label /
  tk.Canvas / tk.Button など sv-ttk が塗らない部品と、ttk.Style の独自スタイルに使う。
"""

# ---------------------------------------------------------------- 色
LIGHT = {
    "bg":           "#FAFAFA",  # 画面背景・ヘッダー・メイン領域
    "surface":      "#F3F3F3",  # サイドバー・フッター
    "panel":        "#FFFFFF",  # 表パネル・選択中タブ
    "tile":         "#FFFFFF",  # 読み込み結果タイル
    "border":       "#E5E5E5",  # 枠・区切り線
    "divider":      "#E0E0E0",  # サイドバー内のセクション区切り
    "row_line":     "#ECECEC",  # 表の見出し下線・グループ見出し下線
    "text":         "#1C1C1C",
    "text2":        "#5C5C5C",  # 補足・列見出し
    "disabled":     "#A0A0A0",  # 無効な入力欄の文字
    "disabled_name":"#8A8A8A",  # 採用オフ行の「元の列名」
    "accent":       "#005FB8",
    "on_accent":    "#FFFFFF",
    "swatch_border":"#BDBDBD",
    "badge_bg":     "#EFEFEF",  # タブの件数バッジ
    "info_bg":      "#EFF6FC", "info_border":    "#CFE4FA",  # 出力予定の枠
    "success":      "#0E700E", "success_bg":     "#DFF6DD",
    "warning":      "#9D5D00", "warning_bg":     "#FFF4CE", "warning_border": "#F2D88A", "warning_row": "#FFFBEA",
    "error":        "#C42B1C", "error_bg":       "#FDE7E9", "error_border":   "#F4BFC4", "error_row":   "#FDF3F4",
    "error_chip_border": "#F1B8BF",
    "drop_bg":      "#F2F7FC", "drop_border":    "#8FB3DC",  # 画面1のドロップ枠
}

DARK = {
    "bg":           "#1C1C1C",
    "surface":      "#202020",
    "panel":        "#272727",
    "tile":         "#2B2B2B",
    "border":       "#333333",
    "divider":      "#333333",
    "row_line":     "#333333",
    "text":         "#F3F3F3",
    "text2":        "#B3B3B3",
    "disabled":     "#6E6E6E",
    "disabled_name":"#7A7A7A",
    "accent":       "#4CC2FF",
    "on_accent":    "#000000",  # ダークの強調ボタンは黒文字
    "swatch_border":"#5A5A5A",
    "badge_bg":     "#383838",
    "info_bg":      "#1B2C3D", "info_border":    "#2C4A66",
    "success":      "#6CCB5F", "success_bg":     "#1F3A1D",
    "warning":      "#FCE100", "warning_bg":     "#433519", "warning_border": "#6B5A1E", "warning_row": "#2E2A1C",
    "error":        "#FF99A4", "error_bg":       "#442726", "error_border":   "#6E3A3E", "error_row":   "#2E2021",
    "error_chip_border": "#6E3A3E",
    "drop_bg":      "#1B2C3D", "drop_border":    "#2C4A66",
}

# ---------------------------------------------------------------- フォント
FONT_FAMILY = "Yu Gothic UI"          # 無い環境では "Meiryo UI"
FONT_FALLBACK = "Meiryo UI"

FONTS = {
    "title":      (FONT_FAMILY, -16, "bold"),  # ヘッダーの現在ステップ、ダイアログ見出し
    "stat":       (FONT_FAMILY, -20, "bold"),  # 読み込み結果の行数
    "drop_title": (FONT_FAMILY, -18, "bold"),  # 画面1 ドロップ枠の見出し
    "body":       (FONT_FAMILY, -14),          # 入力欄・表・ボタン・タブ
    "body_bold":  (FONT_FAMILY, -14, "bold"),  # 選択中タブ・出力予定行数・強調ボタン
    "status":     (FONT_FAMILY, -13, "bold"),  # 状態の帯（欠落なし等）、エラー一覧見出し
    "small_body": (FONT_FAMILY, -13),          # 警告帯の本文、エラー一覧の本文
    "heading":    (FONT_FAMILY, -12, "bold"),  # セクション見出し・列見出し（text2 色）
    "caption":    (FONT_FAMILY, -12),          # 補足（text2 色）
}

# ---------------------------------------------------------------- 寸法（px @100%）
WINDOW = {"init_w": 1120, "init_h": 720, "min_w": 960, "min_h": 560}

SIZE = {
    "header_h": 56, "footer_h": 64, "footer_btnbar_h": 56,  # エラー時はエラー一覧＋56px のボタン行
    "page_pad_x": 24,
    "sidebar_w": 340, "sidebar_pad_x": 20, "sidebar_pad_y": 16, "sidebar_section_gap": 16,
    "main_pad_x": 24, "main_pad_y": 16,
    "tab_h": 40,
    "btn_h": 32, "btn_pad_x": 16, "btn_min_w": 96, "btn_primary_w": 148, "btn_gap": 8, "btn_small_h": 28,
    "entry_h": 30, "entry_h_screen1": 32,
    "table_row_h": 34, "table_header_h": 30, "table_group_h": 28, "table_col_gap": 12,
    "table_pad_x": 16,
    "swatch_w": 44, "swatch_h": 24,
    "check": 18,
    "drop_h": 220, "screen1_content_w": 760, "screen1_top_pad": 40,
    "graph_block_pad": (12, 16), "graph_block_gap": 12, "graph_label_w": 150,
    "palette_swatch": 22, "palette_gap": 4,
}

# 項目の表の列幅（最後の列は残り幅＝状態メッセージ）
ITEM_COLUMNS = [
    ("採用", 36),
    ("元の列名", 150),
    ("ラベル（Excel の見出し）", 210),  # ウィンドウが広いときだけ最大 320 まで伸ばす
    ("係数", 84),
    ("色", 56),
    ("", None),
]

# 間隔・時間オフセットの 2×2 表：[見出し 76 | UART 可変 | ロガー 可変 | 単位 22]
OFFSET_GRID = {"label_w": 76, "unit_w": 22, "col_gap": 8, "row_gap": 6}

# 出力行数のしきい値（現行どおり）
ROWCOUNT_WARN = 100_000
ROWCOUNT_ERROR = 1_000_000

# エラー一覧：2 列 × 最大 3 行（6 件）＋「…ほか N 件」
ERROR_LIST_MAX = 6
ERROR_LIST_COLUMNS = 2


def scale(px: int, root) -> int:
    """96dpi 基準の px を現在の表示倍率に換算する。"""
    return round(px * float(root.tk.call("tk", "scaling")) / (96 / 72))
