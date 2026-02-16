# PowerPoint COM オートメーション: テーブル・グラフ・メディア・高度なオブジェクト

本ドキュメントは、PowerPoint COM オートメーション (win32com) を使用した高度なオブジェクト操作に関する包括的な調査レポートである。MCP サーバーの実装判断に利用される。

---

## 目次

1. [Table (テーブル操作)](#1-table-テーブル操作)
2. [Chart (グラフ操作)](#2-chart-グラフ操作)
3. [SmartArt](#3-smartart)
4. [Media (メディアオブジェクト)](#4-media-メディアオブジェクト)
5. [OLEObject](#5-oleobject)
6. [Selection / Clipboard操作](#6-selection--clipboard操作)
7. [Hyperlink / ActionSettings](#7-hyperlink--actionsettings)
8. [Animation / Transition](#8-animation--transition)
9. [Theme / Design](#9-theme--design)
10. [MCP サーバー実装への推奨事項](#10-mcp-サーバー実装への推奨事項)

---

## 1. Table (テーブル操作)

### 1.1 テーブルの作成 (Shapes.AddTable)

**メソッド:** `Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height)`

| パラメータ | 必須/任意 | 型 | 説明 |
|---|---|---|---|
| NumRows | 必須 | Long | テーブルの行数 |
| NumColumns | 必須 | Long | テーブルの列数 |
| Left | 任意 | Single | スライド左端からの距離 (ポイント) |
| Top | 任意 | Single | スライド上端からの距離 (ポイント) |
| Width | 任意 | Single | テーブルの幅 (ポイント) |
| Height | 任意 | Single | テーブルの高さ (ポイント) |

**戻り値:** Shape オブジェクト

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
prs = app.Presentations.Add()
slide = prs.Slides.Add(1, 12)  # ppLayoutBlank = 12

# 3行4列のテーブルを作成 (位置: 50,100, サイズ: 600x300 ポイント)
shape = slide.Shapes.AddTable(NumRows=3, NumColumns=4, Left=50, Top=100, Width=600, Height=300)
table = shape.Table
```

**注意点:**
- 戻り値は Shape オブジェクトであり、Table オブジェクトではない。`shape.Table` で Table オブジェクトにアクセスする。
- `shape.HasTable` プロパティで Shape がテーブルを含むか確認可能。
- Width/Height を省略するとデフォルトサイズが適用される。

### 1.2 Rows / Columns コレクション

**Table.Rows** と **Table.Columns** でそれぞれ行と列のコレクションにアクセスできる。

```python
table = shape.Table

# 行数・列数の取得
num_rows = table.Rows.Count
num_cols = table.Columns.Count

# 行・列の追加
table.Rows.Add()              # 末尾に行を追加
table.Rows.Add(BeforeRow=1)   # 指定位置の前に行を追加
table.Columns.Add()           # 末尾に列を追加
table.Columns.Add(BeforeColumn=2)  # 指定位置の前に列を追加

# 行・列の削除
table.Rows(3).Delete()       # 3行目を削除
table.Columns(2).Delete()    # 2列目を削除
```

**注意点:**
- `Rows.Add()` / `Columns.Add()` の引数は挿入位置。省略すると末尾に追加される。
- PowerPoint では「セルの追加・削除」は不可。行または列単位での追加・削除のみ可能。

### 1.3 Cell(row, col) アクセス

**メソッド:** `Table.Cell(Row, Column)` - Cell オブジェクトを返す

```python
table = shape.Table

# セルへのアクセス (1-based インデックス)
cell = table.Cell(1, 1)  # 1行1列目のセル

# 全セルの巡回
for row in range(1, table.Rows.Count + 1):
    for col in range(1, table.Columns.Count + 1):
        cell = table.Cell(row, col)
        # セル操作...
```

**注意点:**
- インデックスは **1-based** (VBA と同じ)。
- 結合されたセルにアクセスする場合、結合範囲内のどのセル座標でもアクセス可能だが、同じ Cell オブジェクトが返される。

### 1.4 セルの結合 (Merge) / 分割 (Split)

#### Merge

**メソッド:** `Cell.Merge(MergeTo)`

```python
table = shape.Table

# セル(1,1) と セル(1,2) を結合
table.Cell(1, 1).Merge(table.Cell(1, 2))

# 複数行・列にまたがる結合 (1,1) から (2,3) まで
table.Cell(1, 1).Merge(table.Cell(2, 3))
```

#### Split

**メソッド:** `Cell.Split(NumRows, NumColumns)`

```python
# セル(1,1) を 2行1列に分割 (縦に分割)
table.Cell(1, 1).Split(NumRows=2, NumColumns=1)

# セル(1,1) を 1行3列に分割 (横に分割)
table.Cell(1, 1).Split(NumRows=1, NumColumns=3)
```

**注意点:**
- `Merge` の引数 `MergeTo` は結合先のセル (Cell オブジェクト)。
- 結合は左上から右下方向へ指定する。
- 結合セルの Split は元の結合前の状態に関わらず、指定した行数・列数に分割される。
- 複雑な結合/分割を繰り返すとテーブル構造が不安定になることがある。

### 1.5 セルのテキスト設定

**プロパティ:** `Cell.Shape.TextFrame.TextRange`

```python
table = shape.Table

# テキストの設定
table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "ヘッダー1"
table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "ヘッダー2"

# テキストの取得
text = table.Cell(1, 1).Shape.TextFrame.TextRange.Text

# テキスト書式の設定
text_range = table.Cell(1, 1).Shape.TextFrame.TextRange
text_range.Font.Size = 14
text_range.Font.Bold = True
text_range.Font.Color.RGB = 0xFFFFFF  # 白色 (BGR形式: 0xBBGGRR)
text_range.ParagraphFormat.Alignment = 2  # ppAlignCenter = 2

# TextFrame のマージン設定
tf = table.Cell(1, 1).Shape.TextFrame
tf.MarginLeft = 5
tf.MarginRight = 5
tf.MarginTop = 3
tf.MarginBottom = 3
tf.WordWrap = True  # msoTrue
```

**注意点:**
- Cell オブジェクトには直接 TextFrame がない。`Cell.Shape.TextFrame` 経由でアクセスする。
- RGB 値は BGR 形式 (Blue, Green, Red) で指定する。赤 = 0x0000FF, 青 = 0xFF0000。
- `TextFrame2` も使用可能で、より高度な書式設定が可能。

### 1.6 セルの書式設定 (Borders, Fill)

#### Borders (罫線)

**定数:**
- `ppBorderTop = 1` - 上罫線
- `ppBorderLeft = 2` - 左罫線
- `ppBorderBottom = 3` - 下罫線
- `ppBorderRight = 4` - 右罫線
- `ppBorderDiagonalDown = 5` - 左上から右下の対角線
- `ppBorderDiagonalUp = 6` - 左下から右上の対角線

```python
from win32com.client import constants as c

table = shape.Table
cell = table.Cell(1, 1)

# 罫線の設定
border = cell.Borders(3)  # ppBorderBottom = 3
border.Visible = True     # msoTrue = -1
border.Weight = 2.0       # 線の太さ (ポイント)
border.ForeColor.RGB = 0x000000  # 黒色

# DashStyle の設定
# border.DashStyle = 1  # msoLineSolid = 1

# 全罫線を設定するユーティリティ
for border_type in range(1, 5):  # 上・左・下・右
    border = cell.Borders(border_type)
    border.Visible = True
    border.Weight = 1.0
    border.ForeColor.RGB = 0x808080  # グレー
```

#### Fill (背景色)

```python
cell = table.Cell(1, 1)

# 塗りつぶし色の設定
cell.Shape.Fill.ForeColor.RGB = 0x993300  # 暗い青色
cell.Shape.Fill.Visible = True  # msoTrue

# 透過度の設定
cell.Shape.Fill.Transparency = 0.3  # 30% 透過

# グラデーション
cell.Shape.Fill.TwoColorGradient(1, 1)  # msoGradientHorizontal, variant 1
cell.Shape.Fill.ForeColor.RGB = 0xFF0000
cell.Shape.Fill.BackColor.RGB = 0x0000FF
```

**注意点:**
- Borders は Cell オブジェクトから直接アクセスする (`Cell.Borders`)。
- Fill は `Cell.Shape.Fill` 経由。
- 罫線のスタイル (DashStyle) は全ての線種がサポートされているわけではない。

### 1.7 行の高さ、列の幅

```python
table = shape.Table

# 行の高さを設定 (ポイント単位)
table.Rows(1).Height = 40  # 1行目の高さ
table.Rows(2).Height = 30  # 2行目の高さ

# 列の幅を設定 (ポイント単位)
table.Columns(1).Width = 150  # 1列目の幅
table.Columns(2).Width = 100  # 2列目の幅

# 全行の高さを統一
for i in range(1, table.Rows.Count + 1):
    table.Rows(i).Height = 30

# 全列の幅を統一
for i in range(1, table.Columns.Count + 1):
    table.Columns(i).Width = 120

# テーブル全体の比率を保ったスケーリング
table.ScaleProportionally(1.5)  # 1.5倍にスケール
```

**注意点:**
- 高さ・幅はポイント単位 (1インチ = 72ポイント)。
- `ScaleProportionally` メソッドで比率を維持したままサイズ変更が可能。

### 1.8 TableStyle (テーブルスタイル)

**メソッド:** `Table.ApplyStyle(StyleID, SaveFormatting)`

| パラメータ | 必須/任意 | 型 | 説明 |
|---|---|---|---|
| StyleID | 任意 | String | テーブルスタイルの GUID |
| SaveFormatting | 任意 | Boolean | True で既存の書式を保持 |

**プロパティ:** `Table.Style` - 現在の TableStyle オブジェクトを返す (読み取り専用)

```python
table = shape.Table

# テーブルスタイルの適用 (GUID で指定)
# Medium Style 2 - Accent 1 の例:
table.ApplyStyle("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")

# 書式を保持しながらスタイルを適用
table.ApplyStyle("{2D5ABB26-0587-4C30-8999-92F81FD0307C}", True)

# 現在のスタイル ID を取得
current_style = table.Style  # TableStyle オブジェクト
```

**主要なテーブルスタイル GUID 一覧:**

| スタイル名 | GUID |
|---|---|
| No Style, No Grid | {2D5ABB26-0587-4C30-8999-92F81FD0307C} |
| Themed Style 1 - Accent 1 | {3C2FFA5D-87B4-456A-9821-1D502468CF0F} |
| Medium Style 2 - Accent 1 | {5C22544A-7EE6-4342-B048-85BDC9FD1C3A} |
| Light Style 1 | {9D7B26C5-4107-4FEC-AEDC-1716B250A1EF} |
| Light Style 1 - Accent 1 | {3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5} |
| Light Style 2 | {0E3FDE45-AF77-4B5C-9715-49D594BDF05E} |
| Light Style 3 | {C083E6E3-FA7D-4D7B-A595-EF9225AFEA82} |
| Dark Style 1 | {E8034E78-7F5D-4C2E-B375-FC64B27BC917} |
| Dark Style 2 | {125E5076-3810-47DD-B79F-674D7AD40C01} |

**注意点:**
- スタイル GUID はドキュメント化が不十分。Open XML SDK 等で確認する必要がある。
- `SaveFormatting = True` を指定すると、個別のセル書式が維持される。

### 1.9 FirstRow, LastRow, FirstCol, LastCol のバンド設定

```python
table = shape.Table

# 先頭行の特別書式を有効化 (ヘッダー行)
table.FirstRow = True   # msoTrue = -1

# 最終行の特別書式を有効化 (合計行)
table.LastRow = True

# 先頭列の特別書式を有効化
table.FirstCol = True

# 最終列の特別書式を有効化
table.LastCol = True

# 縞模様 (バンド) の設定
table.HorizBanding = True   # 行方向の縞模様 (偶数行と奇数行で書式が異なる)
table.VertBanding = False    # 列方向の縞模様を無効化
```

**Table オブジェクトの全プロパティ一覧:**

| プロパティ | 型 | 説明 |
|---|---|---|
| AlternativeText | String | 代替テキスト |
| Background | TableBackground | テーブルの背景 |
| Columns | Columns | 列コレクション |
| FirstCol | MsoTriState | 先頭列の特別書式 |
| FirstRow | MsoTriState | 先頭行の特別書式 |
| HorizBanding | MsoTriState | 行方向の縞模様 |
| LastCol | MsoTriState | 最終列の特別書式 |
| LastRow | MsoTriState | 最終行の特別書式 |
| Rows | Rows | 行コレクション |
| Style | TableStyle | テーブルスタイル (読み取り専用) |
| TableDirection | PpDirection | テーブルの方向 (LTR/RTL) |
| Title | String | テーブルのタイトル |
| VertBanding | MsoTriState | 列方向の縞模様 |

**注意点:**
- これらのプロパティは読み書き可能 (MsoTriState 型)。
- テーブルスタイルが適用されている場合にのみ視覚的効果がある。
- `MsoTriState` は `True = -1 (msoTrue)`, `False = 0 (msoFalse)` である。

---

## 2. Chart (グラフ操作)

### 2.1 グラフの作成と種類

**メソッド:** `Shapes.AddChart2(Style, Type, Left, Top, Width, Height, NewLayout)`

| パラメータ | 必須/任意 | 型 | 説明 |
|---|---|---|---|
| Style | 任意 | Long | チャートスタイル (-1 でデフォルト) |
| Type | 任意 | XlChartType | チャートの種類 |
| Left | 任意 | Single | 左端位置 (ポイント) |
| Top | 任意 | Single | 上端位置 (ポイント) |
| Width | 任意 | Single | 幅 (ポイント) |
| Height | 任意 | Single | 高さ (ポイント) |
| NewLayout | 任意 | Boolean | True で動的フォーマットルール使用 |

**主要なチャート種類定数 (XlChartType):**

| 定数名 | 値 | 説明 |
|---|---|---|
| xlColumnClustered | 51 | 集合縦棒 |
| xlColumnStacked | 52 | 積み上げ縦棒 |
| xlColumnStacked100 | 53 | 100% 積み上げ縦棒 |
| xlBarClustered | 57 | 集合横棒 |
| xlBarStacked | 58 | 積み上げ横棒 |
| xlLine | 4 | 折れ線 |
| xlLineMarkers | 65 | マーカー付き折れ線 |
| xlLineStacked | 63 | 積み上げ折れ線 |
| xlPie | 5 | 円グラフ |
| xlPieExploded | 69 | 分割円グラフ |
| xlDoughnut | -4120 | ドーナツ |
| xlArea | 1 | 面グラフ |
| xlAreaStacked | 76 | 積み上げ面 |
| xlXYScatter | -4169 | 散布図 |
| xlXYScatterLines | 74 | 折れ線付き散布図 |
| xlRadar | -4151 | レーダー |
| xlBubble | 15 | バブル |
| xlStockHLC | 88 | 高値-安値-終値 |
| xl3DColumnClustered | 54 | 3D 集合縦棒 |
| xl3DPie | -4102 | 3D 円グラフ |
| xl3DLine | -4101 | 3D 折れ線 |

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
prs = app.Presentations.Add()
slide = prs.Slides.Add(1, 12)  # ppLayoutBlank

# 集合縦棒グラフを作成
chart_shape = slide.Shapes.AddChart2(
    Style=-1,           # デフォルトスタイル
    Type=51,            # xlColumnClustered
    Left=50,
    Top=50,
    Width=500,
    Height=350,
    NewLayout=True
)
chart = chart_shape.Chart
```

**注意点:**
- `AddChart2` は Office 2013 以降で利用可能。旧バージョンには `AddChart` がある。
- Style = -1 はデフォルトスタイル。具体的なスタイル番号はチャート種類により異なる。
- グラフ作成時にExcel が内部的に起動される (背景で Workbook が作られる)。

### 2.2 ChartData (データの設定 - Excel Workbook 経由)

**プロパティ:** `Chart.ChartData` - ChartData オブジェクトを返す

**メソッド:** `ChartData.Activate()` - チャートデータの Workbook をアクティブ化

**プロパティ:** `ChartData.Workbook` - 内部の Excel Workbook オブジェクト

```python
chart = chart_shape.Chart

# チャートデータを操作するためにアクティブ化
chart.ChartData.Activate()
wb = chart.ChartData.Workbook
ws = wb.Worksheets(1)

# データの設定
# カテゴリ (横軸)
ws.Range("A2").Value = "1月"
ws.Range("A3").Value = "2月"
ws.Range("A4").Value = "3月"
ws.Range("A5").Value = "4月"

# 系列1のラベルとデータ
ws.Range("B1").Value = "売上"
ws.Range("B2").Value = 1000
ws.Range("B3").Value = 1500
ws.Range("B4").Value = 1200
ws.Range("B5").Value = 1800

# 系列2のラベルとデータ
ws.Range("C1").Value = "利益"
ws.Range("C2").Value = 300
ws.Range("C3").Value = 450
ws.Range("C4").Value = 360
ws.Range("C5").Value = 540

# データ範囲の設定
chart.SetSourceData(ws.Range("A1:C5"))

# Workbook を閉じる (重要)
wb.Close(SaveChanges=False)
```

**注意点:**
- `ChartData.Activate()` を呼ばないと Workbook にアクセスできない。
- Excel が起動されるため、処理が重くなることがある。
- 操作完了後は `wb.Close()` で閉じることが推奨される。閉じないとプロセスが残る可能性がある。
- `ChartData.ActivateChartDataWindow()` は新しいバージョンで追加された代替メソッド。

### 2.3 ChartTitle, ChartArea, PlotArea

```python
chart = chart_shape.Chart

# チャートタイトル
chart.HasTitle = True
chart.ChartTitle.Text = "月次売上レポート"
chart.ChartTitle.Font.Size = 16
chart.ChartTitle.Font.Bold = True
chart.ChartTitle.Font.Color.RGB = 0x993300  # 暗い青色

# チャートエリア (グラフ全体の領域)
chart_area = chart.ChartArea
chart_area.Format.Fill.ForeColor.RGB = 0xFFF8F0  # 背景色
chart_area.Border.LineStyle = 1  # xlContinuous
chart_area.Border.Color = 0x808080

# プロットエリア (データが描画される領域)
plot_area = chart.PlotArea
plot_area.Format.Fill.ForeColor.RGB = 0xFFFFFF  # 白色
plot_area.Left = 60
plot_area.Top = 40
plot_area.Width = 400
plot_area.Height = 250
```

### 2.4 Series コレクション (データ系列)

```python
chart = chart_shape.Chart

# 系列コレクションへのアクセス
series_count = chart.SeriesCollection().Count

# 各系列の操作
for i in range(1, series_count + 1):
    series = chart.SeriesCollection(i)
    print(f"系列 {i}: {series.Name}")

# 系列の書式設定
series1 = chart.SeriesCollection(1)
series1.Format.Fill.ForeColor.RGB = 0xFF6633  # 系列1の色
series1.Format.Line.ForeColor.RGB = 0xFF6633

# 折れ線グラフの場合
# series1.Format.Line.Weight = 2.5
# series1.MarkerStyle = 8  # xlMarkerStyleCircle
# series1.MarkerSize = 8

# 系列の追加 (ChartData 経由でデータを追加した後)
# chart.SeriesCollection().NewSeries()
# new_series = chart.SeriesCollection(chart.SeriesCollection().Count)
# new_series.Name = "新系列"
# new_series.Values = [100, 200, 300, 400]
```

### 2.5 Axes (軸の設定)

**メソッド:** `Chart.Axes(Type, AxisGroup)`

| 定数 | 値 | 説明 |
|---|---|---|
| xlCategory | 1 | カテゴリ軸 (X 軸) |
| xlValue | 2 | 数値軸 (Y 軸) |
| xlSeriesAxis | 3 | 系列軸 (3D グラフのみ) |

```python
chart = chart_shape.Chart

# カテゴリ軸 (X 軸)
cat_axis = chart.Axes(1)  # xlCategory = 1
cat_axis.HasTitle = True
cat_axis.AxisTitle.Text = "月"
cat_axis.TickLabels.Font.Size = 10
cat_axis.TickLabels.Orientation = 0  # 水平

# 数値軸 (Y 軸)
val_axis = chart.Axes(2)  # xlValue = 2
val_axis.HasTitle = True
val_axis.AxisTitle.Text = "金額 (万円)"
val_axis.MinimumScale = 0
val_axis.MaximumScale = 2000
val_axis.MajorUnit = 500

# 目盛線
val_axis.HasMajorGridlines = True
val_axis.HasMinorGridlines = False
val_axis.MajorGridlines.Format.Line.ForeColor.RGB = 0xD0D0D0

# 数値書式
val_axis.TickLabels.NumberFormat = "#,##0"
```

### 2.6 Legend (凡例)

```python
chart = chart_shape.Chart

# 凡例の表示
chart.HasLegend = True
legend = chart.Legend

# 凡例の位置
legend.Position = -4107  # xlLegendPositionBottom
# xlLegendPositionBottom = -4107
# xlLegendPositionLeft = -4131
# xlLegendPositionRight = -4152
# xlLegendPositionTop = -4160
# xlLegendPositionCorner = 2

# 凡例のフォント
legend.Font.Size = 10

# 凡例エントリの操作
for i in range(1, legend.LegendEntries().Count + 1):
    entry = legend.LegendEntries(i)
    # entry.Delete()  # エントリの削除
```

### 2.7 DataLabels (データラベル)

```python
chart = chart_shape.Chart
series = chart.SeriesCollection(1)

# データラベルの有効化
series.HasDataLabels = True
labels = series.DataLabels()

# データラベルの書式
labels.ShowValue = True            # 値を表示
labels.ShowCategoryName = False    # カテゴリ名を非表示
labels.ShowSeriesName = False      # 系列名を非表示
labels.ShowPercentage = False      # パーセントを非表示 (円グラフ用)
labels.ShowLegendKey = False       # 凡例マーカーを非表示
labels.Font.Size = 9
labels.Font.Color.RGB = 0x333333
labels.NumberFormat = "#,##0"

# 個別のデータラベル
point = series.Points(1)
point.HasDataLabel = True
point.DataLabel.Text = "特記: 1000万"
```

### 2.8 ChartStyle

```python
chart = chart_shape.Chart

# チャートスタイルの適用 (番号で指定)
chart.ChartStyle = 2  # スタイル番号 (1-48等)

# グラフ種類の変更
chart.ChartType = 4    # xlLine (折れ線に変更)
chart.ChartType = 5    # xlPie (円グラフに変更)
chart.ChartType = 51   # xlColumnClustered (集合縦棒に戻す)
```

### 2.9 グラフ要素の書式設定

```python
chart = chart_shape.Chart

# 系列の色
series = chart.SeriesCollection(1)
series.Format.Fill.ForeColor.RGB = 0xFF6633

# 特定のデータポイントの色
point = series.Points(1)
point.Format.Fill.ForeColor.RGB = 0x0000FF  # 赤

# プロットエリアの書式
chart.PlotArea.Format.Fill.ForeColor.RGB = 0xF5F5F5

# 境界線
chart.ChartArea.Format.Line.Visible = True
chart.ChartArea.Format.Line.ForeColor.RGB = 0x808080
chart.ChartArea.Format.Line.Weight = 1.0
```

**グラフ操作の注意点:**
- PowerPoint のグラフは内部的に Excel を使用する。COM 操作中に Excel プロセスが残る場合がある。
- `ChartData.Activate()` 後は必ず `Workbook.Close()` を呼ぶこと。
- グラフの種類を変更すると、データ構造が崩れる場合がある (例: 円グラフは通常1系列のみ)。
- 大量のデータポイントを持つグラフの操作は処理が重くなる。
- `win32com.client.constants` でアクセスできない定数がある場合は、数値を直接使用すること。

---

## 3. SmartArt

### 3.1 SmartArt の作成 (AddSmartArt)

**メソッド:** `Shapes.AddSmartArt(Layout, Left, Top, Width, Height)`

| パラメータ | 必須/任意 | 型 | 説明 |
|---|---|---|---|
| Layout | 必須 | SmartArtLayout | SmartArt レイアウト |
| Left | 任意 | Single | 左端位置 (ポイント) |
| Top | 任意 | Single | 上端位置 (ポイント) |
| Width | 任意 | Single | 幅 (ポイント) |
| Height | 任意 | Single | 高さ (ポイント) |

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
prs = app.Presentations.Add()
slide = prs.Slides.Add(1, 12)  # ppLayoutBlank

# SmartArt レイアウトの取得
# Application.SmartArtLayouts コレクションからレイアウトを選択
layouts = app.SmartArtLayouts

# レイアウトの列挙
for i in range(1, layouts.Count + 1):
    layout = layouts(i)
    print(f"{i}: {layout.Name} - {layout.Description}")

# 特定のレイアウトで SmartArt を追加
# 例: Basic Block List (インデックスはバージョンにより異なる)
smart_art_shape = slide.Shapes.AddSmartArt(
    Layout=app.SmartArtLayouts(1),  # 最初のレイアウト
    Left=50,
    Top=50,
    Width=600,
    Height=400
)
```

### 3.2 SmartArtLayout (レイアウトの種類)

SmartArt のレイアウトは `Application.SmartArtLayouts` コレクションから取得する。

```python
# レイアウトの一覧を取得
layouts = app.SmartArtLayouts
for i in range(1, min(layouts.Count + 1, 20)):  # 最初の20個
    layout = layouts(i)
    print(f"Index {i}: {layout.Name}")

# レイアウトの変更
smart_art = smart_art_shape.SmartArt
smart_art.Layout = app.SmartArtLayouts(5)  # 別のレイアウトに変更
```

**主要なレイアウトカテゴリ:**
- List (リスト): Basic Block List, Lined List, etc.
- Process (プロセス): Basic Process, Step Down Process, etc.
- Cycle (循環): Basic Cycle, Text Cycle, etc.
- Hierarchy (階層): Organization Chart, Hierarchy, etc.
- Relationship (関係): Balance, Funnel, etc.
- Matrix (マトリックス): Basic Matrix, etc.
- Pyramid (ピラミッド): Basic Pyramid, etc.

**注意点:**
- レイアウトのインデックスは PowerPoint のバージョンやインストールされたアドインにより異なる場合がある。
- Name プロパティでレイアウトを特定する方が安全。

### 3.3 SmartArtNodes (ノードの操作)

```python
smart_art = smart_art_shape.SmartArt

# 全ノードへのアクセス
all_nodes = smart_art.AllNodes
print(f"ノード数: {all_nodes.Count}")

# トップレベルのノード
nodes = smart_art.Nodes
print(f"トップレベルノード数: {nodes.Count}")

# ノードの追加
new_node = smart_art.AllNodes.Add()  # 末尾にノード追加

# ノードのテキスト設定
for i in range(1, all_nodes.Count + 1):
    node = all_nodes(i)
    node.TextFrame2.TextRange.Text = f"項目 {i}"

# 特定のノードにテキスト設定
smart_art.AllNodes(1).TextFrame2.TextRange.Text = "第1ステップ"
smart_art.AllNodes(2).TextFrame2.TextRange.Text = "第2ステップ"
smart_art.AllNodes(3).TextFrame2.TextRange.Text = "第3ステップ"

# ノードの削除
smart_art.AllNodes(3).Delete()

# 子ノードの追加 (階層型レイアウトの場合)
parent_node = smart_art.Nodes(1)
child_node = parent_node.Nodes.Add()
child_node.TextFrame2.TextRange.Text = "子項目"
```

### 3.4 SmartArtNode.TextFrame2 (テキスト設定)

```python
node = smart_art.AllNodes(1)

# テキストの設定
node.TextFrame2.TextRange.Text = "メインテキスト"

# テキスト書式
node.TextFrame2.TextRange.Font.Size = 14
node.TextFrame2.TextRange.Font.Bold = True
node.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x000000
```

### 3.5 Color scheme 変更

```python
smart_art = smart_art_shape.SmartArt

# カラースタイルの変更
# SmartArt.Color プロパティで SmartArtColor を設定
# app.SmartArtColors コレクションから色を選択
colors = app.SmartArtColors
for i in range(1, min(colors.Count + 1, 10)):
    print(f"Color {i}: {colors(i).Name}")

smart_art.Color = app.SmartArtColors(2)  # 別の配色に変更

# クイックスタイルの変更
styles = app.SmartArtQuickStyles
smart_art.QuickStyle = styles(3)
```

**SmartArt 操作の注意点:**
- SmartArt は内部的に複雑な XML 構造を持つ。COM 経由の操作は限定的。
- レイアウトを変更するとノード構造が再配置される。
- `AllNodes` はフラットなリスト、`Nodes` は階層的なアクセスを提供。
- レイアウトの種類によって利用可能なノード操作が異なる。
- SmartArt の個々のシェイプ要素を直接操作するのは推奨されない (レイアウトエンジンが管理するため)。

---

## 4. Media (メディアオブジェクト)

### 4.1 動画の挿入 (AddMediaObject2)

**メソッド:** `Shapes.AddMediaObject2(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)`

| パラメータ | 必須/任意 | 型 | 説明 |
|---|---|---|---|
| FileName | 必須 | String | メディアファイルのパス |
| LinkToFile | 任意 | MsoTriState | ファイルにリンクするか |
| SaveWithDocument | 任意 | MsoTriState | ドキュメントに保存するか |
| Left | 任意 | Single | 左端位置 (ポイント) |
| Top | 任意 | Single | 上端位置 (ポイント) |
| Width | 任意 | Single | 幅 (ポイント、-1でデフォルト) |
| Height | 任意 | Single | 高さ (ポイント、-1でデフォルト) |

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
prs = app.ActivePresentation
slide = prs.Slides(1)

# 動画の挿入 (埋め込み)
video_shape = slide.Shapes.AddMediaObject2(
    FileName=r"C:\Videos\sample.mp4",
    LinkToFile=False,       # msoFalse = 0
    SaveWithDocument=True,  # msoTrue = -1
    Left=100,
    Top=100,
    Width=480,
    Height=270
)

# 動画の挿入 (リンク)
video_linked = slide.Shapes.AddMediaObject2(
    FileName=r"C:\Videos\sample.mp4",
    LinkToFile=True,        # msoTrue = -1
    SaveWithDocument=False,
    Left=100,
    Top=400,
    Width=320,
    Height=180
)
```

### 4.2 音声の挿入

```python
# 音声ファイルの挿入 (埋め込み)
audio_shape = slide.Shapes.AddMediaObject2(
    FileName=r"C:\Audio\bgm.mp3",
    LinkToFile=False,
    SaveWithDocument=True,
    Left=50,
    Top=50,
    Width=-1,    # デフォルトサイズ
    Height=-1
)
```

**注意点:**
- `LinkToFile` と `SaveWithDocument` の両方を `False` にするとエラーが発生する。少なくとも一方は `True` にする必要がある。
- 旧メソッド `AddMediaObject` は非推奨。`AddMediaObject2` を使用すること。
- サポートされるフォーマット: MP4, WMV, AVI, MP3, WAV, WMA, M4A 等。
- Width/Height に -1 を指定するとデフォルトサイズが適用される。

### 4.3 MediaFormat プロパティ

**プロパティ:** `Shape.MediaFormat` - MediaFormat オブジェクトを返す

**MediaFormat オブジェクトのプロパティ:**

| プロパティ | 型 | 説明 |
|---|---|---|
| AudioCompressionType | String | オーディオ圧縮タイプ |
| AudioSamplingRate | Long | オーディオサンプリングレート |
| EndPoint | Long | メディアの終了位置 (ミリ秒) |
| FadeInDuration | Long | フェードイン時間 (ミリ秒) |
| FadeOutDuration | Long | フェードアウト時間 (ミリ秒) |
| IsEmbedded | Boolean | 埋め込みかどうか |
| IsLinked | Boolean | リンクかどうか |
| Length | Long | メディアの長さ (ミリ秒) |
| MediaType | PpMediaType | メディアの種類 |
| Muted | Boolean | ミュート状態 |
| SampleHeight | Long | サンプルの高さ |
| SampleWidth | Long | サンプルの幅 |
| StartPoint | Long | メディアの開始位置 (ミリ秒) |
| VideoCompressionType | String | ビデオ圧縮タイプ |
| VideoFrameRate | Long | ビデオフレームレート |
| Volume | Single | ボリューム (0.0 - 1.0) |

**メソッド:**

| メソッド | 説明 |
|---|---|
| Resample() | メディアの再サンプリング |
| ResampleFromProfile() | プロファイルに基づく再サンプリング |
| SetDisplayPicture(FilePath) | 表示画像の設定 |
| SetDisplayPictureFromFile(FilePath) | ファイルからの表示画像設定 |

### 4.4 トリミング (StartPoint, EndPoint)

```python
media = video_shape.MediaFormat

# トリム範囲の設定 (ミリ秒単位)
media.StartPoint = 5000    # 5秒後から開始
media.EndPoint = 30000     # 30秒で終了

# メディアの長さを取得
total_length = media.Length  # ミリ秒
print(f"メディア長さ: {total_length / 1000} 秒")
```

### 4.5 Volume, Mute

```python
media = video_shape.MediaFormat

# ボリュームの設定 (0.0 - 1.0)
media.Volume = 0.5   # 50% のボリューム
media.Volume = 1.0   # 最大ボリューム
media.Volume = 0.0   # 無音

# ミュートの切り替え
media.Muted = True    # ミュート
media.Muted = False   # ミュート解除
```

### 4.6 FadeIn / FadeOut

```python
media = video_shape.MediaFormat

# フェードイン・アウトの設定 (ミリ秒)
media.FadeInDuration = 2000   # 2秒のフェードイン
media.FadeOutDuration = 3000  # 3秒のフェードアウト
```

### 4.7 PlaySettings (再生設定)

**プロパティ:** `Shape.AnimationSettings.PlaySettings` - PlaySettings オブジェクトを返す

**PlaySettings オブジェクトのプロパティ:**

| プロパティ | 型 | 説明 |
|---|---|---|
| HideWhileNotPlaying | MsoTriState | 再生中以外は非表示 |
| LoopUntilStopped | MsoTriState | 停止するまでループ |
| PauseAnimation | MsoTriState | 再生中アニメーション一時停止 |
| PlayOnEntry | MsoTriState | スライド表示時に自動再生 |
| RewindMovie | MsoTriState | 再生後に巻き戻し |
| StopAfterSlides | Long | 指定スライド数後に停止 |

```python
# PlaySettings へのアクセス
play = video_shape.AnimationSettings.PlaySettings

# 自動再生の設定
play.PlayOnEntry = True    # msoTrue = -1

# ループ再生
play.LoopUntilStopped = True

# 再生していない間は非表示
play.HideWhileNotPlaying = True

# アニメーションとの連携
play.PauseAnimation = True  # メディア再生中は他のアニメーションを一時停止

# 巻き戻し
play.RewindMovie = True     # 再生後に先頭に戻す
```

**メディア操作の注意点:**
- `PlaySettings` は `AnimationSettings.PlaySettings` 経由でアクセスする。
- `PlayOnEntry` を `True` にするには、`AnimationSettings.Animate` も `True` にする必要がある場合がある。
- メディアファイルのフォーマットによっては一部のプロパティが機能しない場合がある。
- 大きなメディアファイルを埋め込む場合、プレゼンテーションファイルのサイズが大幅に増加する。
- リンク形式の場合、ファイルパスが無効になるとメディアが再生できなくなる。

---

## 5. OLEObject

### 5.1 OLE オブジェクトの挿入

**メソッド:** `Shapes.AddOLEObject(Left, Top, Width, Height, ClassName, FileName, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link)`

| パラメータ | 必須/任意 | 型 | 説明 |
|---|---|---|---|
| Left | 任意 | Single | 左端位置 |
| Top | 任意 | Single | 上端位置 |
| Width | 任意 | Single | 幅 |
| Height | 任意 | Single | 高さ |
| ClassName | 任意 | String | OLE クラス名 / ProgID |
| FileName | 任意 | String | ファイルパス |
| DisplayAsIcon | 任意 | MsoTriState | アイコンとして表示 |
| IconFileName | 任意 | String | アイコンファイルパス |
| IconIndex | 任意 | Long | アイコンインデックス |
| IconLabel | 任意 | String | アイコン下のラベル |
| Link | 任意 | MsoTriState | リンクするか |

```python
slide = prs.Slides(1)

# Excel ワークシートの埋め込み (新規作成)
ole_shape = slide.Shapes.AddOLEObject(
    Left=100, Top=100,
    Width=400, Height=300,
    ClassName="Excel.Sheet"
)

# Excel ファイルの埋め込み (既存ファイル)
ole_shape2 = slide.Shapes.AddOLEObject(
    Left=100, Top=100,
    Width=400, Height=300,
    FileName=r"C:\Data\report.xlsx"
)

# Word ドキュメントのリンク
ole_linked = slide.Shapes.AddOLEObject(
    Left=100, Top=100,
    Width=200, Height=300,
    FileName=r"C:\Docs\testing.doc",
    Link=True  # msoTrue
)

# アイコンとして表示
ole_icon = slide.Shapes.AddOLEObject(
    Left=100, Top=100,
    Width=200, Height=300,
    ClassName="Excel.Sheet",
    DisplayAsIcon=True
)
```

### 5.2 リンク vs 埋め込み

| 特徴 | 埋め込み (Embedded) | リンク (Linked) |
|---|---|---|
| ファイルサイズ | 大きい (データを内包) | 小さい (参照のみ) |
| ソースファイルとの関係 | 独立 | 依存 |
| 更新 | 手動で編集 | ソース変更時に自動/手動更新 |
| 可搬性 | 高い | 低い (パスが必要) |
| Shape.Type | msoEmbeddedOLEObject (7) | msoLinkedOLEObject (10) |

```python
# オブジェクトの種類を判定
shape = slide.Shapes(1)
if shape.Type == 7:   # msoEmbeddedOLEObject
    print("埋め込みOLEオブジェクト")
elif shape.Type == 10:  # msoLinkedOLEObject
    print("リンクOLEオブジェクト")
```

### 5.3 OLEFormat.Activate

```python
ole_shape = slide.Shapes(1)  # OLE オブジェクトを含む Shape

# OLE オブジェクトのアクティブ化 (編集モードに入る)
ole_shape.OLEFormat.Activate()
```

**注意点:**
- `Activate()` により OLE サーバー (例: Excel) が起動し、オブジェクトが編集可能になる。
- ユーザーが見える状態 (Visible = True) でないと正常に動作しない場合がある。

### 5.4 OLEFormat.Object (内部オブジェクトへのアクセス)

```python
ole_shape = slide.Shapes(1)

# OLE オブジェクトの内部オブジェクトにアクセス
ole_object = ole_shape.OLEFormat.Object

# ProgID の確認
prog_id = ole_shape.OLEFormat.ProgID
print(f"ProgID: {prog_id}")  # 例: "Excel.Sheet.12"

# Excel ワークシートの場合
if "Excel" in prog_id:
    # Activate してから操作
    ole_shape.OLEFormat.Activate()
    wb = ole_shape.OLEFormat.Object  # Workbook オブジェクト
    ws = wb.Sheets(1)
    ws.Range("A1").Value = "Hello from COM"
    # 操作後にデアクティベート
    # (別のシェイプをクリックするか、スライドをクリック)

# OLEFormat の他のプロパティ
follow_colors = ole_shape.OLEFormat.FollowColors  # 配色追従
```

**OLE 操作の注意点:**
- `ClassName` と `FileName` は排他的。どちらか一方のみ指定する。
- `ClassName` 使用時は `Link = msoFalse` にする必要がある。
- OLE サーバーがインストールされていないとオブジェクトを操作できない。
- `OLEFormat.Object` は OLE サーバーのオブジェクトを返す。型はサーバーに依存する。
- COM オートメーションで OLE オブジェクトを操作する場合、メモリリークやプロセス残留に注意が必要。

---

## 6. Selection / Clipboard 操作

### 6.1 ActiveWindow.Selection

**プロパティ:** `Application.ActiveWindow.Selection` - Selection オブジェクトを返す

```python
app = win32com.client.Dispatch("PowerPoint.Application")
selection = app.ActiveWindow.Selection
```

**Selection オブジェクトの構造:**

| メソッド | 説明 |
|---|---|
| Copy() | 選択をクリップボードにコピー |
| Cut() | 選択を切り取り |
| Delete() | 選択を削除 |
| Unselect() | 選択を解除 |

| プロパティ | 説明 |
|---|---|
| ChildShapeRange | グループ内の子シェイプ範囲 |
| HasChildShapeRange | 子シェイプ範囲を持つか |
| ShapeRange | 選択されたシェイプ範囲 |
| SlideRange | 選択されたスライド範囲 |
| TextRange | 選択されたテキスト範囲 |
| TextRange2 | 選択されたテキスト範囲 (拡張版) |
| Type | 選択の種類 |

### 6.2 Selection.Type (選択の種類)

**PpSelectionType 列挙:**

| 定数 | 値 | 説明 |
|---|---|---|
| ppSelectionNone | 0 | 何も選択されていない |
| ppSelectionSlides | 1 | スライドが選択されている |
| ppSelectionShapes | 2 | シェイプが選択されている |
| ppSelectionText | 3 | テキストが選択されている |

```python
selection = app.ActiveWindow.Selection

# 選択の種類を確認
sel_type = selection.Type

if sel_type == 0:  # ppSelectionNone
    print("何も選択されていません")
elif sel_type == 1:  # ppSelectionSlides
    print("スライドが選択されています")
    slide_range = selection.SlideRange
elif sel_type == 2:  # ppSelectionShapes
    print("シェイプが選択されています")
    shape_range = selection.ShapeRange
elif sel_type == 3:  # ppSelectionText
    print("テキストが選択されています")
    text_range = selection.TextRange
```

**注意点:**
- スライドが切り替わると Selection はリセットされる (Type = ppSelectionNone)。
- 対応する Type でない場合に ShapeRange / TextRange / SlideRange にアクセスするとエラーが発生する。
- 必ず Type を確認してから適切なプロパティにアクセスすること。

### 6.3 Selection.ShapeRange

```python
selection = app.ActiveWindow.Selection

if selection.Type == 2:  # ppSelectionShapes
    shapes = selection.ShapeRange

    # 選択されたシェイプの数
    print(f"選択シェイプ数: {shapes.Count}")

    # 各シェイプの操作
    for i in range(1, shapes.Count + 1):
        shape = shapes(i)
        print(f"シェイプ {i}: {shape.Name}, 種類: {shape.Type}")

    # 全選択シェイプの塗りつぶし変更
    shapes.Fill.ForeColor.RGB = 0xFF6633

    # 全選択シェイプの位置変更
    # shapes.Left = 100
    # shapes.Top = 100
```

### 6.4 Selection.TextRange

```python
selection = app.ActiveWindow.Selection

if selection.Type == 3:  # ppSelectionText
    text_range = selection.TextRange

    # 選択されたテキストを取得
    selected_text = text_range.Text
    print(f"選択テキスト: {selected_text}")

    # 選択テキストの書式変更
    text_range.Font.Bold = True
    text_range.Font.Size = 18
    text_range.Font.Color.RGB = 0x0000FF  # 赤

    # テキストの置換
    text_range.Text = "新しいテキスト"
```

### 6.5 Selection.SlideRange

```python
selection = app.ActiveWindow.Selection

if selection.Type == 1:  # ppSelectionSlides
    slides = selection.SlideRange

    # 選択されたスライドの数
    print(f"選択スライド数: {slides.Count}")

    # スライドの複製
    slides.Duplicate()

    # スライドの背景変更
    for i in range(1, slides.Count + 1):
        slide = slides(i)
        slide.FollowMasterBackground = False
        slide.Background.Fill.ForeColor.RGB = 0xFFF8F0
```

### 6.6 コピー＆ペースト操作

```python
# シェイプのコピー＆ペースト
slide = prs.Slides(1)
shape = slide.Shapes(1)
shape.Copy()

# 同じスライドにペースト
slide.Shapes.Paste()

# 別のスライドにペースト
target_slide = prs.Slides(2)
target_slide.Shapes.Paste()

# スライドのコピー＆ペースト
slide.Copy()
prs.Slides.Paste()           # 末尾にペースト
prs.Slides.Paste(Index=3)    # 3番目の位置にペースト

# Selection 経由のコピー
app.ActiveWindow.Selection.Copy()
app.ActiveWindow.Selection.Cut()
```

### 6.7 特殊貼り付け (PasteSpecial)

**メソッド:** `Shapes.PasteSpecial(DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link)`

**PpPasteDataType 列挙:**

| 定数 | 値 | 説明 |
|---|---|---|
| ppPasteDefault | 0 | デフォルト |
| ppPasteBitmap | 1 | ビットマップ |
| ppPasteEnhancedMetafile | 2 | 拡張メタファイル |
| ppPasteMetafilePicture | 3 | メタファイル |
| ppPasteGIF | 4 | GIF画像 |
| ppPasteJPG | 5 | JPG画像 |
| ppPastePNG | 6 | PNG画像 |
| ppPasteText | 7 | テキスト |
| ppPasteHTML | 8 | HTML |
| ppPasteRTF | 9 | RTF |
| ppPasteOLEObject | 10 | OLEオブジェクト |
| ppPasteShape | 11 | シェイプ |

```python
slide = prs.Slides(1)

# Excel から範囲をコピーした後に PasteSpecial
# (Excel側で Range.Copy() を実行済みの前提)

# 拡張メタファイルとして貼り付け (高品質画像)
shape_range = slide.Shapes.PasteSpecial(DataType=2)  # ppPasteEnhancedMetafile

# ビットマップとして貼り付け
shape_range = slide.Shapes.PasteSpecial(DataType=1)  # ppPasteBitmap

# HTML として貼り付け
shape_range = slide.Shapes.PasteSpecial(DataType=8)  # ppPasteHTML

# OLE オブジェクトとして貼り付け (編集可能)
shape_range = slide.Shapes.PasteSpecial(DataType=10)  # ppPasteOLEObject

# リンクとして貼り付け
shape_range = slide.Shapes.PasteSpecial(
    DataType=10,  # ppPasteOLEObject
    Link=True     # msoTrue
)

# View 経由の PasteSpecial (特定のビューで)
app.ActiveWindow.View.PasteSpecial(DataType=2)
```

**Selection / Clipboard 操作の注意点:**
- クリップボードにデータがない状態で Paste/PasteSpecial を呼ぶとエラーが発生する。
- `PasteSpecial` はクリップボードの内容に対応する DataType のみ使用可能。
- COM オートメーションでの Clipboard 操作は、GUIのインタラクションが必要な場合があり、完全にヘッドレスでの操作は困難な場合がある。
- `View.PasteSpecial` と `Shapes.PasteSpecial` は異なるオブジェクトのメソッド。用途に応じて使い分けること。
- 複数のアプリケーション間でのコピー＆ペーストでは、タイミングの問題が発生することがある。

---

## 7. Hyperlink / ActionSettings

### 7.1 ハイパーリンクの追加・取得

**Hyperlink オブジェクトのプロパティ:**

| プロパティ | 型 | 説明 |
|---|---|---|
| Address | String | リンク先 URL またはファイルパス |
| SubAddress | String | ドキュメント内の位置 (スライド番号等) |
| ScreenTip | String | マウスオーバー時のヒントテキスト |
| TextToDisplay | String | 表示テキスト (図形に関連付いていない場合) |
| Type | MsoHyperlinkType | ハイパーリンクの種類 |

```python
slide = prs.Slides(1)
shape = slide.Shapes(1)

# ActionSettings 経由でハイパーリンクを設定 (推奨方法)
# クリック時のアクション
action = shape.ActionSettings(1)  # ppMouseClick = 1
action.Action = 7                  # ppActionHyperlink = 7
action.Hyperlink.Address = "https://www.example.com"
action.Hyperlink.ScreenTip = "Example Site"

# テキスト範囲にハイパーリンクを設定
text_range = shape.TextFrame.TextRange
action = text_range.ActionSettings(1)  # ppMouseClick = 1
action.Action = 7  # ppActionHyperlink
action.Hyperlink.Address = "https://www.google.com"

# 既存のハイパーリンクの取得
for i in range(1, slide.Hyperlinks.Count + 1):
    hl = slide.Hyperlinks(i)
    print(f"Address: {hl.Address}, SubAddress: {hl.SubAddress}")
```

#### SubAddress の形式 (スライド内リンク)

```python
# 別のスライドへのリンク
action = shape.ActionSettings(1)
action.Action = 7  # ppActionHyperlink
action.Hyperlink.Address = ""  # 同一ファイル内
action.Hyperlink.SubAddress = "256,3,スライドタイトル"
# 形式: "SlideID,SlideIndex,SlideTitle"
# SlideID: スライドの一意ID
# SlideIndex: スライド番号
# SlideTitle: スライドタイトル (いずれかで指定可能)
```

### 7.2 ActionSettings (クリック時、マウスオーバー時)

**ActionSettings コレクション:**
- `ActionSettings(1)` = `ppMouseClick` - クリック時のアクション
- `ActionSettings(2)` = `ppMouseOver` - マウスオーバー時のアクション

**PpActionType 列挙 (Action プロパティの値):**

| 定数 | 値 | 説明 |
|---|---|---|
| ppActionNone | 0 | アクションなし |
| ppActionNextSlide | 1 | 次のスライドに移動 |
| ppActionPreviousSlide | 2 | 前のスライドに移動 |
| ppActionFirstSlide | 3 | 最初のスライドに移動 |
| ppActionLastSlide | 4 | 最後のスライドに移動 |
| ppActionLastSlideViewed | 5 | 最後に表示したスライド |
| ppActionEndShow | 6 | スライドショーを終了 |
| ppActionHyperlink | 7 | ハイパーリンク |
| ppActionRunMacro | 8 | マクロを実行 |
| ppActionRunProgram | 9 | プログラムを実行 |
| ppActionNamedSlideShow | 10 | 指定のスライドショー |
| ppActionOLEVerb | 11 | OLE 動詞 |

```python
shape = slide.Shapes(1)

# クリック時: 次のスライドに移動
shape.ActionSettings(1).Action = 1  # ppActionNextSlide

# マウスオーバー時: 前のスライドに移動
shape.ActionSettings(2).Action = 2  # ppActionPreviousSlide

# クリック時: ハイパーリンク (URL)
action = shape.ActionSettings(1)
action.Action = 7  # ppActionHyperlink
action.Hyperlink.Address = "https://www.example.com"

# クリック時: 特定のスライドに移動
action = shape.ActionSettings(1)
action.Action = 7  # ppActionHyperlink
action.Hyperlink.Address = ""
action.Hyperlink.SubAddress = ",5,"  # スライド5に移動

# クリック時: プログラムの実行
action = shape.ActionSettings(1)
action.Action = 9  # ppActionRunProgram
action.Run = r"C:\Windows\notepad.exe"

# クリック時: サウンドの再生
action = shape.ActionSettings(1)
action.SoundEffect.ImportFromFile(r"C:\Sounds\click.wav")

# アクションの確認
current_action = shape.ActionSettings(1).Action
print(f"クリック時アクション: {current_action}")
```

### 7.3 ジャンプ先の種類

```python
# URL へのリンク
action.Action = 7  # ppActionHyperlink
action.Hyperlink.Address = "https://www.example.com"

# ファイルへのリンク
action.Action = 7
action.Hyperlink.Address = r"C:\Documents\report.pdf"

# メールリンク
action.Action = 7
action.Hyperlink.Address = "mailto:user@example.com?subject=Subject"

# スライドへのリンク
action.Action = 7
action.Hyperlink.Address = ""
action.Hyperlink.SubAddress = ",3,"  # スライド3

# 別のプレゼンテーションの特定スライドへ
action.Action = 7
action.Hyperlink.Address = r"C:\Presentations\other.pptx"
action.Hyperlink.SubAddress = ",2,"  # other.pptx のスライド2
```

**Hyperlink / ActionSettings の注意点:**
- ハイパーリンクは ActionSettings 経由で設定するのが確実。
- `Action` プロパティを `ppActionHyperlink (7)` に設定してから Hyperlink プロパティを設定する。
- SubAddress の形式は "SlideID,SlideIndex,SlideTitle" でカンマ区切り。不要な項目は空白にする。
- TextToDisplay はテキストに関連付けられたハイパーリンクでのみ使用可能。グラフィック関連付けの場合はエラーになる。
- マウスオーバーアクションはスライドショー中のみ機能する。

---

## 8. Animation / Transition

### 8.1 AnimationSettings vs Timeline.MainSequence

PowerPoint には 2 つのアニメーション API がある:

| 特徴 | AnimationSettings | Timeline.MainSequence |
|---|---|---|
| 対象バージョン | 旧バージョン (互換性用) | PowerPoint 2002 以降 (推奨) |
| 機能 | 基本的なアニメーション | 完全なアニメーション制御 |
| アクセス | Shape.AnimationSettings | Slide.TimeLine.MainSequence |
| 推奨 | 非推奨 | 推奨 |

```python
# 旧 API (AnimationSettings) - 非推奨
shape.AnimationSettings.EntryEffect = 3  # ppEffectBlindsHorizontal

# 新 API (Timeline.MainSequence) - 推奨
slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=1,   # msoAnimEffectAppear
    trigger=1     # msoAnimTriggerOnPageClick
)
```

### 8.2 Effect の追加 (AddEffect)

**メソッド:** `Sequence.AddEffect(Shape, effectId, Level, trigger, Index)`

| パラメータ | 必須/任意 | 型 | 説明 |
|---|---|---|---|
| Shape | 必須 | Shape | アニメーション対象のシェイプ |
| effectId | 必須 | MsoAnimEffect | アニメーション効果の種類 |
| Level | 任意 | MsoAnimateByLevel | レベル (チャート・テキスト用) |
| trigger | 任意 | MsoAnimTriggerType | トリガーの種類 |
| Index | 任意 | Long | シーケンス内の位置 (-1で末尾) |

```python
slide = prs.Slides(1)
shape = slide.Shapes(1)

# アニメーション効果の追加
effect = slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=1,   # msoAnimEffectAppear
    trigger=1     # msoAnimTriggerOnPageClick
)

# フェードインの追加
effect_fade = slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=10,   # msoAnimEffectFade
    trigger=1      # msoAnimTriggerOnPageClick
)

# フライインの追加
effect_fly = slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=2,    # msoAnimEffectFly
    trigger=1
)
```

### 8.3 主要な MsoAnimEffect 定数

| 定数 | 値 | 説明 | カテゴリ |
|---|---|---|---|
| msoAnimEffectAppear | 1 | 出現 | 開始 |
| msoAnimEffectFly | 2 | フライイン | 開始 |
| msoAnimEffectBlinds | 3 | ブラインド | 開始 |
| msoAnimEffectBox | 4 | ボックス | 開始 |
| msoAnimEffectCheckerboard | 5 | チェッカーボード | 開始 |
| msoAnimEffectCircle | 6 | 円 | 開始 |
| msoAnimEffectCrawl | 7 | クロール | 開始 |
| msoAnimEffectDiamond | 8 | ダイヤモンド | 開始 |
| msoAnimEffectDissolve | 9 | ディゾルブ | 開始 |
| msoAnimEffectFade | 10 | フェード | 開始 |
| msoAnimEffectFlashOnce | 11 | フラッシュ | 開始 |
| msoAnimEffectPeek | 12 | ピーク | 開始 |
| msoAnimEffectPlus | 13 | プラス | 開始 |
| msoAnimEffectRandomBars | 14 | ランダムバー | 開始 |
| msoAnimEffectSpiral | 15 | スパイラル | 開始 |
| msoAnimEffectSplit | 16 | スプリット | 開始 |
| msoAnimEffectStretch | 17 | ストレッチ | 開始 |
| msoAnimEffectStrips | 18 | ストリップ | 開始 |
| msoAnimEffectSwivel | 19 | スウィベル | 開始 |
| msoAnimEffectWedge | 20 | ウェッジ | 開始 |
| msoAnimEffectWheel | 21 | ホイール | 開始 |
| msoAnimEffectWipe | 22 | ワイプ | 開始 |
| msoAnimEffectZoom | 23 | ズーム | 開始 |
| msoAnimEffectRandomEffects | 24 | ランダム | 開始 |
| msoAnimEffectBoomerang | 25 | ブーメラン | 開始 |
| msoAnimEffectBounce | 26 | バウンス | 開始 |
| msoAnimEffectColorReveal | 27 | カラーリビール | 開始 |
| msoAnimEffectFloat | 56 | フロート | 開始 |
| msoAnimEffectGrowAndTurn | 57 | 拡大して回転 | 開始 |
| msoAnimEffectSpin | 61 | スピン | 強調 |
| msoAnimEffectChangeFillColor | 54 | 塗りつぶし色変更 | 強調 |
| msoAnimEffectChangeFont | 55 | フォント変更 | 強調 |
| msoAnimEffectChangeFontColor | 58 | フォント色変更 | 強調 |
| msoAnimEffectChangeFontSize | 59 | フォントサイズ変更 | 強調 |
| msoAnimEffectTransparency | 62 | 透過 | 強調 |
| msoAnimEffectPathDown | 64 | 下へ移動 | モーションパス |
| msoAnimEffectPathUp | 65 | 上へ移動 | モーションパス |
| msoAnimEffectPathLeft | 66 | 左へ移動 | モーションパス |
| msoAnimEffectPathRight | 67 | 右へ移動 | モーションパス |
| msoAnimEffectPathDiamond | 68 | ダイヤモンドパス | モーションパス |
| msoAnimEffectPathCircle | 69 | 円パス | モーションパス |

### 8.4 Effect.Timing (タイミング設定)

```python
effect = slide.TimeLine.MainSequence(1)

# 再生時間 (秒)
effect.Timing.Duration = 2.0

# 遅延時間 (秒)
effect.Timing.TriggerDelayTime = 0.5

# 繰り返し回数
effect.Timing.RepeatCount = 3

# 繰り返し期間 (秒)
effect.Timing.RepeatDuration = 10.0

# 加速 (0.0 - 1.0)
effect.Timing.Accelerate = 0.3  # 開始30%で加速

# 減速 (0.0 - 1.0)
effect.Timing.Decelerate = 0.3  # 終了30%で減速

# 自動的に逆再生
effect.Timing.AutoReverse = True

# 効果終了後に巻き戻し
effect.Timing.RewindAtEnd = True

# スムースな開始/終了
effect.Timing.SmoothStart = True
effect.Timing.SmoothEnd = True

# リスタートの設定
# msoAnimEffectRestartAlways = 1
# msoAnimEffectRestartWhenOff = 2
# msoAnimEffectRestartNever = 3
effect.Timing.Restart = 3  # msoAnimEffectRestartNever
```

### 8.5 Trigger (トリガーの種類)

**MsoAnimTriggerType 列挙:**

| 定数 | 値 | 説明 |
|---|---|---|
| msoAnimTriggerNone | 0 | トリガーなし |
| msoAnimTriggerOnPageClick | 1 | クリック時 |
| msoAnimTriggerWithPrevious | 2 | 前と同時 |
| msoAnimTriggerAfterPrevious | 3 | 前の後 |
| msoAnimTriggerOnShapeClick | 4 | シェイプクリック時 |

```python
# クリック時に開始
effect = slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=10,  # msoAnimEffectFade
    trigger=1     # msoAnimTriggerOnPageClick
)

# 前のアニメーションと同時に開始
effect2 = slide.TimeLine.MainSequence.AddEffect(
    Shape=another_shape,
    effectId=10,
    trigger=2  # msoAnimTriggerWithPrevious
)

# 前のアニメーションの後に開始
effect3 = slide.TimeLine.MainSequence.AddEffect(
    Shape=third_shape,
    effectId=10,
    trigger=3  # msoAnimTriggerAfterPrevious
)

# シェイプクリック時に開始 (インタラクティブトリガー)
# InteractiveSequences を使用
interactive_seq = slide.TimeLine.InteractiveSequences.Add()
effect_interactive = interactive_seq.AddEffect(
    Shape=target_shape,
    effectId=1,  # msoAnimEffectAppear
    trigger=4    # msoAnimTriggerOnShapeClick
)
effect_interactive.Timing.TriggerShape = trigger_shape
```

### 8.6 MotionEffect (モーションパス)

```python
# モーションパスアニメーションの追加
effect = slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=64,  # msoAnimEffectPathDown
    trigger=1
)

# Behavior を通じたモーションエフェクトの操作
for i in range(1, effect.Behaviors.Count + 1):
    behavior = effect.Behaviors(i)
    if behavior.Type == 9:  # msoAnimTypeMotion
        motion = behavior.MotionEffect
        motion.FromX = 0
        motion.FromY = 0
        motion.ToX = 0.5   # 相対座標 (スライド幅に対する割合)
        motion.ToY = 0.5
        # motion.Path = "M 0 0 L 0.5 0.5 E"  # SVG パス形式

# カスタムモーションパスの追加
effect_custom = slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=63,  # msoAnimEffectPath (カスタムパス)
    trigger=1
)
```

### 8.7 Effect の Exit (終了) アニメーション

```python
# 終了アニメーションの追加
effect = slide.TimeLine.MainSequence.AddEffect(
    Shape=shape,
    effectId=10,  # msoAnimEffectFade
    trigger=1
)
effect.Exit = True  # msoTrue - 終了アニメーションに設定

# Exit = True にすると、同じ effectId でも「フェードアウト」になる
```

### 8.8 SlideShowTransition (スライド切り替え効果)

**プロパティ:** `Slide.SlideShowTransition` - SlideShowTransition オブジェクトを返す

**SlideShowTransition プロパティ:**

| プロパティ | 型 | 説明 |
|---|---|---|
| AdvanceOnClick | MsoTriState | クリックで進むか |
| AdvanceOnTime | MsoTriState | 時間で自動進行するか |
| AdvanceTime | Long | 自動進行までの時間 (秒) |
| Duration | Single | トランジションの再生時間 (秒) |
| EntryEffect | PpEntryEffect | 切り替え効果の種類 |
| Hidden | MsoTriState | スライドを非表示にするか |
| LoopSoundUntilNext | MsoTriState | 次までサウンドをループ |
| SoundEffect | SoundEffect | サウンドエフェクト |
| Speed | PpTransitionSpeed | 切り替え速度 |

**PpTransitionSpeed 列挙:**

| 定数 | 値 | 説明 |
|---|---|---|
| ppTransitionSpeedSlow | 3 | 遅い |
| ppTransitionSpeedMedium | 2 | 普通 |
| ppTransitionSpeedFast | 1 | 速い |
| ppTransitionSpeedMixed | -2 | 混合 |

```python
slide = prs.Slides(1)

# スライド切り替え効果の設定
transition = slide.SlideShowTransition
transition.EntryEffect = 3844  # ppEffectFade
transition.Duration = 1.5      # 1.5秒

# 自動進行の設定
transition.AdvanceOnClick = True   # クリックでも進める
transition.AdvanceOnTime = True    # 時間でも進める
transition.AdvanceTime = 5         # 5秒後に自動進行

# 切り替え速度
transition.Speed = 1  # ppTransitionSpeedFast

# サウンド効果
transition.SoundEffect.ImportFromFile(r"C:\Sounds\transition.wav")
transition.LoopSoundUntilNext = False

# スライドを非表示にする
transition.Hidden = False

# 全スライドに同じ設定を適用するには SlideShowSettings を使用
prs.SlideShowSettings.AdvanceMode = 2  # ppSlideShowUseSlideTimings
```

**主要な PpEntryEffect 定数:**

| 定数 | 値 | 説明 |
|---|---|---|
| ppEffectNone | 0 | 効果なし |
| ppEffectCut | 257 | カット |
| ppEffectFade | 3844 | フェード |
| ppEffectPush | 3845 | プッシュ |
| ppEffectWipe | 3846 | ワイプ |
| ppEffectSplit | 3847 | スプリット |
| ppEffectReveal | 3848 | リビール |
| ppEffectBlindsHorizontal | 769 | 水平ブラインド |
| ppEffectBlindsVertical | 770 | 垂直ブラインド |
| ppEffectCheckerboardAcross | 1025 | チェッカーボード(横) |
| ppEffectDissolve | 1537 | ディゾルブ |
| ppEffectCoverDown | 1284 | カバー(下) |
| ppEffectCoverRight | 1281 | カバー(右) |
| ppEffectStripsDownLeft | 2305 | ストリップ(左下) |
| ppEffectRandom | 513 | ランダム |

**Animation / Transition の注意点:**
- `MainSequence` が空の状態でインデックスアクセスするとエラーが発生する。
- `effect.Exit = True` で終了アニメーション (例: フェードアウト) になる。
- InteractiveSequences はクリックトリガーのアニメーション用。
- モーションパスの座標はスライドサイズに対する相対値 (0.0 - 1.0)。
- SlideShowTransition は各スライドに個別に設定される。
- `Duration` プロパティ (新しい) と `Speed` プロパティ (旧) は両方存在するが、`Duration` の使用が推奨される。
- EntryEffect の定数は非常に多い。使用頻度の高いものを上記に列挙している。

---

## 9. Theme / Design

### 9.1 SlideMaster (スライドマスター)

```python
prs = app.ActivePresentation

# スライドマスターへのアクセス
# Design (デザイン) 経由
design = prs.Designs(1)
master = design.SlideMaster

# シェイプの操作
for i in range(1, master.Shapes.Count + 1):
    shape = master.Shapes(i)
    print(f"マスターシェイプ: {shape.Name}, Type: {shape.Type}")

# マスターの背景
master.Background.Fill.ForeColor.RGB = 0xF5F5F5

# プレゼンテーションの SlideMaster
slide_master = prs.SlideMaster
# 注意: プレゼンテーションに複数のデザインがある場合は Designs コレクション経由が正確
```

### 9.2 CustomLayout

```python
# カスタムレイアウトの一覧
master = prs.Designs(1).SlideMaster
layouts = master.CustomLayouts

for i in range(1, layouts.Count + 1):
    layout = layouts(i)
    print(f"Layout {i}: {layout.Name}")

# スライドにカスタムレイアウトを適用
slide = prs.Slides(1)
slide.CustomLayout = master.CustomLayouts(2)  # 2番目のレイアウトを適用

# レイアウト名で検索して適用
for i in range(1, layouts.Count + 1):
    if layouts(i).Name == "タイトルとコンテンツ":
        slide.CustomLayout = layouts(i)
        break

# カスタムレイアウトの追加
new_layout = layouts.Add(1)  # インデックス1の位置に追加
new_layout.Name = "カスタムレイアウト"

# レイアウト上のプレースホルダー操作
for i in range(1, new_layout.Shapes.Placeholders.Count + 1):
    ph = new_layout.Shapes.Placeholders(i)
    print(f"Placeholder: {ph.PlaceholderFormat.Type}")
```

**PpSlideLayout 定数 (主要なもの):**

| 定数 | 値 | 説明 |
|---|---|---|
| ppLayoutTitle | 1 | タイトルスライド |
| ppLayoutText | 2 | タイトルとコンテンツ |
| ppLayoutTwoColumnText | 3 | 2段組テキスト |
| ppLayoutTable | 4 | テーブル |
| ppLayoutTextAndChart | 5 | テキストとグラフ |
| ppLayoutChartAndText | 6 | グラフとテキスト |
| ppLayoutOrgchart | 7 | 組織図 |
| ppLayoutChart | 8 | グラフ |
| ppLayoutTextAndClipart | 9 | テキストとクリップアート |
| ppLayoutBlank | 12 | 白紙 |
| ppLayoutTitleOnly | 11 | タイトルのみ |
| ppLayoutCustom | 32 | カスタム |

### 9.3 Theme.ThemeColorScheme

```python
# スライドのテーマカラースキームにアクセス
slide = prs.Slides(1)
color_scheme = slide.ThemeColorScheme

# カラースキームの色を取得 (12色)
# MsoThemeColorSchemeIndex:
# msoThemeDark1 = 1, msoThemeLight1 = 2
# msoThemeDark2 = 3, msoThemeLight2 = 4
# msoThemeAccent1 = 5, msoThemeAccent2 = 6
# msoThemeAccent3 = 7, msoThemeAccent4 = 8
# msoThemeAccent5 = 9, msoThemeAccent6 = 10
# msoThemeHyperlink = 11, msoThemeFollowedHyperlink = 12

for i in range(1, 13):
    color = color_scheme(i)
    print(f"Color {i}: RGB={color.RGB}")

# テーマカラーの変更
color_scheme(5).RGB = 0xFF6633  # Accent1 の色を変更

# テーマカラースキームファイルの読み込み
# color_scheme.Load(r"C:\Themes\custom_colors.xml")
```

### 9.4 Theme.ThemeFontScheme

```python
# テーマフォントスキーム
# Designs 経由でアクセス
design = prs.Designs(1)
theme = design.SlideMaster.Theme

# フォント情報へのアクセス
font_scheme = theme.ThemeFontScheme
major_font = font_scheme.MajorFont  # 見出しフォント
minor_font = font_scheme.MinorFont  # 本文フォント

# Latin フォントの取得/設定
print(f"見出しフォント: {major_font(1).Name}")  # Latin
print(f"本文フォント: {minor_font(1).Name}")    # Latin

# 東アジアフォントの取得
# major_font / minor_font コレクションには言語別のフォント設定が含まれる
```

### 9.5 ColorScheme の変更

```python
# 旧 API (ColorScheme) - 互換性のため残されている
# 新しい API は ThemeColorScheme を使用

# スライドのカラースキーム (旧)
# slide.ColorScheme.Colors(1).RGB = 0x000000  # Background
# slide.ColorScheme.Colors(2).RGB = 0xFFFFFF  # Text
# slide.ColorScheme.Colors(3).RGB = 0x808080  # Shadow
# slide.ColorScheme.Colors(4).RGB = 0x000000  # Title
# slide.ColorScheme.Colors(5).RGB = 0xFF6633  # Fill
# slide.ColorScheme.Colors(6).RGB = 0x6699FF  # Accent
# slide.ColorScheme.Colors(7).RGB = 0xFF0000  # Accent2
# slide.ColorScheme.Colors(8).RGB = 0x0000FF  # Accent3
```

### 9.6 デザインテンプレートの適用

```python
# テーマの適用
slide = prs.Slides(1)
slide.ApplyTheme(r"C:\Program Files\Microsoft Office\root\Document Themes 16\Facet.thmx")

# テンプレートの適用 (バリアントを含む)
slide.ApplyTemplate2(
    r"C:\Program Files\Microsoft Office\root\Document Themes 16\Facet.thmx",
    "Variant1"  # バリアント名
)

# SlideRange に適用
slides = prs.Slides.Range([1, 2, 3])
slides.ApplyTheme(r"C:\Themes\custom.thmx")

# テーマカラースキームの適用
slide.ApplyThemeColorScheme(r"C:\Themes\custom_colors.xml")

# プレゼンテーション全体にデザインを適用
prs.ApplyTemplate(r"C:\Templates\corporate.potx")
```

**Theme / Design の注意点:**
- テーマファイル (.thmx) のパスは環境により異なる。
- `ApplyTemplate2` は Office 2013 以降で利用可能。
- ThemeColorScheme の変更はそのスライド (またはマスター) に適用される。
- CustomLayouts の操作は SlideMaster 経由で行う。
- 複数の Designs を持つプレゼンテーションでは、各 Design の SlideMaster が異なる。
- テーマの変更は既存のコンテンツの見た目に影響を与える。
- フォントスキームの変更は「テーマフォント」を使用している要素にのみ影響する。

---

## 10. MCP サーバー実装への推奨事項

### 10.1 実装優先度の推奨

#### 高優先度 (必須)

| 機能 | 理由 |
|---|---|
| **Table 基本操作** | ビジネス資料で頻繁に使用。作成・テキスト設定・書式設定は必須 |
| **Table セルアクセス** | Cell(row, col) でのデータ設定は最も基本的な操作 |
| **Hyperlink** | URL リンクの追加は頻出ニーズ |
| **SlideShowTransition (基本)** | 切り替え効果の設定は一般的な要件 |

#### 中優先度 (推奨)

| 機能 | 理由 |
|---|---|
| **Chart 基本操作** | グラフ作成のニーズは多いが、Excel Workbook 経由のため実装が複雑 |
| **Table スタイル・バンド** | テーブルの見た目改善に有用 |
| **Animation 基本** | AddEffect でのシンプルなアニメーション追加 |
| **Selection 読み取り** | 現在の選択状態の確認は有用 |
| **Theme 適用** | テーマの切り替えは一般的な操作 |

#### 低優先度 (将来実装)

| 機能 | 理由 |
|---|---|
| **SmartArt** | レイアウトがバージョン依存で不安定 |
| **Media** | メディアファイルのパス依存が大きい |
| **OLEObject** | 複雑な OLE サーバー依存。メモリリスクが高い |
| **Clipboard 操作** | GUI インタラクション依存で自動化が困難 |
| **MotionPath** | 高度なアニメーション。ニーズが限定的 |
| **ActionSettings 高度** | マクロ実行等はセキュリティリスクがある |

### 10.2 実装時の注意点

#### COM オブジェクトのライフサイクル管理

```python
import win32com.client
import pythoncom

# COM オブジェクトの適切な解放
def safe_release(*com_objects):
    """COM オブジェクトを安全に解放する"""
    for obj in com_objects:
        if obj is not None:
            try:
                del obj
            except:
                pass
    pythoncom.CoUninitialize()
```

#### Chart 操作時の Excel プロセス管理

```python
# Chart 操作後は必ず Workbook を閉じる
def set_chart_data(chart, data):
    """チャートデータを安全に設定する"""
    try:
        chart.ChartData.Activate()
        wb = chart.ChartData.Workbook
        ws = wb.Worksheets(1)
        # データ設定...
    finally:
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
```

#### エラーハンドリング

```python
# Selection 操作時の安全なアクセス
def get_selected_shapes(app):
    """選択されたシェイプを安全に取得する"""
    try:
        selection = app.ActiveWindow.Selection
        if selection.Type == 2:  # ppSelectionShapes
            return selection.ShapeRange
    except Exception as e:
        print(f"選択取得エラー: {e}")
    return None
```

### 10.3 MCP ツール設計案

```
# Table 関連ツール
- add_table(slide_index, rows, cols, left, top, width, height)
- set_table_cell_text(slide_index, shape_index, row, col, text)
- set_table_cell_format(slide_index, shape_index, row, col, format_options)
- merge_table_cells(slide_index, shape_index, from_row, from_col, to_row, to_col)
- apply_table_style(slide_index, shape_index, style_id)
- add_table_row(slide_index, shape_index, before_row?)
- add_table_column(slide_index, shape_index, before_col?)
- delete_table_row(slide_index, shape_index, row_index)
- delete_table_column(slide_index, shape_index, col_index)

# Chart 関連ツール
- add_chart(slide_index, chart_type, data, left, top, width, height)
- set_chart_data(slide_index, shape_index, data)
- set_chart_title(slide_index, shape_index, title_text)
- set_chart_style(slide_index, shape_index, style_number)

# Hyperlink 関連ツール
- add_hyperlink(slide_index, shape_index, url, screen_tip?)
- add_slide_link(slide_index, shape_index, target_slide)

# Animation 関連ツール
- add_animation(slide_index, shape_index, effect_type, trigger, duration)
- set_slide_transition(slide_index, effect, duration, advance_time?)

# Theme 関連ツール
- apply_theme(theme_path)
- set_slide_layout(slide_index, layout_index_or_name)
```

### 10.4 定数管理の推奨

win32com で定数にアクセスする方法:

```python
import win32com.client

# EnsureDispatch を使うと定数が利用可能になる
app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")

# constants 経由でアクセス
from win32com.client import constants
print(constants.ppLayoutBlank)  # 12

# ただし、全ての定数が利用可能とは限らない
# 数値を直接使用する方が安全な場合がある
```

MCP サーバーでは定数を Python の辞書またはEnum として管理することを推奨:

```python
from enum import IntEnum

class PpSlideLayout(IntEnum):
    ppLayoutTitle = 1
    ppLayoutText = 2
    ppLayoutBlank = 12
    ppLayoutTitleOnly = 11
    ppLayoutCustom = 32

class XlChartType(IntEnum):
    xlColumnClustered = 51
    xlLine = 4
    xlPie = 5
    xlBarClustered = 57
    xlArea = 1
    xlXYScatter = -4169

class MsoAnimEffect(IntEnum):
    msoAnimEffectAppear = 1
    msoAnimEffectFly = 2
    msoAnimEffectFade = 10
    msoAnimEffectWipe = 22
    msoAnimEffectSpin = 61

class PpPasteDataType(IntEnum):
    ppPasteDefault = 0
    ppPasteBitmap = 1
    ppPasteEnhancedMetafile = 2
    ppPasteHTML = 8
    ppPasteOLEObject = 10

class MsoAnimTriggerType(IntEnum):
    msoAnimTriggerNone = 0
    msoAnimTriggerOnPageClick = 1
    msoAnimTriggerWithPrevious = 2
    msoAnimTriggerAfterPrevious = 3
    msoAnimTriggerOnShapeClick = 4
```

---

## 参考資料

- [Shapes.AddTable method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addtable)
- [Table object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.table)
- [Cell.Merge method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.cell.merge)
- [Cell.Split method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Cell.Split)
- [Table.ApplyStyle method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Table.ApplyStyle)
- [Shapes.AddChart2 method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2)
- [Chart.ChartData property (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Chart.ChartData)
- [ChartData.Activate method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.ChartData.Activate)
- [Shapes.AddSmartArt method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addsmartart)
- [Shapes.AddMediaObject2 method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2)
- [MediaFormat object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.MediaFormat)
- [PlaySettings object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.playsettings)
- [Shapes.AddOLEObject method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addoleobject)
- [OLEFormat object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.oleformat)
- [Selection object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.selection)
- [Shapes.PasteSpecial method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.pastespecial)
- [PpPasteDataType enumeration (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype)
- [Hyperlink object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.hyperlink)
- [ActionSetting object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.ActionSetting)
- [Sequence.AddEffect method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sequence.addeffect)
- [Effect object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.effect)
- [Timing object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.timing)
- [MsoAnimEffect enumeration (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect)
- [SlideShowTransition object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowtransition)
- [CustomLayouts object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customlayouts)
- [Slide.ApplyTemplate2 method (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.slide.applytemplate2)
- [WIN32 automation of PowerPoint - GitHub Gist](https://gist.github.com/dmahugh/f642607d50cd008cc752f1344e9809e6)
- [sigma_coding_youtube - Win32COM PowerPoint Chart Objects](https://github.com/areed1192/sigma_coding_youtube/blob/master/python/python-vba-powerpoint/Win32COM%20-%20PowerPoint%20To%20Excel%20-%20Chart%20Objects.py)
