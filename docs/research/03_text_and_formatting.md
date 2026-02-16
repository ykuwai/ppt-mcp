# PowerPoint COM テキスト書式・フォント・色・塗りつぶし・線 調査レポート

> **調査日**: 2026-02-16
> **目的**: MCP サーバーで実装すべきテキスト書式設定・色制御・塗りつぶし・線の機能を洗い出す
> **対象**: PowerPoint COM オブジェクトモデル (win32com 経由で Python から利用)

---

## 目次

1. [TextFrame / TextFrame2 - テキストフレーム](#1-textframe--textframe2---テキストフレーム)
2. [TextRange - テキスト操作](#2-textrange---テキスト操作)
3. [Font - フォント設定](#3-font---フォント設定)
4. [ParagraphFormat - 段落書式](#4-paragraphformat---段落書式)
5. [Fill - 塗りつぶし](#5-fill---塗りつぶし)
6. [Line - 枠線](#6-line---枠線)
7. [Shadow / Glow / Reflection / SoftEdge - 効果](#7-shadow--glow--reflection--softedge---効果)
8. [ThreeDFormat - 3D 効果](#8-threedformat---3d-効果)
9. [部分テキスト色変更の詳細実装例（重要）](#9-部分テキスト色変更の詳細実装例重要)
10. [MCP サーバー実装への提言](#10-mcp-サーバー実装への提言)

---

## 1. TextFrame / TextFrame2 - テキストフレーム

### 1.1 TextFrame オブジェクト

**階層パス**: `Shape.TextFrame`

TextFrame はシェイプ内のテキストコンテナを表すオブジェクト。テキストの内容と、テキストフレームの配置・アンカーを制御するプロパティ・メソッドを持つ。

#### プロパティ一覧

| プロパティ | 型 | 説明 |
|---|---|---|
| `HasText` | MsoTriState | テキストフレームにテキストが含まれているかどうか。読み取り専用 |
| `TextRange` | TextRange | テキストフレーム内のテキスト範囲オブジェクトを返す |
| `WordWrap` | MsoTriState | テキストがシェイプ内に収まるように自動折り返しするか |
| `AutoSize` | PpAutoSize | テキストに合わせてシェイプのサイズを自動調整するかどうか |
| `MarginLeft` | Single | テキストフレームの左余白（ポイント単位） |
| `MarginRight` | Single | テキストフレームの右余白（ポイント単位） |
| `MarginTop` | Single | テキストフレームの上余白（ポイント単位） |
| `MarginBottom` | Single | テキストフレームの下余白（ポイント単位） |
| `Orientation` | MsoTextOrientation | テキストの方向（横書き、縦書きなど） |
| `HorizontalAnchor` | MsoHorizontalAnchor | 水平方向のアンカー位置 |
| `VerticalAnchor` | MsoVerticalAnchor | 垂直方向のアンカー位置 |
| `Ruler` | Ruler | ルーラーオブジェクト（タブ位置やインデント） |

#### メソッド

| メソッド | 説明 |
|---|---|
| `DeleteText` | テキストフレーム内のテキストとその書式プロパティを全て削除 |

#### AutoSize の定数値

| 定数 | 値 | 説明 |
|---|---|---|
| `ppAutoSizeNone` | 0 | 自動サイズ調整なし |
| `ppAutoSizeMixed` | -2 | 混在 |
| `ppAutoSizeShapeToFitText` | 1 | テキストに合わせてシェイプを調整 |

#### Orientation の定数値

| 定数 | 値 | 説明 |
|---|---|---|
| `msoTextOrientationHorizontal` | 1 | 横書き |
| `msoTextOrientationVertical` | 5 | 縦書き（上から下、右から左） |
| `msoTextOrientationVerticalFarEast` | 6 | 縦書き（東アジア向け） |
| `msoTextOrientationUpward` | 2 | 下から上 |
| `msoTextOrientationDownward` | 3 | 上から下 |

#### Python win32com 例

```python
import win32com.client

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True
pres = ppt.Presentations.Open(r"C:\path\to\file.pptx")
slide = pres.Slides(1)

shape = slide.Shapes(1)

# テキストフレームが存在するか確認
if shape.HasTextFrame:
    tf = shape.TextFrame

    # テキストがあるか確認
    if tf.HasText:
        print(tf.TextRange.Text)

    # 余白設定（ポイント単位）
    tf.MarginLeft = 10
    tf.MarginRight = 10
    tf.MarginTop = 5
    tf.MarginBottom = 5

    # ワードラップ有効化（msoTrue = -1）
    tf.WordWrap = -1  # msoTrue

    # テキスト方向を横書きに
    tf.Orientation = 1  # msoTextOrientationHorizontal

    # AutoSize 無効化
    tf.AutoSize = 0  # ppAutoSizeNone
```

### 1.2 TextFrame2 オブジェクト

**階層パス**: `Shape.TextFrame2`

TextFrame2 は Office 2007 以降で追加された強化版テキストフレーム。TextFrame の全機能に加え、段組み・3D・ワードアートなどの高度な機能を持つ。

#### TextFrame2 の追加プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `Column` | TextColumn2 | 段組み（カラム）設定オブジェクト |
| `ThreeD` | ThreeDFormat | テキストの 3D 書式 |
| `WarpFormat` | MsoWarpFormat | テキストのワープ（変形）書式 |
| `WordArtFormat` | MsoPresetTextEffect | ワードアート書式 |
| `PathFormat` | MsoPathFormat | テキストパス書式 |
| `NoTextRotation` | MsoTriState | テキストがシェイプと一緒に回転しないか |

#### TextFrame と TextFrame2 の違い

| 機能 | TextFrame | TextFrame2 |
|---|---|---|
| 返す TextRange | TextRange | TextRange2（拡張版） |
| 返す Font | Font | Font2（拡張フォント効果を含む） |
| 段組み (Column) | なし | あり |
| 3D 効果 (ThreeD) | なし | あり |
| ワードアート | なし | あり |
| 導入バージョン | Office 2003 以前 | Office 2007 |
| 文字間隔/カーニング | 限定的 | 詳細制御可能 |

**注意**: TextFrame2.TextRange は `TextRange2` オブジェクトを返す。これは `TextRange` とは異なるオブジェクトで、メソッド名やプロパティ名が一部異なる。MCP サーバー実装では `TextFrame`（TextRange を返す）の方が一般的で情報も多いため、まずこちらを優先実装し、TextFrame2 固有の機能（段組みなど）は段階的に追加する方針が望ましい。

#### Python win32com 例（TextFrame2）

```python
shape = slide.Shapes(1)
tf2 = shape.TextFrame2

# テキスト設定
tf2.TextRange.Text = "サンプルテキスト"

# 余白設定
tf2.MarginBottom = 10
tf2.MarginLeft = 10
tf2.MarginRight = 10
tf2.MarginTop = 10

# 段組み設定（TextFrame2 固有）
tf2.Column.Number = 2    # 2段組み
tf2.Column.Spacing = 36  # カラム間のスペース（ポイント単位）
```

---

## 2. TextRange - テキスト操作

**階層パス**: `Shape.TextFrame.TextRange`

TextRange オブジェクトはシェイプに付属するテキストを保持し、テキストの操作・書式設定のためのプロパティとメソッドを提供する。

### 2.1 プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `Text` | String | テキスト範囲のプレーンテキスト文字列。読み書き可能 |
| `Font` | Font | テキスト範囲のフォント書式オブジェクト。読み取り専用 |
| `ParagraphFormat` | ParagraphFormat | 段落書式オブジェクト。読み取り専用 |
| `IndentLevel` | Long | インデントレベル（1〜9） |
| `LanguageID` | MsoLanguageID | 言語 ID |
| `Length` | Long | テキスト範囲の文字数。読み取り専用 |
| `Start` | Long | テキスト範囲の開始位置（1始まり）。読み取り専用 |
| `Count` | Long | テキスト範囲のアイテム数。読み取り専用 |
| `BoundHeight` | Single | テキスト範囲のバウンディングボックスの高さ |
| `BoundLeft` | Single | テキスト範囲のバウンディングボックスの左位置 |
| `BoundTop` | Single | テキスト範囲のバウンディングボックスの上位置 |
| `BoundWidth` | Single | テキスト範囲のバウンディングボックスの幅 |
| `ActionSettings` | ActionSettings | ハイパーリンク/アクション設定 |

### 2.2 テキスト部分取得メソッド

これらのメソッドは全て TextRange オブジェクトを返す。返された TextRange に対してさらに `.Font` や `.ParagraphFormat` でスタイルを変更可能。

#### Characters(Start, Length)

文字単位でテキストの部分範囲を取得する。**部分テキストの書式変更に最も重要なメソッド。**

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| Start | 省略可 | Long | 開始文字位置（1始まり） |
| Length | 省略可 | Long | 取得する文字数 |

**動作仕様**:
- Start, Length 両方省略: 全文字を含む範囲
- Start のみ指定: 1文字分の範囲
- Length のみ指定: 先頭から Length 文字分
- Start がテキスト長を超える場合: 最後の文字から開始

```python
tr = shape.TextFrame.TextRange
tr.Text = "Hello World"

# 先頭5文字 "Hello" を取得
chars = tr.Characters(1, 5)
print(chars.Text)  # "Hello"

# 1文字目を取得
first_char = tr.Characters(1)
print(first_char.Text)  # "H"

# ★ 部分テキストの色変更（最重要パターン）
tr.Characters(1, 5).Font.Color.RGB = 0x0000FF  # "Hello" を赤に
tr.Characters(7, 5).Font.Color.RGB = 0xFF0000  # "World" を青に
```

#### Words(Start, Length)

単語単位でテキストの部分範囲を取得する。

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| Start | 省略可 | Long | 開始単語位置 |
| Length | 省略可 | Long | 取得する単語数 |

```python
tr = shape.TextFrame.TextRange
tr.Text = "Hello World Today"

# 最初の単語を取得
word1 = tr.Words(1)
print(word1.Text)  # "Hello "（末尾スペースを含む場合あり）

# 2番目の単語の色を変更
tr.Words(2).Font.Color.RGB = 0x00FF00  # 緑色
```

#### Lines(Start, Length)

表示上の行単位でテキストの部分範囲を取得する。

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| Start | 省略可 | Long | 開始行番号 |
| Length | 省略可 | Long | 取得する行数 |

**注意**: Lines はテキストフレームの幅による自動折り返し後の「表示行」を基準とする。つまり、シェイプの幅が変わると行の区切りも変わる。

```python
tr = shape.TextFrame.TextRange
# 1行目のテキストを取得
line1 = tr.Lines(1)
print(line1.Text)

# 2行目を太字に
tr.Lines(2).Font.Bold = -1  # msoTrue
```

#### Sentences(Start, Length)

文単位でテキストの部分範囲を取得する。文はピリオド（.）で区切られる。

```python
tr = shape.TextFrame.TextRange
tr.Text = "First sentence. Second sentence. Third."

# 2番目の文を取得
sent2 = tr.Sentences(2)
print(sent2.Text)  # " Second sentence."
```

#### Paragraphs(Start, Length)

段落単位でテキストの部分範囲を取得する。段落は改行（vbCr / \r）で区切られる。

```python
tr = shape.TextFrame.TextRange
tr.Text = "First paragraph\rSecond paragraph\rThird paragraph"

# 最初の段落
para1 = tr.Paragraphs(1)
print(para1.Text)

# 2番目の段落のインデントレベルを変更
tr.Paragraphs(2).IndentLevel = 2
```

#### Runs(Start, Length)

同一書式が連続する区間（ラン）単位でテキストの部分範囲を取得する。

**ランとは**: フォント属性が変わる箇所の間にある、同じ書式が適用された連続テキスト。例えば "This **bold** word" の場合、"This " (通常), "bold" (太字), " word" (通常) の3つのランになる。

```python
tr = shape.TextFrame.TextRange

# ラン数を取得
run_count = tr.Runs().Count
print(f"ラン数: {run_count}")

# 各ランの情報を表示
for i in range(1, run_count + 1):
    run = tr.Runs(i)
    print(f"Run {i}: '{run.Text}', Bold={run.Font.Bold}, Size={run.Font.Size}")

# 2番目のランが斜体なら太字にもする
run2 = tr.Runs(2)
if run2.Font.Italic:
    run2.Font.Bold = -1  # msoTrue
```

### 2.3 テキスト挿入メソッド

#### InsertBefore(NewText)

テキスト範囲の前にテキストを挿入する。挿入されたテキストの TextRange を返す。

```python
tr = shape.TextFrame.TextRange
tr.Text = "World"

# 先頭に挿入
inserted = tr.InsertBefore("Hello ")
print(tr.Text)  # "Hello World"
# inserted は "Hello " の部分を参照する TextRange
inserted.Font.Bold = -1  # 挿入したテキストを太字に
```

#### InsertAfter(NewText)

テキスト範囲の後にテキストを挿入する。

```python
tr = shape.TextFrame.TextRange
tr.Text = "Hello"

inserted = tr.InsertAfter(" World")
print(tr.Text)  # "Hello World"
inserted.Font.Color.RGB = 0x0000FF  # 挿入部分を赤に
```

#### InsertDateTime(DateTimeFormat, InsertAsField)

日時を挿入する。

```python
# ppDateTimeFigureOut = 14
tr.InsertDateTime(14)
```

#### InsertSlideNumber()

スライド番号を挿入する。

#### InsertSymbol(FontName, CharNumber, Unicode)

特殊記号を挿入する。

```python
# Wingdings フォントの記号を挿入
tr.InsertSymbol("Wingdings", 252)
```

### 2.4 検索・置換メソッド

#### Find(FindWhat, After, MatchCase, WholeWords)

テキスト範囲内で文字列を検索する。見つかった場合は TextRange を返し、見つからない場合は Nothing (None) を返す。

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| FindWhat | 必須 | String | 検索文字列 |
| After | 省略可 | Long | 検索開始位置 |
| MatchCase | 省略可 | MsoTriState | 大文字小文字を区別 |
| WholeWords | 省略可 | MsoTriState | 単語全体で一致 |

```python
tr = shape.TextFrame.TextRange
tr.Text = "Hello World Hello"

# "Hello" を検索
found = tr.Find("Hello")
if found is not None:
    print(f"Found at position {found.Start}, length {found.Length}")
    found.Font.Color.RGB = 0x0000FF  # 見つかった部分を赤に
```

#### Replace(FindWhat, ReplaceWhat, After, MatchCase, WholeWords)

テキスト範囲内で文字列を検索し、置換する。

```python
tr = shape.TextFrame.TextRange
tr.Text = "Hello World"

# "World" を "Python" に置換
result = tr.Replace("World", "Python")
print(tr.Text)  # "Hello Python"
```

### 2.5 その他のメソッド

| メソッド | 説明 |
|---|---|
| `Copy()` | テキスト範囲をクリップボードにコピー |
| `Cut()` | テキスト範囲を切り取ってクリップボードに配置 |
| `Delete()` | テキスト範囲を削除 |
| `Paste()` | クリップボードからテキストを貼り付け |
| `PasteSpecial(DataType, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link)` | 特殊貼り付け |
| `Select()` | テキスト範囲を選択状態にする |
| `ChangeCase(Type)` | 大文字/小文字の変更 |
| `AddPeriods()` | 各段落末にピリオドを追加（日本語対応: 句点） |
| `RemovePeriods()` | 各段落末のピリオドを削除 |
| `TrimText()` | 前後の空白を除去した TextRange を返す |
| `RotatedBounds(X1,Y1,X2,Y2,X3,Y3,X4,Y4)` | 回転を考慮したバウンディングボックスの4隅座標を返す |
| `LtrRun()` | テキスト方向を左から右に設定 |
| `RtlRun()` | テキスト方向を右から左に設定 |

---

## 3. Font - フォント設定

**階層パス**: `Shape.TextFrame.TextRange.Font` または `Shape.TextFrame.TextRange.Characters(n, m).Font`

Font オブジェクトは文字の書式設定を表す。TextRange の任意の部分範囲に対して個別に設定可能。

### 3.1 プロパティ一覧

| プロパティ | 型 | 説明 | 読み書き |
|---|---|---|---|
| `Name` | String | フォント名（例: "Arial", "游ゴシック"） | R/W |
| `NameAscii` | String | ASCII 文字用フォント名 | R/W |
| `NameFarEast` | String | 東アジア文字用フォント名 | R/W |
| `NameComplexScript` | String | 複合スクリプト文字用フォント名 | R/W |
| `NameOther` | String | その他の文字用フォント名 | R/W |
| `Size` | Single | フォントサイズ（ポイント単位） | R/W |
| `Bold` | MsoTriState | 太字 | R/W |
| `Italic` | MsoTriState | 斜体 | R/W |
| `Underline` | MsoTriState | 下線 | R/W |
| `Shadow` | MsoTriState | 影 | R/W |
| `Emboss` | MsoTriState | エンボス（浮き出し） | R/W |
| `Subscript` | MsoTriState | 下付き文字 | R/W |
| `Superscript` | MsoTriState | 上付き文字 | R/W |
| `BaselineOffset` | Single | ベースラインオフセット（-1.0〜1.0、下付き=負、上付き=正） | R/W |
| `Color` | ColorFormat | フォントの色（ColorFormat オブジェクト） | R/O |
| `AutoRotateNumbers` | MsoTriState | 縦書き時に数字を自動回転 | R/W |
| `Embedded` | MsoTriState | フォントがプレゼンテーションに埋め込まれているか | R/O |
| `Embeddable` | MsoTriState | フォントが埋め込み可能か | R/O |

**注意**: MsoTriState の値は `msoTrue = -1`, `msoFalse = 0`, `msoTriStateMixed = -2`

### 3.2 Color プロパティ（ColorFormat オブジェクト）

**階層パス**: `TextRange.Font.Color`

ColorFormat オブジェクトは色の設定を管理する。RGB 直接指定、テーマカラー、スキームカラーの3つの方式で色を設定できる。

| プロパティ | 型 | 説明 |
|---|---|---|
| `RGB` | Long | RGB 値で色を直接指定。**VBA の RGB() 関数値と同じ形式** |
| `SchemeColor` | PpColorSchemeIndex | レガシーカラースキームの色を指定（Office 2003 以前互換） |
| `ObjectThemeColor` | MsoThemeColorIndex | テーマカラーで色を指定（Office 2007+） |
| `TintAndShade` | Single | 色の濃淡調整（-1.0〜1.0） |
| `Brightness` | Single | 色の明るさ調整（-1.0〜1.0） |
| `Type` | MsoColorType | 色の種類（RGB / スキーム / テーマ）。読み取り専用 |

#### RGB 値の計算方法

**重要**: PowerPoint COM の RGB 値は `R + G*256 + B*256^2` の形式（BGR 順の整数エンコード）。VBA の `RGB(r, g, b)` 関数がこの形式を返す。

Python での実装:

```python
def RGB(r, g, b):
    """VBA 互換の RGB 関数"""
    return r + (g << 8) + (b << 16)

# 使用例
red = RGB(255, 0, 0)      # 255 (0x0000FF)
green = RGB(0, 255, 0)    # 65280 (0x00FF00)
blue = RGB(0, 0, 255)     # 16711680 (0xFF0000)
white = RGB(255, 255, 255) # 16777215 (0xFFFFFF)
```

**注意**: Python の 0xRRGGBB 表記とは逆順。`0x0000FF` は青ではなく赤（R=255, G=0, B=0）。

#### テーマカラー定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoThemeColorDark1` | 1 | 暗い色 1（通常は黒） |
| `msoThemeColorLight1` | 2 | 明るい色 1（通常は白） |
| `msoThemeColorDark2` | 3 | 暗い色 2 |
| `msoThemeColorLight2` | 4 | 明るい色 2 |
| `msoThemeColorAccent1` | 5 | アクセント色 1 |
| `msoThemeColorAccent2` | 6 | アクセント色 2 |
| `msoThemeColorAccent3` | 7 | アクセント色 3 |
| `msoThemeColorAccent4` | 8 | アクセント色 4 |
| `msoThemeColorAccent5` | 9 | アクセント色 5 |
| `msoThemeColorAccent6` | 10 | アクセント色 6 |
| `msoThemeColorHyperlink` | 11 | ハイパーリンクの色 |
| `msoThemeColorFollowedHyperlink` | 12 | 訪問済みハイパーリンクの色 |

#### SchemeColor 定数（レガシー）

| 定数 | 値 | 説明 |
|---|---|---|
| `ppBackground` | 1 | 背景色 |
| `ppForeground` | 2 | 前景色 |
| `ppShadow` | 3 | 影の色 |
| `ppTitle` | 4 | タイトルの色 |
| `ppFill` | 5 | 塗りつぶしの色 |
| `ppAccent1` | 6 | アクセント色 1 |
| `ppAccent2` | 7 | アクセント色 2 |
| `ppAccent3` | 8 | アクセント色 3 |

### 3.3 多言語フォント設定

PowerPoint は文字種別ごとに異なるフォントを適用できる。

```python
font = shape.TextFrame.TextRange.Font

# ASCII 文字用フォント（英数字）
font.NameAscii = "Arial"

# 東アジア文字用フォント（日本語・中国語・韓国語）
font.NameFarEast = "游ゴシック"

# 複合スクリプト文字用フォント（アラビア語・ヘブライ語など）
font.NameComplexScript = "Arial"

# Name を設定すると全ての NameXxx が上書きされる場合がある点に注意
font.Name = "Meiryo"
```

### 3.4 Python win32com 総合例

```python
import win32com.client

def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True
pres = ppt.Presentations.Add()
slide = pres.Slides.Add(1, 1)  # ppLayoutTitle

shape = slide.Shapes(1)
tf = shape.TextFrame
tr = tf.TextRange

# タイトルテキストを設定
tr.Text = "Volcano Coffee"

# フォント全体の設定
with_font = tr.Font
with_font.Name = "Palatino"
with_font.Size = 48
with_font.Bold = -1   # msoTrue
with_font.Italic = -1  # msoTrue
with_font.Color.RGB = RGB(0, 0, 255)  # 青

# 下付き文字の例
tr2 = shape.TextFrame.TextRange
tr2.Text = "H2O"
tr2.Characters(2, 1).Font.BaselineOffset = -0.2  # "2" を下付きに
```

---

## 4. ParagraphFormat - 段落書式

**階層パス**: `Shape.TextFrame.TextRange.ParagraphFormat` または `Shape.TextFrame.TextRange.Paragraphs(n).ParagraphFormat`

### 4.1 プロパティ一覧

| プロパティ | 型 | 説明 |
|---|---|---|
| `Alignment` | PpParagraphAlignment | テキストの配置 |
| `SpaceBefore` | Single | 段落前のスペース（ポイントまたは行倍率） |
| `SpaceAfter` | Single | 段落後のスペース（ポイントまたは行倍率） |
| `SpaceWithin` | Single | 段落内の行間（ポイントまたは行倍率） |
| `LineRuleBefore` | MsoTriState | SpaceBefore を行数倍率として扱うか |
| `LineRuleAfter` | MsoTriState | SpaceAfter を行数倍率として扱うか |
| `LineRuleWithin` | MsoTriState | SpaceWithin を行数倍率として扱うか |
| `Bullet` | BulletFormat | 箇条書き設定オブジェクト |
| `TextDirection` | PpDirection | テキストの方向（LTR/RTL） |
| `BaseLineAlignment` | PpBaselineAlignment | ベースラインの配置 |
| `FarEastLineBreakControl` | MsoTriState | 東アジア文字の禁則処理 |
| `HangingPunctuation` | MsoTriState | 句読点のぶら下げ |
| `WordWrap` | MsoTriState | ワードラップ |

**注意**: `IndentLevel`、`FirstLineIndent`、`LeftIndent`、`RightIndent` は TextRange のプロパティとして、または Ruler オブジェクト経由でアクセスする。

#### Alignment の定数値

| 定数 | 値 | 説明 |
|---|---|---|
| `ppAlignLeft` | 1 | 左揃え |
| `ppAlignCenter` | 2 | 中央揃え |
| `ppAlignRight` | 3 | 右揃え |
| `ppAlignJustify` | 4 | 両端揃え |
| `ppAlignDistribute` | 5 | 均等割り付け |
| `ppAlignmentMixed` | -2 | 混在 |

#### 行間設定の仕組み

`LineRuleWithin` が `msoTrue (-1)` の場合、`SpaceWithin` は行数の倍率（例: 1.5 = 1.5行間隔）。
`LineRuleWithin` が `msoFalse (0)` の場合、`SpaceWithin` はポイント値。

```python
pf = shape.TextFrame.TextRange.ParagraphFormat

# 行間を 1.5 行に設定
pf.LineRuleWithin = -1  # msoTrue（行数倍率モード）
pf.SpaceWithin = 1.5

# 段落前のスペースを 12 ポイントに設定
pf.LineRuleBefore = 0  # msoFalse（ポイント値モード）
pf.SpaceBefore = 12

# 段落後のスペースを 6 ポイントに設定
pf.LineRuleAfter = 0
pf.SpaceAfter = 6
```

### 4.2 Bullet - 箇条書き設定

**階層パス**: `TextRange.ParagraphFormat.Bullet`

BulletFormat オブジェクトは段落の箇条書き書式を管理する。

#### BulletFormat プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `Visible` | MsoTriState | 箇条書きマークの表示/非表示 |
| `Type` | PpBulletType | 箇条書きの種類 |
| `Character` | Long | 箇条書き記号の Unicode 文字コード |
| `Font` | Font | 箇条書き記号のフォント設定 |
| `RelativeSize` | Single | 箇条書き記号のテキストに対する相対サイズ（0.25〜4.0） |
| `StartValue` | Long | 番号付き箇条書きの開始値（1〜32767） |
| `Number` | Long | 現在の段落の番号値。読み取り専用 |
| `Style` | PpNumberedBulletStyle | 番号付き箇条書きのスタイル |
| `UseTextColor` | MsoTriState | テキストの色を箇条書き記号に使用 |
| `UseTextFont` | MsoTriState | テキストのフォントを箇条書き記号に使用 |

#### BulletFormat メソッド

| メソッド | 説明 |
|---|---|
| `Picture(Picture)` | 画像を箇条書きマークとして設定 |

#### PpBulletType 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `ppBulletNone` | 0 | 箇条書きなし |
| `ppBulletUnnumbered` | 1 | 記号付き箇条書き |
| `ppBulletNumbered` | 2 | 番号付き箇条書き |
| `ppBulletPicture` | 3 | 画像付き箇条書き |
| `ppBulletMixed` | -2 | 混在 |

#### PpNumberedBulletStyle 定数（一部）

| 定数 | 値 | 説明 |
|---|---|---|
| `ppBulletArabicParenRight` | 2 | 1) 2) 3) |
| `ppBulletArabicPeriod` | 3 | 1. 2. 3. |
| `ppBulletArabicParenBoth` | 12 | (1) (2) (3) |
| `ppBulletRomanUCPeriod` | 4 | I. II. III. |
| `ppBulletRomanLCPeriod` | 5 | i. ii. iii. |
| `ppBulletAlphaUCPeriod` | 6 | A. B. C. |
| `ppBulletAlphaLCPeriod` | 7 | a. b. c. |

#### Python win32com 例

```python
def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

tr = shape.TextFrame.TextRange
tr.Text = "First item\rSecond item\rThird item"

# 全段落に記号付き箇条書きを設定
bullet = tr.ParagraphFormat.Bullet
bullet.Visible = -1  # msoTrue
bullet.Type = 1       # ppBulletUnnumbered
bullet.Character = 8226  # Unicode bullet character "•"
bullet.RelativeSize = 1.25
bullet.Font.Color.RGB = RGB(255, 0, 0)  # 赤い箇条書きマーク
bullet.Font.Name = "Arial"

# 番号付き箇条書きに変更
bullet.Type = 2  # ppBulletNumbered
bullet.Style = 3  # ppBulletArabicPeriod: "1. 2. 3."
bullet.StartValue = 1

# 特定の段落のみ箇条書きレベルを変更
tr.Paragraphs(2).IndentLevel = 2  # 2番目の段落をレベル2に
```

### 4.3 インデント設定

TextRange の `IndentLevel` プロパティと、Ruler オブジェクト経由のインデント設定がある。

```python
tr = shape.TextFrame.TextRange

# 段落ごとのインデントレベル（1〜9）
tr.Paragraphs(1).IndentLevel = 1
tr.Paragraphs(2).IndentLevel = 2

# Ruler 経由でのインデント設定（TextFrame のプロパティ）
ruler = shape.TextFrame.Ruler

# レベル1のインデント設定
ruler.Levels(1).FirstMargin = 0     # 1行目のインデント（ポイント）
ruler.Levels(1).LeftMargin = 36     # 左マージン（ポイント）

# レベル2のインデント設定
ruler.Levels(2).FirstMargin = 36
ruler.Levels(2).LeftMargin = 72
```

### 4.4 TabStops（タブ位置）

```python
ruler = shape.TextFrame.Ruler
tab_stops = ruler.TabStops

# タブ位置を追加
tab_stops.Add(1, 72)    # ppTabStopLeft=1, 72ポイント位置
tab_stops.Add(2, 216)   # ppTabStopCenter=2, 216ポイント位置
tab_stops.Add(3, 360)   # ppTabStopRight=3, 360ポイント位置
```

---

## 5. Fill - 塗りつぶし

**階層パス**: `Shape.Fill`

FillFormat オブジェクトはシェイプの塗りつぶし書式を管理する。単色、グラデーション、パターン、テクスチャ、画像など様々な塗りつぶし種類をサポート。

### 5.1 プロパティ一覧

| プロパティ | 型 | 説明 |
|---|---|---|
| `ForeColor` | ColorFormat | 前景色（単色塗りつぶしの色、グラデーションの開始色） |
| `BackColor` | ColorFormat | 背景色（パターンの背景色、グラデーションの終了色） |
| `Type` | MsoFillType | 塗りつぶしの種類。読み取り専用 |
| `Transparency` | Single | 透明度（0.0=不透明 〜 1.0=完全透明） |
| `Visible` | MsoTriState | 塗りつぶしの表示/非表示 |
| `GradientAngle` | Single | グラデーションの角度（度） |
| `GradientColorType` | MsoGradientColorType | グラデーションの色タイプ |
| `GradientDegree` | Single | 1色グラデーションの濃淡（0.0〜1.0） |
| `GradientStops` | GradientStops | グラデーション停止点コレクション |
| `GradientStyle` | MsoGradientStyle | グラデーションのスタイル |
| `GradientVariant` | Long | グラデーションのバリアント（1〜4） |
| `Pattern` | MsoPatternType | パターンの種類 |
| `PresetGradientType` | MsoPresetGradientType | プリセットグラデーションの種類 |
| `PresetTexture` | MsoPresetTexture | プリセットテクスチャの種類 |
| `TextureName` | String | カスタムテクスチャのファイル名 |
| `TextureType` | MsoTextureType | テクスチャの種類 |
| `TextureAlignment` | MsoTextureAlignment | テクスチャの配置 |
| `TextureHorizontalScale` | Single | テクスチャの水平スケール |
| `TextureVerticalScale` | Single | テクスチャの垂直スケール |
| `TextureOffsetX` | Single | テクスチャの X オフセット |
| `TextureOffsetY` | Single | テクスチャの Y オフセット |
| `TextureTile` | MsoTriState | テクスチャのタイリング |
| `RotateWithObject` | MsoTriState | オブジェクトと一緒に回転するか |
| `PictureEffects` | PictureEffects | ピクチャ効果コレクション |

#### MsoFillType 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoFillSolid` | 1 | 単色 |
| `msoFillPatterned` | 2 | パターン |
| `msoFillGradient` | 3 | グラデーション |
| `msoFillTextured` | 4 | テクスチャ |
| `msoFillPicture` | 6 | 画像 |
| `msoFillBackground` | 5 | 背景 |
| `msoFillMixed` | -2 | 混在 |

### 5.2 メソッド一覧

| メソッド | 説明 |
|---|---|
| `Solid()` | 単色塗りつぶしに設定 |
| `Patterned(Pattern)` | パターン塗りつぶしに設定 |
| `OneColorGradient(Style, Variant, Degree)` | 1色グラデーションに設定 |
| `TwoColorGradient(Style, Variant)` | 2色グラデーションに設定 |
| `PresetGradient(Style, Variant, PresetGradientType)` | プリセットグラデーションに設定 |
| `PresetTextured(PresetTexture)` | プリセットテクスチャに設定 |
| `UserPicture(PictureFile)` | ユーザー画像で塗りつぶし |
| `UserTextured(TextureFile)` | ユーザーテクスチャでタイル塗りつぶし |
| `Background()` | スライド背景に合わせる |

### 5.3 単色塗りつぶし

```python
def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

shape = slide.Shapes(1)
fill = shape.Fill

# 単色塗りつぶし
fill.Solid()
fill.ForeColor.RGB = RGB(255, 0, 0)  # 赤
fill.Transparency = 0  # 不透明

# テーマカラーで設定
fill.Solid()
fill.ForeColor.ObjectThemeColor = 5  # msoThemeColorAccent1
fill.ForeColor.Brightness = 0.4      # 明るさ調整

# 塗りつぶし非表示
fill.Visible = 0  # msoFalse
```

### 5.4 グラデーション

#### 2色グラデーション

```python
fill = shape.Fill
fill.ForeColor.RGB = RGB(128, 0, 0)      # 前景色: 暗い赤
fill.BackColor.RGB = RGB(170, 170, 170)   # 背景色: グレー
fill.TwoColorGradient(1, 1)               # msoGradientHorizontal, Variant 1
```

#### MsoGradientStyle 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoGradientHorizontal` | 1 | 水平 |
| `msoGradientVertical` | 2 | 垂直 |
| `msoGradientDiagonalUp` | 3 | 対角線（右上がり） |
| `msoGradientDiagonalDown` | 4 | 対角線（右下がり） |
| `msoGradientFromCorner` | 5 | 角から |
| `msoGradientFromCenter` | 7 | 中心から |
| `msoGradientFromTitle` | 6 | タイトルから |

#### 1色グラデーション

```python
fill = shape.Fill
fill.ForeColor.RGB = RGB(0, 0, 255)
fill.OneColorGradient(2, 1, 0.5)  # msoGradientVertical, Variant 1, Degree 0.5
```

#### プリセットグラデーション

```python
fill = shape.Fill
# msoGradientDiagonalDown=4, Variant=1, msoGradientOcean=7
fill.PresetGradient(4, 1, 7)
```

#### GradientStops（グラデーション停止点の詳細制御）

既にグラデーションが適用されたシェイプに対して、グラデーション停止点を追加・変更する。

```python
fill = shape.Fill

# まずグラデーションを適用
fill.TwoColorGradient(1, 1)

# グラデーション停止点を追加
# Insert(RGB_Color, Position, Transparency=0, Index=-1)
gradient_stops = fill.GradientStops
gradient_stops.Insert(RGB(255, 0, 255), 0.5)  # 50%位置にマゼンタを追加

# 既存の停止点を変更
stop = gradient_stops(1)
stop.Color.RGB = RGB(255, 255, 0)  # 黄色に変更
stop.Position = 0.3                 # 30%位置に移動
stop.Transparency = 0.2             # 20%透明

# 停止点の数を確認
print(f"停止点数: {gradient_stops.Count}")

# 停止点を削除
gradient_stops.Delete(2)  # 2番目の停止点を削除
```

### 5.5 パターン塗りつぶし

```python
fill = shape.Fill
fill.Patterned(48)  # msoPatternDiagonalBrick = 48 など

fill.ForeColor.RGB = RGB(0, 0, 128)  # パターンの前景色
fill.BackColor.RGB = RGB(255, 255, 255)  # パターンの背景色
```

#### MsoPatternType 定数（一部）

| 定数 | 値 | 説明 |
|---|---|---|
| `msoPattern5Percent` | 1 | 5% |
| `msoPattern10Percent` | 2 | 10% |
| `msoPattern25Percent` | 3 | 25% |
| `msoPattern50Percent` | 4 | 50% |
| `msoPattern75Percent` | 5 | 75% |
| `msoPatternHorizontal` | 6 | 横線 |
| `msoPatternVertical` | 7 | 縦線 |
| `msoPatternDiagonalBrick` | 48 | 対角レンガ |
| `msoPatternCross` | 11 | 十字 |
| `msoPatternChecks` | 12 | チェック |

### 5.6 テクスチャ・画像塗りつぶし

```python
fill = shape.Fill

# プリセットテクスチャ
fill.PresetTextured(1)  # msoTexturePapyrus など

# ユーザー画像（1枚の大きな画像）
fill.UserPicture(r"C:\path\to\image.jpg")

# ユーザーテクスチャ（タイル状に繰り返し）
fill.UserTextured(r"C:\path\to\texture.jpg")
```

---

## 6. Line - 枠線

**階層パス**: `Shape.Line`

LineFormat オブジェクトはシェイプの枠線（境界線）や、線シェイプの書式を管理する。

### 6.1 プロパティ一覧

| プロパティ | 型 | 説明 |
|---|---|---|
| `ForeColor` | ColorFormat | 線の前景色（単色線の場合はこれが線の色） |
| `BackColor` | ColorFormat | 線の背景色（パターン線の場合に使用） |
| `Weight` | Single | 線の太さ（ポイント単位） |
| `DashStyle` | MsoLineDashStyle | 破線スタイル |
| `Style` | MsoLineStyle | 線のスタイル（単線、二重線など） |
| `Transparency` | Single | 線の透明度（0.0〜1.0） |
| `Visible` | MsoTriState | 線の表示/非表示 |
| `Pattern` | MsoPatternType | 線のパターン |
| `BeginArrowheadStyle` | MsoArrowheadStyle | 始点の矢印スタイル |
| `BeginArrowheadLength` | MsoArrowheadLength | 始点の矢印の長さ |
| `BeginArrowheadWidth` | MsoArrowheadWidth | 始点の矢印の幅 |
| `EndArrowheadStyle` | MsoArrowheadStyle | 終点の矢印スタイル |
| `EndArrowheadLength` | MsoArrowheadLength | 終点の矢印の長さ |
| `EndArrowheadWidth` | MsoArrowheadWidth | 終点の矢印の幅 |
| `InsetPen` | MsoTriState | 線をシェイプ内側に描画するか |

#### MsoLineDashStyle 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoLineSolid` | 1 | 実線 |
| `msoLineDash` | 4 | 破線 |
| `msoLineDot` | 3 | 点線 |
| `msoLineDashDot` | 5 | 一点鎖線 |
| `msoLineDashDotDot` | 6 | 二点鎖線 |
| `msoLineRoundDot` | 2 | 丸点線 |
| `msoLineLongDash` | 7 | 長い破線 |
| `msoLineLongDashDot` | 8 | 長い一点鎖線 |
| `msoLineSquareDot` | 2 | 角点線 |
| `msoLineDashStyleMixed` | -2 | 混在 |

#### MsoLineStyle 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoLineSingle` | 1 | 単線 |
| `msoLineThinThin` | 2 | 二重線（細-細） |
| `msoLineThinThick` | 3 | 二重線（細-太） |
| `msoLineThickThin` | 4 | 二重線（太-細） |
| `msoLineThickBetweenThin` | 5 | 三重線（細-太-細） |

#### MsoArrowheadStyle 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoArrowheadNone` | 1 | 矢印なし |
| `msoArrowheadTriangle` | 2 | 三角形 |
| `msoArrowheadOpen` | 3 | 開いた矢印 |
| `msoArrowheadStealth` | 4 | ステルス矢印 |
| `msoArrowheadDiamond` | 5 | ひし形 |
| `msoArrowheadOval` | 6 | 楕円 |

#### MsoArrowheadLength 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoArrowheadShort` | 1 | 短い |
| `msoArrowheadLengthMedium` | 2 | 中間 |
| `msoArrowheadLong` | 3 | 長い |

#### MsoArrowheadWidth 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoArrowheadNarrow` | 1 | 狭い |
| `msoArrowheadWidthMedium` | 2 | 中間 |
| `msoArrowheadWide` | 3 | 広い |

### 6.2 Python win32com 例

```python
def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

# シェイプの枠線設定
shape = slide.Shapes(1)
line = shape.Line

line.Visible = -1  # msoTrue
line.ForeColor.RGB = RGB(0, 0, 128)   # 暗い青
line.Weight = 2.5                       # 2.5ポイント
line.DashStyle = 1                      # msoLineSolid（実線）
line.Style = 1                          # msoLineSingle（単線）
line.Transparency = 0                   # 不透明

# 枠線を非表示にする
shape.Line.Visible = 0  # msoFalse

# 線シェイプの矢印設定
line_shape = slide.Shapes.AddLine(100, 100, 300, 200)
with_line = line_shape.Line

with_line.DashStyle = 6                    # msoLineDashDotDot
with_line.ForeColor.RGB = RGB(50, 0, 128)  # 紫
with_line.Weight = 2

# 始点: 短い狭い楕円
with_line.BeginArrowheadLength = 1   # msoArrowheadShort
with_line.BeginArrowheadStyle = 6    # msoArrowheadOval
with_line.BeginArrowheadWidth = 1    # msoArrowheadNarrow

# 終点: 長い広い三角形
with_line.EndArrowheadLength = 3     # msoArrowheadLong
with_line.EndArrowheadStyle = 2      # msoArrowheadTriangle
with_line.EndArrowheadWidth = 3      # msoArrowheadWide
```

---

## 7. Shadow / Glow / Reflection / SoftEdge - 効果

### 7.1 Shadow（影）

**階層パス**: `Shape.Shadow`

ShadowFormat オブジェクトはシェイプの影効果を管理する。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `Type` | MsoShadowType | 影の種類（プリセット番号） |
| `ForeColor` | ColorFormat | 影の色 |
| `OffsetX` | Single | 影の水平オフセット（ポイント、正=右方向） |
| `OffsetY` | Single | 影の垂直オフセット（ポイント、正=下方向） |
| `Blur` | Single | 影のぼかし半径（ポイント） |
| `Transparency` | Single | 影の透明度（0.0=不透明 〜 1.0=完全透明） |
| `Size` | Single | 影のサイズ（シェイプサイズに対する百分率、0〜200） |
| `Style` | MsoShadowStyle | 影のスタイル（内側/外側） |
| `Visible` | MsoTriState | 影の表示/非表示 |
| `Obscured` | MsoTriState | 影がシェイプに隠れるか |
| `RotateWithShape` | MsoTriState | シェイプの回転に追随するか |

#### MsoShadowType 定数（一部）

| 定数 | 値 | 説明 |
|---|---|---|
| `msoShadow1` 〜 `msoShadow20` | 1〜20 | プリセット影 1〜20 |
| `msoShadow21` 〜 `msoShadow43` | 21〜43 | 追加プリセット影 |

#### MsoShadowStyle 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoShadowStyleInnerShadow` | 1 | 内側の影 |
| `msoShadowStyleOuterShadow` | 2 | 外側の影 |
| `msoShadowStyleMixed` | -2 | 混在 |

#### メソッド

| メソッド | 説明 |
|---|---|
| `IncrementOffsetX(Increment)` | 水平オフセットを指定値分変更 |
| `IncrementOffsetY(Increment)` | 垂直オフセットを指定値分変更 |

#### Python win32com 例

```python
def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

shape = slide.Shapes(1)
shadow = shape.Shadow

# プリセット影を適用
shadow.Type = 17  # msoShadow17

# カスタム影設定
shadow.Visible = -1  # msoTrue
shadow.ForeColor.RGB = RGB(0, 0, 128)  # 暗い青の影
shadow.OffsetX = 5    # 右に5ポイント
shadow.OffsetY = 3    # 下に3ポイント
shadow.Blur = 8       # ぼかし8ポイント
shadow.Transparency = 0.5  # 50%透明
shadow.Size = 100      # サイズ100%（元のサイズ）
shadow.Style = 2       # msoShadowStyleOuterShadow（外側の影）
shadow.RotateWithShape = -1  # msoTrue
```

### 7.2 Glow（光彩）

**階層パス**: `Shape.Glow`

GlowFormat オブジェクトはシェイプの光彩（グロー）効果を管理する。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `Color` | ColorFormat | 光彩の色 |
| `Radius` | Single | 光彩の半径（ポイント） |
| `Transparency` | Single | 光彩の透明度（0.0〜1.0） |

**注意**: Glow は Shape オブジェクトのプロパティとして直接アクセスする。`Type` プロパティは Glow オブジェクトには直接存在しない。プリセットのグロー効果は存在しない（手動で Color, Radius, Transparency を設定する）。

#### Python win32com 例

```python
def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

shape = slide.Shapes(1)
glow = shape.Glow

# 光彩効果を設定
glow.Color.RGB = RGB(255, 215, 0)  # ゴールド
glow.Radius = 10                    # 半径10ポイント
glow.Transparency = 0.5             # 50%透明

# テーマカラーで光彩を設定
glow.Color.ObjectThemeColor = 5  # msoThemeColorAccent1
glow.Radius = 15
glow.Transparency = 0.6

# 光彩を除去（半径を0にする）
glow.Radius = 0
```

### 7.3 Reflection（反射）

**階層パス**: `Shape.Reflection`

ReflectionFormat オブジェクトはシェイプの反射効果を管理する。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `Type` | MsoReflectionType | プリセット反射の種類 |
| `Blur` | Single | 反射のぼかし（ポイント） |
| `Offset` | Single | 反射のオフセット距離（ポイント） |
| `Size` | Single | 反射のサイズ（百分率） |
| `Transparency` | Single | 反射の透明度（0.0〜1.0） |

#### MsoReflectionType 定数（一部）

| 定数 | 値 | 説明 |
|---|---|---|
| `msoReflectionTypeNone` | 0 | 反射なし |
| `msoReflectionType1` | 1 | プリセット反射 1（密接な反射、接触） |
| `msoReflectionType2` | 2 | プリセット反射 2 |
| `msoReflectionType3` | 3 | プリセット反射 3 |
| ... | ... | ... |
| `msoReflectionType9` | 9 | プリセット反射 9（完全反射、オフセット大） |

#### Python win32com 例

```python
shape = slide.Shapes(1)
reflection = shape.Reflection

# プリセット反射を適用
reflection.Type = 1  # msoReflectionType1

# カスタム反射設定
reflection.Blur = 0.5
reflection.Offset = 2.4
reflection.Size = 100
reflection.Transparency = 0.5

# 反射を除去
reflection.Type = 0  # msoReflectionTypeNone
```

### 7.4 SoftEdge（ぼかし）

**階層パス**: `Shape.SoftEdge`

SoftEdgeFormat オブジェクトはシェイプのエッジぼかし効果を管理する。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `Type` | MsoSoftEdgeType | プリセットぼかしの種類 |
| `Radius` | Single | ぼかしの半径（ポイント） |

#### MsoSoftEdgeType 定数

| 定数 | 値 | 説明 |
|---|---|---|
| `msoSoftEdgeTypeNone` | 0 | ぼかしなし |
| `msoSoftEdgeType1` | 1 | 1ポイントのぼかし |
| `msoSoftEdgeType2` | 2 | 2.5ポイントのぼかし |
| `msoSoftEdgeType3` | 3 | 5ポイントのぼかし |
| `msoSoftEdgeType4` | 4 | 10ポイントのぼかし |
| `msoSoftEdgeType5` | 5 | 25ポイントのぼかし |
| `msoSoftEdgeType6` | 6 | 50ポイントのぼかし |

#### Python win32com 例

```python
shape = slide.Shapes(1)
soft_edge = shape.SoftEdge

# プリセットぼかしを適用
soft_edge.Type = 3  # msoSoftEdgeType3（5ポイント）

# カスタム半径を直接設定
soft_edge.Radius = 8  # 8ポイントのぼかし

# ぼかしを除去
soft_edge.Type = 0  # msoSoftEdgeTypeNone
```

---

## 8. ThreeDFormat - 3D 効果

**階層パス**: `Shape.ThreeD`

ThreeDFormat オブジェクトはシェイプの 3D（立体）効果を管理する。ベベル、押し出し、回転、照明、マテリアルの設定が可能。

### 8.1 プロパティ一覧

| プロパティ | 型 | 説明 |
|---|---|---|
| `BevelTopType` | MsoBevelType | 上面ベベルの種類 |
| `BevelTopDepth` | Single | 上面ベベルの深さ |
| `BevelTopInset` | Single | 上面ベベルのインセット |
| `BevelBottomType` | MsoBevelType | 底面ベベルの種類 |
| `BevelBottomDepth` | Single | 底面ベベルの深さ |
| `BevelBottomInset` | Single | 底面ベベルのインセット |
| `Depth` | Single | 押し出しの深さ（ポイント） |
| `ExtrusionColor` | ColorFormat | 押し出し部分の色 |
| `ExtrusionColorType` | MsoExtrusionColorType | 押し出し色の種類 |
| `RotationX` | Single | X軸回転（-90〜90度、正=上向き） |
| `RotationY` | Single | Y軸回転（-90〜90度、正=左向き） |
| `RotationZ` | Single | Z軸回転（度） |
| `PresetLightingDirection` | MsoPresetLightingDirection | 照明の方向 |
| `PresetLightingSoftness` | MsoPresetLightingSoftness | 照明の柔らかさ |
| `PresetMaterial` | MsoPresetMaterial | マテリアル（質感）の種類 |
| `PresetThreeDFormat` | MsoPresetThreeDFormat | プリセット 3D 書式 |
| `PresetExtrusionDirection` | MsoPresetExtrusionDirection | 押し出しの方向 |
| `Perspective` | MsoTriState | パースペクティブ（遠近法）の有無 |
| `FieldOfView` | Single | 視野角 |
| `LightAngle` | Single | 光源の角度 |
| `ContourColor` | ColorFormat | 輪郭線の色 |
| `ContourWidth` | Single | 輪郭線の幅 |
| `Visible` | MsoTriState | 3D 効果の表示/非表示 |
| `ProjectText` | MsoTriState | テキストの投影 |

#### MsoBevelType 定数（一部）

| 定数 | 値 | 説明 |
|---|---|---|
| `msoBevelNone` | 1 | ベベルなし |
| `msoBevelCircle` | 3 | 円形 |
| `msoBevelRelaxedInset` | 2 | リラックスインセット |
| `msoBevelCross` | 4 | 十字 |
| `msoBevelSlope` | 5 | スロープ |
| `msoBevelAngle` | 6 | 角度 |
| `msoBevelSoftRound` | 7 | ソフトラウンド |
| `msoBevelConvex` | 8 | 凸 |
| `msoBevelCoolSlant` | 9 | クールスラント |
| `msoBevelDivot` | 10 | ディボット |
| `msoBevelRiblet` | 11 | リブレット |
| `msoBevelHardEdge` | 12 | ハードエッジ |
| `msoBevelArtDeco` | 13 | アールデコ |

#### MsoPresetMaterial 定数（一部）

| 定数 | 値 | 説明 |
|---|---|---|
| `msoMaterialMatte` | 1 | マット |
| `msoMaterialMatte2` | 5 | マット2 |
| `msoMaterialPlastic` | 2 | プラスチック |
| `msoMaterialMetal` | 3 | 金属 |
| `msoMaterialMetal2` | 6 | 金属2 |
| `msoMaterialWireFrame` | 4 | ワイヤフレーム |
| `msoMaterialDarkEdge` | 7 | ダークエッジ |
| `msoMaterialSoftEdge` | 8 | ソフトエッジ |
| `msoMaterialFlat` | 9 | フラット |
| `msoMaterialSoftMetal` | 10 | ソフトメタル |
| `msoMaterialPowder` | 11 | パウダー |
| `msoMaterialWarmMatte` | 12 | ウォームマット |
| `msoMaterialTranslucentPowder` | 13 | 半透明パウダー |
| `msoMaterialClear` | 14 | クリア |

#### MsoPresetLightingDirection 定数（一部）

| 定数 | 値 | 説明 |
|---|---|---|
| `msoLightingTop` | 13 | 上 |
| `msoLightingBottom` | 14 | 下 |
| `msoLightingLeft` | 15 | 左 |
| `msoLightingRight` | 16 | 右 |
| `msoLightingTopLeft` | 7 | 左上 |
| `msoLightingTopRight` | 8 | 右上 |
| `msoLightingBottomLeft` | 9 | 左下 |
| `msoLightingBottomRight` | 10 | 右下 |

### 8.2 メソッド

| メソッド | 説明 |
|---|---|
| `IncrementRotationX(Increment)` | X軸回転を指定角度分変更 |
| `IncrementRotationY(Increment)` | Y軸回転を指定角度分変更 |
| `IncrementRotationZ(Increment)` | Z軸回転を指定角度分変更 |
| `ResetRotation()` | 回転をリセット |
| `SetThreeDFormat(PresetFormat)` | プリセット 3D 書式を適用 |
| `SetExtrusionDirection(Direction)` | 押し出しの方向を設定 |

### 8.3 Python win32com 例

```python
def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

shape = slide.Shapes(1)
three_d = shape.ThreeD

# ベベル効果
three_d.BevelTopType = 3       # msoBevelCircle
three_d.BevelTopDepth = 6      # 深さ6ポイント
three_d.BevelTopInset = 6      # インセット6ポイント
three_d.BevelBottomType = 1    # msoBevelNone

# 押し出し
three_d.Depth = 36             # 36ポイントの押し出し
three_d.ExtrusionColor.RGB = RGB(128, 128, 128)  # グレー

# 3D 回転
three_d.RotationX = 15         # X軸15度回転
three_d.RotationY = -20        # Y軸-20度回転

# 照明とマテリアル
three_d.PresetLightingDirection = 7  # msoLightingTopLeft
three_d.PresetMaterial = 2           # msoMaterialPlastic

# プリセット 3D フォーマットの適用
three_d.SetThreeDFormat(1)  # msoThreeD1
```

---

## 9. 部分テキスト色変更の詳細実装例（重要）

ここでは MCP サーバーの最重要ユースケースである「1つのテキストボックス内でテキストの一部だけ色を変える」方法を詳しく解説する。

### 9.1 基本原理

PowerPoint COM では、テキストの部分書式変更は以下の流れで行う:

1. `TextFrame.TextRange` でテキスト全体の TextRange を取得
2. `Characters(Start, Length)` で変更したい部分の TextRange を取得
3. 取得した部分 TextRange の `.Font` プロパティで書式を変更

**この操作は VBA / COM の根本機能であり、python-pptx ライブラリとは異なり「Run」を事前に分割する必要がない。** Characters() で任意の位置・長さの部分を指定するだけで、PowerPoint が内部的にランを自動分割して書式を適用する。

### 9.2 例1: 虹色テキスト

```python
import win32com.client

def RGB(r, g, b):
    return r + (g << 8) + (b << 16)

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True
pres = ppt.Presentations.Add()
slide = pres.Slides.Add(1, 12)  # ppLayoutBlank

# テキストボックスを追加
# msoTextOrientationHorizontal = 1
textbox = slide.Shapes.AddTextbox(1, 100, 100, 500, 100)
tr = textbox.TextFrame.TextRange
tr.Text = "RAINBOW"
tr.Font.Size = 48

# 各文字に異なる色を設定
colors = [
    RGB(255, 0, 0),     # R - 赤
    RGB(255, 165, 0),   # A - オレンジ
    RGB(255, 255, 0),   # I - 黄
    RGB(0, 128, 0),     # N - 緑
    RGB(0, 0, 255),     # B - 青
    RGB(75, 0, 130),    # O - 藍
    RGB(238, 130, 238), # W - 紫
]

for i, color in enumerate(colors):
    tr.Characters(i + 1, 1).Font.Color.RGB = color
```

### 9.3 例2: キーワードのハイライト

```python
def highlight_keywords(text_range, keyword, color):
    """テキスト内のキーワードを指定色にハイライトする"""
    text = text_range.Text
    start = 0
    while True:
        pos = text.find(keyword, start)
        if pos == -1:
            break
        # Characters は 1-based index
        text_range.Characters(pos + 1, len(keyword)).Font.Color.RGB = color
        start = pos + len(keyword)

# 使用例
tr = textbox.TextFrame.TextRange
tr.Text = "Python is great. Python is powerful. I love Python."
tr.Font.Size = 24
tr.Font.Color.RGB = RGB(0, 0, 0)  # デフォルトは黒

# "Python" を赤色にハイライト
highlight_keywords(tr, "Python", RGB(255, 0, 0))
```

### 9.4 例3: 複合書式（色+太字+サイズ変更）

```python
tr = textbox.TextFrame.TextRange
tr.Text = "重要: この文書は機密です。取扱注意。"
tr.Font.Size = 18
tr.Font.Name = "游ゴシック"
tr.Font.Color.RGB = RGB(0, 0, 0)  # 全体を黒

# "重要:" を赤太字に
chars_important = tr.Characters(1, 3)  # "重要:"
chars_important.Font.Color.RGB = RGB(255, 0, 0)
chars_important.Font.Bold = -1  # msoTrue
chars_important.Font.Size = 24

# "機密" を青下線に
text = tr.Text
pos = text.find("機密")
if pos >= 0:
    chars_secret = tr.Characters(pos + 1, 2)
    chars_secret.Font.Color.RGB = RGB(0, 0, 255)
    chars_secret.Font.Underline = -1  # msoTrue

# "取扱注意" をオレンジ斜体に
pos2 = text.find("取扱注意")
if pos2 >= 0:
    chars_caution = tr.Characters(pos2 + 1, 4)
    chars_caution.Font.Color.RGB = RGB(255, 165, 0)
    chars_caution.Font.Italic = -1
```

### 9.5 例4: Find を使った部分書式変更

```python
tr = textbox.TextFrame.TextRange
tr.Text = "Error: File not found. Warning: Disk space low."
tr.Font.Size = 14
tr.Font.Color.RGB = RGB(0, 0, 0)

# "Error" を赤太字に
found = tr.Find("Error")
if found is not None:
    found.Font.Color.RGB = RGB(255, 0, 0)
    found.Font.Bold = -1

# "Warning" をオレンジ太字に
found = tr.Find("Warning")
if found is not None:
    found.Font.Color.RGB = RGB(255, 165, 0)
    found.Font.Bold = -1
```

### 9.6 例5: InsertBefore / InsertAfter を使った書式付きテキスト構築

InsertBefore / InsertAfter は挿入した部分の TextRange を返すため、挿入と同時に書式設定が可能。

```python
tr = textbox.TextFrame.TextRange
tr.Text = ""  # テキストをクリア

# テキストを段階的に構築しながら書式を設定
part1 = tr.InsertAfter("売上報告: ")
part1.Font.Size = 20
part1.Font.Bold = -1
part1.Font.Color.RGB = RGB(0, 0, 0)

part2 = tr.InsertAfter("前年比 ")
part2.Font.Size = 16
part2.Font.Color.RGB = RGB(0, 0, 0)

part3 = tr.InsertAfter("+15%")
part3.Font.Size = 20
part3.Font.Bold = -1
part3.Font.Color.RGB = RGB(0, 128, 0)  # 緑（増加を示す）

part4 = tr.InsertAfter(" 達成")
part4.Font.Size = 16
part4.Font.Color.RGB = RGB(0, 0, 0)
```

### 9.7 例6: Runs を使った既存テキストの書式分析と変更

```python
tr = textbox.TextFrame.TextRange
# 既にテキストに複数の書式が適用されている場合

# 全ランを走査して書式情報を取得
runs = tr.Runs()
for i in range(1, runs.Count + 1):
    run = tr.Runs(i)
    print(f"Run {i}:")
    print(f"  Text: '{run.Text}'")
    print(f"  Font: {run.Font.Name}")
    print(f"  Size: {run.Font.Size}")
    print(f"  Bold: {run.Font.Bold}")
    print(f"  Color RGB: {run.Font.Color.RGB}")

# 特定のランの書式を変更
if runs.Count >= 2:
    run2 = tr.Runs(2)
    run2.Font.Size = run2.Font.Size + 4  # サイズを4ポイント大きく
```

### 9.8 例7: 段落ごとに異なる書式

```python
tr = textbox.TextFrame.TextRange
tr.Text = "タイトル\r本文テキスト。\rフッター"

# 1段落目: タイトルスタイル
p1 = tr.Paragraphs(1)
p1.Font.Size = 28
p1.Font.Bold = -1
p1.Font.Color.RGB = RGB(0, 51, 102)
p1.ParagraphFormat.Alignment = 2  # ppAlignCenter

# 2段落目: 本文スタイル
p2 = tr.Paragraphs(2)
p2.Font.Size = 14
p2.Font.Color.RGB = RGB(51, 51, 51)
p2.ParagraphFormat.Alignment = 1  # ppAlignLeft

# 3段落目: フッタースタイル
p3 = tr.Paragraphs(3)
p3.Font.Size = 10
p3.Font.Italic = -1
p3.Font.Color.RGB = RGB(128, 128, 128)
p3.ParagraphFormat.Alignment = 3  # ppAlignRight
```

### 9.9 制限事項と注意点

1. **Characters の Start は 1始まり**: Python の 0始まりインデックスとは異なるため、変換が必要。

2. **RGB 値の順序**: PowerPoint COM の RGB は `R + G*256 + B*65536` 形式（BGR エンコード）。Python の一般的な `0xRRGGBB` 表記とは逆なので注意。

3. **改行文字**: PowerPoint の段落区切りは `\r`（CR）。`\n`（LF）ではない。テキスト設定時に `\r` を使用すること。

4. **Font.Color.RGB 設定後の Type**: RGB を直接設定すると `Color.Type` が自動的に `msoColorTypeRGB (1)` に変わる。テーマカラーに戻したい場合は `ObjectThemeColor` を再設定する必要がある。

5. **Characters(Start, Length) の範囲外指定**: Start がテキスト長を超えると最後の文字から開始し、Length がテキスト残りを超えるとテキスト末尾までの範囲になる。エラーにはならないが意図しない結果になりうる。

6. **InsertBefore / InsertAfter の戻り値**: 挿入されたテキスト部分の TextRange を返す。この戻り値を使って書式設定するのが確実。テキスト全体のインデックスが変わるため、Characters() で後から参照する場合は位置の再計算が必要。

7. **パフォーマンス**: 大量の文字に個別に書式を適用する場合（例: 数千文字の虹色テキスト）、COM の呼び出しオーバーヘッドにより遅くなる可能性がある。可能な限り連続した範囲をまとめて処理すること。

8. **TextFrame vs TextFrame2**: TextFrame.TextRange は TextRange を返すが、TextFrame2.TextRange は TextRange2 を返す。両者のメソッド名は似ているが完全に同一ではない。混在して使用しないこと。

---

## 10. MCP サーバー実装への提言

### 10.1 優先実装機能（高優先度）

以下は、ユーザーがプレゼンテーション作成で最も頻繁に使用する機能であり、MCP サーバーで最初に実装すべきもの。

| 機能カテゴリ | 具体的な API | 理由 |
|---|---|---|
| **部分テキスト色変更** | `TextRange.Characters(start, length).Font.Color.RGB` | 最重要要件。ハイライト、強調、多色テキストに必須 |
| **フォント基本設定** | `Font.Name`, `.Size`, `.Bold`, `.Italic`, `.Underline` | テキスト書式の基本 |
| **テキスト配置** | `ParagraphFormat.Alignment` | 左/中央/右揃え |
| **テキスト設定** | `TextRange.Text`, `.InsertBefore`, `.InsertAfter` | テキスト操作の基本 |
| **塗りつぶし基本** | `Shape.Fill.Solid()`, `.ForeColor.RGB` | シェイプの色設定基本 |
| **枠線基本** | `Shape.Line.ForeColor.RGB`, `.Weight`, `.DashStyle` | シェイプの枠線基本 |
| **テキストフレーム余白** | `TextFrame.MarginLeft/Right/Top/Bottom` | テキスト位置調整 |
| **検索・置換** | `TextRange.Find`, `.Replace` | テキスト操作の自動化 |

### 10.2 2次優先機能

| 機能カテゴリ | 具体的な API |
|---|---|
| **行間設定** | `ParagraphFormat.SpaceBefore/After/Within`, `LineRule*` |
| **箇条書き** | `Bullet.Visible`, `.Type`, `.Character`, `.Font.Color` |
| **テーマカラー** | `Color.ObjectThemeColor`, `.TintAndShade`, `.Brightness` |
| **グラデーション** | `Fill.TwoColorGradient`, `GradientStops` |
| **影効果** | `Shadow.ForeColor`, `.OffsetX/Y`, `.Blur`, `.Transparency` |
| **矢印設定** | `Line.BeginArrowheadStyle/Length/Width`, `End*` |
| **多言語フォント** | `Font.NameFarEast`, `.NameAscii` |
| **インデント** | `IndentLevel`, `Ruler.Levels` |

### 10.3 3次優先機能（高度な装飾）

| 機能カテゴリ | 具体的な API |
|---|---|
| **光彩効果** | `Shape.Glow.Color/Radius/Transparency` |
| **反射効果** | `Shape.Reflection.Type/Blur/Offset/Size` |
| **ぼかし効果** | `Shape.SoftEdge.Type/Radius` |
| **3D 効果** | `Shape.ThreeD.BevelTopType/Depth/RotationX/Y` |
| **パターン塗りつぶし** | `Fill.Patterned()`, `.Pattern` |
| **テクスチャ/画像塗りつぶし** | `Fill.UserPicture()`, `.PresetTextured()` |
| **段組み** | `TextFrame2.Column.Number/Spacing` |
| **ワードアート** | `TextFrame2.WordArtFormat` |

### 10.4 RGB ヘルパー関数の実装

MCP サーバーには以下のユーティリティ関数を必ず含めること:

```python
def RGB(r: int, g: int, b: int) -> int:
    """VBA 互換の RGB 関数。PowerPoint COM の Color.RGB に設定する値を生成する。

    Args:
        r: 赤 (0-255)
        g: 緑 (0-255)
        b: 青 (0-255)

    Returns:
        PowerPoint COM 形式の RGB 整数値
    """
    return r + (g << 8) + (b << 16)


def rgb_to_components(rgb_value: int) -> tuple[int, int, int]:
    """PowerPoint COM の RGB 値を (R, G, B) タプルに変換する。

    Args:
        rgb_value: PowerPoint COM 形式の RGB 整数値

    Returns:
        (R, G, B) のタプル
    """
    r = rgb_value & 0xFF
    g = (rgb_value >> 8) & 0xFF
    b = (rgb_value >> 16) & 0xFF
    return (r, g, b)


def hex_to_rgb(hex_color: str) -> int:
    """'#RRGGBB' 形式の16進数カラーコードを PowerPoint COM 形式の RGB 値に変換する。

    Args:
        hex_color: '#RRGGBB' 形式のカラーコード

    Returns:
        PowerPoint COM 形式の RGB 整数値
    """
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGB(r, g, b)
```

### 10.5 テキスト書式設定 API の設計案

MCP サーバーのツール設計として、以下のような構造を推奨:

```
# ツール1: テキストの部分書式設定
format_text_range:
  - slide_index: int
  - shape_name_or_index: str | int
  - start: int (1-based)
  - length: int
  - font_name: str (optional)
  - font_size: float (optional)
  - bold: bool (optional)
  - italic: bool (optional)
  - underline: bool (optional)
  - color: str (optional, "#RRGGBB" format)
  - theme_color: int (optional)

# ツール2: 段落書式設定
format_paragraph:
  - slide_index: int
  - shape_name_or_index: str | int
  - paragraph_index: int (1-based)
  - alignment: str ("left", "center", "right", "justify")
  - space_before: float (optional)
  - space_after: float (optional)
  - line_spacing: float (optional)
  - indent_level: int (optional, 1-9)
  - bullet_visible: bool (optional)
  - bullet_type: str (optional, "none", "unnumbered", "numbered")
  - bullet_character: int (optional, unicode codepoint)

# ツール3: シェイプの塗りつぶし設定
set_shape_fill:
  - slide_index: int
  - shape_name_or_index: str | int
  - fill_type: str ("solid", "gradient", "pattern", "none")
  - color: str (optional, "#RRGGBB")
  - transparency: float (optional, 0.0-1.0)
  - gradient_colors: list (optional)
  - gradient_style: str (optional)

# ツール4: シェイプの枠線設定
set_shape_line:
  - slide_index: int
  - shape_name_or_index: str | int
  - visible: bool
  - color: str (optional, "#RRGGBB")
  - weight: float (optional)
  - dash_style: str (optional)
  - arrow_begin_style: str (optional)
  - arrow_end_style: str (optional)
```

---

## 参考リンク

- [TextRange object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange)
- [TextRange.Characters method (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.TextRange.Characters)
- [TextRange.Runs method (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.runs)
- [TextRange.Font property (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.font)
- [Font object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.font)
- [ColorFormat object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.colorformat)
- [TextFrame object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textframe)
- [TextFrame2 object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textframe2)
- [ParagraphFormat object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.paragraphformat)
- [BulletFormat object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.bulletformat)
- [Shape.Fill property (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.fill)
- [FillFormat.GradientStops property (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.fillformat.gradientstops)
- [LineFormat object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.lineformat)
- [ShadowFormat object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.ShadowFormat)
- [Shape object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape)
- [Working with Glow and Reflection Properties in Office 2010 | Microsoft Learn](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh148188(v=office.14))
- [When to Use TextFrame or TextFrame2 in VBA](https://copyprogramming.com/howto/when-to-use-textframe-or-textframe2-in-vba)
