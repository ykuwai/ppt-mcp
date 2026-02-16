# PowerPoint COM オートメーション: Slides & Shapes 管理 詳細リファレンス

> 本ドキュメントは MCP サーバー実装の判断材料として、PowerPoint COM オートメーションにおけるスライドおよびシェイプ操作の全機能を調査・整理したものである。

---

## 目次

1. [Slides コレクション](#1-slides-コレクション)
2. [Shapes コレクション - 基本操作](#2-shapes-コレクション---基本操作)
3. [Shape 個別操作](#3-shape-個別操作)
4. [Shape の種類 (MsoShapeType)](#4-shape-の種類-msoshapetype)
5. [主要定数・列挙型一覧](#5-主要定数列挙型一覧)
6. [Python win32com での定数利用方法](#6-python-win32com-での定数利用方法)
7. [MCP サーバー実装への推奨事項](#7-mcp-サーバー実装への推奨事項)

---

## 1. Slides コレクション

### 1.1 スライドの追加

#### Slides.Add (レガシー)

古い API。`Layout` パラメータに `PpSlideLayout` 列挙値（整数）を渡す。

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
prs = app.Presentations.Add()

# ppLayoutBlank = 12
slide = prs.Slides.Add(Index=1, Layout=12)

# ppLayoutTitle = 1
title_slide = prs.Slides.Add(Index=2, Layout=1)
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Index | Long | Yes | 挿入位置（1始まり） |
| Layout | PpSlideLayout | Yes | スライドレイアウト定数 |

**戻り値:** `Slide` オブジェクト

#### Slides.AddSlide (新しい API、PowerPoint 2007+)

`CustomLayout` オブジェクトを直接指定できる新しいメソッド。

```python
# SlideMaster の最初の CustomLayout を取得して使用
prs = app.Presentations.Add()
custom_layout = prs.SlideMaster.CustomLayouts(1)
slide = prs.Slides.AddSlide(Index=1, pCustomLayout=custom_layout)
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Index | Long | Yes | 挿入位置 |
| pCustomLayout | CustomLayout | Yes | カスタムレイアウトオブジェクト |

**戻り値:** `Slide` オブジェクト

**注意点:**
- `Slides.Add` は古い API だが、現在も動作する。整数のレイアウト番号を使う場合はこちらが簡単。
- `Slides.AddSlide` はカスタムレイアウトを柔軟に指定できるため、テーマに依存するレイアウトを使う場合に推奨。

---

### 1.2 スライドの削除・複製・移動

#### Slide.Delete

```python
# スライド1を削除
prs.Slides(1).Delete()
```

**注意:** 削除後、残りのスライドのインデックスが自動的にシフトする。ループ中に削除する場合は後ろから削除すること。

#### Slide.Duplicate

```python
# スライド1を複製（直後に挿入される）
dup_range = prs.Slides(1).Duplicate()
# 戻り値は SlideRange オブジェクト
```

**注意:** Duplicate は複製先を指定できない。複製後に `MoveTo` で移動する必要がある。

#### Slide.Copy

```python
# スライドをクリップボードにコピー
prs.Slides(1).Copy()
```

#### Slide.MoveTo

```python
# スライド1を最後に移動
prs.Slides(1).MoveTo(toPos=prs.Slides.Count)

# スライド3を先頭に移動
prs.Slides(3).MoveTo(toPos=1)
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| toPos | Long | Yes | 移動先の位置 |

---

### 1.3 スライドのレイアウト (PpSlideLayout)

`Slide.Layout` プロパティで取得・設定可能。

```python
# 現在のレイアウトを取得
current_layout = prs.Slides(1).Layout

# レイアウトを変更
prs.Slides(1).Layout = 12  # ppLayoutBlank
```

#### PpSlideLayout 主要定数一覧

| 定数名 | 値 | 説明 |
|---|---|---|
| ppLayoutTitle | 1 | タイトルスライド |
| ppLayoutText | 2 | テキスト |
| ppLayoutTwoColumnText | 3 | 2段組テキスト |
| ppLayoutTable | 4 | テーブル |
| ppLayoutTextAndChart | 5 | テキストとグラフ |
| ppLayoutChartAndText | 6 | グラフとテキスト |
| ppLayoutOrgchart | 7 | 組織図 |
| ppLayoutChart | 8 | グラフ |
| ppLayoutTextAndClipArt | 9 | テキストとクリップアート |
| ppLayoutClipArtAndText | 10 | クリップアートとテキスト |
| ppLayoutTitleOnly | 11 | タイトルのみ |
| ppLayoutBlank | 12 | 白紙 |
| ppLayoutTextAndObject | 13 | テキストとオブジェクト |
| ppLayoutObjectAndText | 14 | オブジェクトとテキスト |
| ppLayoutLargeObject | 15 | 大きなオブジェクト |
| ppLayoutObject | 16 | オブジェクト |
| ppLayoutTextAndMediaClip | 17 | テキストとメディアクリップ |
| ppLayoutMediaClipAndText | 18 | メディアクリップとテキスト |
| ppLayoutObjectOverText | 19 | テキスト上のオブジェクト |
| ppLayoutTextOverObject | 20 | オブジェクト上のテキスト |
| ppLayoutTextAndTwoObjects | 21 | テキストと2つのオブジェクト |
| ppLayoutTwoObjectsAndText | 22 | 2つのオブジェクトとテキスト |
| ppLayoutTwoObjectsOverText | 23 | テキスト上の2つのオブジェクト |
| ppLayoutFourObjects | 24 | 4つのオブジェクト |
| ppLayoutVerticalText | 25 | 縦書きテキスト |
| ppLayoutClipArtAndVerticalText | 26 | クリップアートと縦書きテキスト |
| ppLayoutVerticalTitleAndText | 27 | 縦書きタイトルとテキスト |
| ppLayoutVerticalTitleAndTextOverChart | 28 | 縦書きタイトルとグラフ上のテキスト |
| ppLayoutTwoObjects | 29 | 2つのオブジェクト |
| ppLayoutObjectAndTwoObjects | 30 | オブジェクトと2つのオブジェクト |
| ppLayoutTwoObjectsAndObject | 31 | 2つのオブジェクトとオブジェクト |
| ppLayoutCustom | 32 | カスタム |
| ppLayoutSectionHeader | 33 | セクションヘッダー |
| ppLayoutComparison | 34 | 比較 |
| ppLayoutContentWithCaption | 35 | キャプション付きコンテンツ |
| ppLayoutPictureWithCaption | 36 | キャプション付き画像 |
| ppLayoutMixed | -2 | 混合（複数スライド選択時の戻り値） |

---

### 1.4 カスタムレイアウト (CustomLayout)

#### SlideMaster と CustomLayouts の取得

```python
# プレゼンテーションの SlideMaster を取得
master = prs.SlideMaster

# CustomLayouts コレクション（利用可能なレイアウト一覧）
layouts = master.CustomLayouts

# レイアウト数を取得
print(f"レイアウト数: {layouts.Count}")

# 各レイアウトの名前を表示
for i in range(1, layouts.Count + 1):
    layout = layouts(i)
    print(f"  {i}: {layout.Name}")
```

#### 特定のレイアウトを使ってスライドを追加

```python
# 名前でレイアウトを検索
def find_layout(master, layout_name):
    for i in range(1, master.CustomLayouts.Count + 1):
        if master.CustomLayouts(i).Name == layout_name:
            return master.CustomLayouts(i)
    return None

layout = find_layout(prs.SlideMaster, "Title Slide")
if layout:
    slide = prs.Slides.AddSlide(Index=1, pCustomLayout=layout)
```

#### 複数の SlideMaster がある場合

```python
# Designs コレクションから各 SlideMaster にアクセス
for i in range(1, prs.Designs.Count + 1):
    design = prs.Designs(i)
    master = design.SlideMaster
    print(f"Design {i}: {design.Name}")
    for j in range(1, master.CustomLayouts.Count + 1):
        print(f"  Layout {j}: {master.CustomLayouts(j).Name}")
```

#### SlideLayout の取得

```python
# スライドに適用されている CustomLayout を取得
slide = prs.Slides(1)
layout = slide.CustomLayout
print(f"現在のレイアウト: {layout.Name}")

# レイアウトを変更
new_layout = prs.SlideMaster.CustomLayouts(2)
slide.CustomLayout = new_layout
```

**注意点:**
- `Slide.Layout` プロパティ（PpSlideLayout 整数値）と `Slide.CustomLayout` プロパティ（CustomLayout オブジェクト）は異なる。
- `CustomLayout` を使うと、テーマに定義されたすべてのレイアウトにアクセスできる。
- `Slide.Layout` は標準レイアウトのみサポート。カスタムテーマのレイアウトには `CustomLayout` を使う。

---

### 1.5 スライド番号・名前でのアクセス

```python
# インデックスでアクセス（1始まり）
slide = prs.Slides(1)

# スライドIDでアクセス（一意のID）
slide_id = prs.Slides(1).SlideID
slide = prs.Slides.FindBySlideID(slide_id)

# スライド名でアクセス（Slidesコレクションの Item）
# ※ スライド名は通常 "Slide1", "Slide2" 等
slide = prs.Slides("Slide1")

# スライド番号の取得
slide_number = prs.Slides(1).SlideNumber
slide_index = prs.Slides(1).SlideIndex

# スライド名の取得・設定
slide_name = prs.Slides(1).Name
prs.Slides(1).Name = "MyCustomSlide"
```

**注意点:**
- `SlideIndex` はコレクション内の位置（1始まり）。
- `SlideNumber` は表示されるスライド番号（ページ番号設定に影響される）。
- `SlideID` はプレゼンテーション内で一意の不変ID。スライドの追加・削除で変わらない。
- `FindBySlideID` は ID からスライドを検索するメソッド。

---

### 1.6 スライドのコピー＆ペースト（プレゼンテーション間）

#### 方法1: Copy + Paste

```python
# 元のプレゼンテーションのスライドをコピー
source_prs = app.Presentations.Open(r"C:\source.pptx")
source_prs.Slides(1).Copy()

# 別のプレゼンテーションにペースト
dest_prs = app.Presentations.Open(r"C:\dest.pptx")
dest_prs.Slides.Paste(Index=dest_prs.Slides.Count + 1)
```

#### 方法2: InsertFromFile

```python
# ファイルからスライドを挿入（より効率的）
dest_prs.Slides.InsertFromFile(
    FileName=r"C:\source.pptx",
    Index=dest_prs.Slides.Count,  # 挿入位置
    SlideStart=1,                  # 開始スライド番号
    SlideEnd=3                     # 終了スライド番号
)
```

**Slides.InsertFromFile パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| FileName | String | Yes | ソースファイルのパス |
| Index | Long | Yes | 挿入位置（この位置の後に挿入） |
| SlideStart | Long | No | コピー開始スライド番号（デフォルト: 1） |
| SlideEnd | Long | No | コピー終了スライド番号（デフォルト: 最後） |

**Slides.Paste パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Index | Long | No | ペースト位置（省略時は最後に追加） |

**戻り値:** `SlideRange` オブジェクト

**注意点:**
- `Copy + Paste` ではソースの書式が目的プレゼンテーションのテーマに合わせて変換される場合がある。
- `InsertFromFile` の方がパフォーマンスが良い（クリップボード不使用）。
- どちらの方法でも、元のプレゼンテーションのマスタースライドは自動的にコピーされない場合がある。

---

### 1.7 NotesPage (ノートの追加・取得)

```python
# ノートテキストの取得
slide = prs.Slides(1)
notes_text = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
print(f"ノート: {notes_text}")

# ノートテキストの設定
slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "これはスピーカーノートです"

# ノートにテキストを追加（既存テキストの後に）
notes_tf = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange
notes_tf.InsertAfter("\n追加テキスト")
```

**オブジェクト階層:**
```
Slide
  └── NotesPage (SlideRange)
        └── Shapes
              ├── Placeholders(1)  ... スライドのサムネイル画像
              └── Placeholders(2)  ... ノートのテキスト
                    └── TextFrame
                          └── TextRange
                                └── Text (String)
```

**注意点:**
- `NotesPage` は `SlideRange` オブジェクトを返す。
- ノートのテキストプレースホルダーは通常 `Placeholders(2)` だが、レイアウトによっては異なる場合がある。
- `HasTextFrame` を使って事前にテキストフレームの有無を確認するとより安全。
- ノートページの背景やレイアウトも変更可能。

---

### 1.8 スライドの背景設定 (Background.Fill)

```python
slide = prs.Slides(1)

# マスタースライドの背景を追従しない設定
slide.FollowMasterBackground = False

# 単色背景
slide.Background.Fill.Solid()
slide.Background.Fill.ForeColor.RGB = 0xFF0000  # 赤 (BGR形式: 0xBBGGRR)

# RGB値を正しく設定するヘルパー
def rgb_to_bgr(r, g, b):
    """RGBをPowerPointのBGR形式に変換"""
    return r + (g << 8) + (b << 16)

slide.Background.Fill.ForeColor.RGB = rgb_to_bgr(100, 149, 237)  # コーンフラワーブルー

# グラデーション背景
# msoGradientHorizontal = 1, msoGradientLateSunset = 8
slide.Background.Fill.PresetGradient(
    Style=1,          # msoGradientHorizontal
    Variant=1,
    PresetGradientType=8  # msoGradientLateSunset
)

# パターン背景
# msoPatternHorizontalBrick = 35
slide.Background.Fill.Patterned(Pattern=35)

# 画像背景
slide.Background.Fill.UserPicture(PictureFile=r"C:\bg_image.jpg")

# テクスチャ背景
# msoTextureGranite = 12
slide.Background.Fill.PresetTextured(PresetTexture=12)
```

**主要な Fill メソッド:**
| メソッド | 説明 |
|---|---|
| `.Solid()` | 単色塗りつぶし |
| `.PresetGradient(Style, Variant, PresetGradientType)` | プリセットグラデーション |
| `.TwoColorGradient(Style, Variant)` | 2色グラデーション |
| `.OneColorGradient(Style, Variant, Degree)` | 1色グラデーション |
| `.Patterned(Pattern)` | パターン |
| `.PresetTextured(PresetTexture)` | プリセットテクスチャ |
| `.UserPicture(PictureFile)` | ユーザー画像 |
| `.UserTextured(TextureFile)` | ユーザーテクスチャ |

**注意点:**
- `FollowMasterBackground = False` を設定しないと、個別の背景設定が反映されない。
- 色は BGR 形式（Blue-Green-Red）で指定する必要がある。Python から VBA の `RGB()` 関数は使えないため、手動で変換する。

---

### 1.9 スライドのトランジション (SlideShowTransition)

```python
slide = prs.Slides(1)
transition = slide.SlideShowTransition

# トランジション効果の設定
# ppEffectStripsDownLeft = 2321
transition.EntryEffect = 2321
transition.Speed = 2  # ppTransitionSpeedFast = 3, Medium = 2, Slow = 1

# 自動送り設定
transition.AdvanceOnClick = True   # クリックで進む
transition.AdvanceOnTime = True    # 時間で自動進行
transition.AdvanceTime = 5         # 5秒後に自動進行

# トランジションの継続時間（秒）
transition.Duration = 1.5

# サウンド効果
transition.SoundEffect.ImportFromFile(r"C:\sound\transition.wav")
transition.LoopSoundUntilNext = False

# スライドを非表示にする
transition.Hidden = True  # スライドショーで非表示
```

**SlideShowTransition プロパティ一覧:**
| プロパティ | 型 | 説明 |
|---|---|---|
| AdvanceOnClick | Boolean | クリックで次のスライドに進むか |
| AdvanceOnTime | Boolean | 時間経過で自動進行するか |
| AdvanceTime | Long | 自動進行までの秒数 |
| Duration | Single | トランジションの継続時間（秒） |
| EntryEffect | PpEntryEffect | トランジション効果の種類 |
| Hidden | Boolean | スライドショーで非表示にするか |
| LoopSoundUntilNext | Boolean | 次のスライドまでサウンドをループ |
| SoundEffect | SoundEffect | サウンド効果オブジェクト |
| Speed | PpTransitionSpeed | トランジションの速度 |

**PpTransitionSpeed 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| ppTransitionSpeedSlow | 1 | 遅い |
| ppTransitionSpeedMedium | 2 | 普通 |
| ppTransitionSpeedFast | 3 | 速い |
| ppTransitionSpeedMixed | -2 | 混合 |

**注意点:**
- `PpEntryEffect` 列挙型には数百の定数がある（ppEffectNone = 0 から各種エフェクトまで）。
- NotesPage に対して `SlideShowTransition` を設定しようとするとエラーになる。

---

## 2. Shapes コレクション - 基本操作

### 2.1 Shapes.AddShape (オートシェイプ)

```python
slide = prs.Slides(1)

# 四角形を追加
# msoShapeRectangle = 1
rect = slide.Shapes.AddShape(
    Type=1,       # msoShapeRectangle
    Left=100,     # 左端からの距離（ポイント）
    Top=100,      # 上端からの距離（ポイント）
    Width=200,    # 幅（ポイント）
    Height=100    # 高さ（ポイント）
)

# 円を追加
# msoShapeOval = 9
oval = slide.Shapes.AddShape(Type=9, Left=350, Top=100, Width=100, Height=100)

# 角丸四角形
# msoShapeRoundedRectangle = 5
rounded = slide.Shapes.AddShape(Type=5, Left=100, Top=250, Width=200, Height=100)

# 右矢印
# msoShapeRightArrow = 33
arrow = slide.Shapes.AddShape(Type=33, Left=350, Top=250, Width=150, Height=80)

# 5角星
# msoShape5pointStar = 92
star = slide.Shapes.AddShape(Type=92, Left=100, Top=400, Width=120, Height=120)

# 名前を設定
rect.Name = "MyRectangle"
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Type | MsoAutoShapeType | Yes | シェイプの種類 |
| Left | Single | Yes | 左端位置（ポイント） |
| Top | Single | Yes | 上端位置（ポイント） |
| Width | Single | Yes | 幅（ポイント） |
| Height | Single | Yes | 高さ（ポイント） |

**戻り値:** `Shape` オブジェクト

**主要な MsoAutoShapeType 定数（よく使われるもの）:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoShapeRectangle | 1 | 四角形 |
| msoShapeParallelogram | 2 | 平行四辺形 |
| msoShapeTrapezoid | 3 | 台形 |
| msoShapeDiamond | 4 | ひし形 |
| msoShapeRoundedRectangle | 5 | 角丸四角形 |
| msoShapeOctagon | 6 | 八角形 |
| msoShapeIsoscelesTriangle | 7 | 二等辺三角形 |
| msoShapeRightTriangle | 8 | 直角三角形 |
| msoShapeOval | 9 | 楕円 |
| msoShapeHexagon | 10 | 六角形 |
| msoShapeCross | 11 | 十字形 |
| msoShapeRegularPentagon | 12 | 正五角形 |
| msoShapeCan | 13 | 缶 |
| msoShapeCube | 14 | 立方体 |
| msoShapeDonut | 18 | ドーナツ |
| msoShapeNoSymbol | 19 | 禁止マーク |
| msoShapeHeart | 21 | ハート |
| msoShapeLightningBolt | 22 | 稲妻 |
| msoShapeSun | 23 | 太陽 |
| msoShapeMoon | 24 | 月 |
| msoShapeArc | 25 | 弧 |
| msoShapeSmileyFace | 17 | スマイリー |
| msoShapeRightArrow | 33 | 右矢印ブロック |
| msoShapeLeftArrow | 34 | 左矢印ブロック |
| msoShapeUpArrow | 35 | 上矢印ブロック |
| msoShapeDownArrow | 36 | 下矢印ブロック |
| msoShapeLeftRightArrow | 37 | 左右矢印ブロック |
| msoShapeUpDownArrow | 38 | 上下矢印ブロック |
| msoShapeQuadArrow | 39 | 四方向矢印ブロック |
| msoShapePentagon | 51 | 五角形 |
| msoShapeChevron | 52 | シェブロン |
| msoShape4pointStar | 91 | 4角星 |
| msoShape5pointStar | 92 | 5角星 |
| msoShape8pointStar | 93 | 8角星 |
| msoShape16pointStar | 94 | 16角星 |
| msoShape24pointStar | 95 | 24角星 |
| msoShape32pointStar | 96 | 32角星 |
| msoShapeExplosion1 | 89 | 爆発1 |
| msoShapeExplosion2 | 90 | 爆発2 |
| msoShapeFlowchartProcess | 61 | フローチャート: 処理 |
| msoShapeFlowchartDecision | 63 | フローチャート: 判断 |
| msoShapeFlowchartData | 64 | フローチャート: データ |
| msoShapeFlowchartTerminator | 69 | フローチャート: 端子 |
| msoShapeFlowchartDocument | 67 | フローチャート: 書類 |
| msoShapeFlowchartConnector | 73 | フローチャート: 結合子 |
| msoShapeCloud | 179 | 雲 |
| msoShapeFunnel | 174 | 漏斗 |

> MsoAutoShapeType には全部で 180 以上の定数が定義されている。フローチャート、矢印、吹き出し、アクションボタン等を含む。

---

### 2.2 Shapes.AddTextbox (テキストボックス)

```python
# msoTextOrientationHorizontal = 1
textbox = slide.Shapes.AddTextbox(
    Orientation=1,   # msoTextOrientationHorizontal
    Left=100,
    Top=100,
    Width=300,
    Height=50
)

# テキストを設定
textbox.TextFrame.TextRange.Text = "Hello, PowerPoint!"

# フォント設定
textbox.TextFrame.TextRange.Font.Size = 24
textbox.TextFrame.TextRange.Font.Bold = True
textbox.TextFrame.TextRange.Font.Color.RGB = rgb_to_bgr(0, 0, 255)

# テキストの配置
# ppAlignCenter = 2
textbox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

# 自動サイズ調整
# ppAutoSizeNone = 0, ppAutoSizeShapeToFitText = 1, ppAutoSizeMixed = -2
textbox.TextFrame.AutoSize = 1
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Orientation | MsoTextOrientation | Yes | テキスト方向 |
| Left | Single | Yes | 左端位置（ポイント） |
| Top | Single | Yes | 上端位置（ポイント） |
| Width | Single | Yes | 幅（ポイント） |
| Height | Single | Yes | 高さ（ポイント） |

**MsoTextOrientation 主要定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoTextOrientationHorizontal | 1 | 横書き |
| msoTextOrientationVertical | 5 | 縦書き |
| msoTextOrientationUpward | 2 | 上向き（90度回転） |
| msoTextOrientationDownward | 3 | 下向き（270度回転） |
| msoTextOrientationVerticalFarEast | 6 | 縦書き（東アジア） |

---

### 2.3 Shapes.AddPicture (画像の挿入)

```python
# 画像を挿入
pic = slide.Shapes.AddPicture(
    FileName=r"C:\images\photo.jpg",
    LinkToFile=False,    # msoFalse = 0
    SaveWithDocument=True,  # msoTrue = -1
    Left=100,
    Top=100,
    Width=300,
    Height=200
)

# 元の縦横比を維持
pic.LockAspectRatio = True  # msoTrue = -1

# 画像フォーマットにアクセス
pic.PictureFormat.Brightness = 0.5
pic.PictureFormat.Contrast = 0.7
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| FileName | String | Yes | 画像ファイルパス |
| LinkToFile | MsoTriState | Yes | ファイルにリンクするか |
| SaveWithDocument | MsoTriState | Yes | ドキュメントに保存するか |
| Left | Single | Yes | 左端位置（ポイント） |
| Top | Single | Yes | 上端位置（ポイント） |
| Width | Single | Yes | 幅（ポイント）。-1 で元のサイズ |
| Height | Single | Yes | 高さ（ポイント）。-1 で元のサイズ |

#### Shapes.AddPicture2 (PowerPoint 2013+)

圧縮品質を指定できる拡張版。

```python
# msoPictureCompressTrue = 1, msoPictureCompressDocDefault = -1
pic = slide.Shapes.AddPicture2(
    FileName=r"C:\images\photo.jpg",
    LinkToFile=False,
    SaveWithDocument=True,
    Left=100,
    Top=100,
    Width=300,
    Height=200,
    Compress=-1  # msoPictureCompressDocDefault
)
```

**注意点:**
- `LinkToFile=True` かつ `SaveWithDocument=False` の場合、ファイルの参照のみが保存される（ファイル移動で壊れる）。
- `LinkToFile=False` かつ `SaveWithDocument=False` はエラーになる。
- Width/Height に `-1` を指定すると元の画像サイズが使われる。

---

### 2.4 Shapes.AddChart2 (グラフの挿入)

```python
# グラフを追加
# XlChartType: xlColumnClustered = 51, xlLine = 4, xlPie = 5, xlBar = 57
chart_shape = slide.Shapes.AddChart2(
    Style=-1,         # デフォルトスタイル
    Type=51,          # xlColumnClustered（集合縦棒）
    Left=100,
    Top=100,
    Width=400,
    Height=300,
    NewLayout=True
)

# グラフオブジェクトにアクセス
chart = chart_shape.Chart

# グラフデータの編集（Excel ワークブック経由）
wb = chart.ChartData.Workbook
ws = wb.Worksheets(1)
ws.Range("A1").Value = "カテゴリ"
ws.Range("B1").Value = "系列1"
ws.Range("A2").Value = "項目A"
ws.Range("B2").Value = 100
ws.Range("A3").Value = "項目B"
ws.Range("B3").Value = 200
chart.ChartData.Activate()  # データを反映

# グラフタイトル
chart.HasTitle = True
chart.ChartTitle.Text = "売上グラフ"
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Style | Long | No | グラフスタイル（-1でデフォルト） |
| Type | XlChartType | No | グラフの種類 |
| Left | Single | No | 左端位置 |
| Top | Single | No | 上端位置 |
| Width | Single | No | 幅 |
| Height | Single | No | 高さ |
| NewLayout | Boolean | No | 新しい動的書式ルールを使うか |

**主要な XlChartType 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| xlColumnClustered | 51 | 集合縦棒 |
| xlColumnStacked | 52 | 積み上げ縦棒 |
| xlBarClustered | 57 | 集合横棒 |
| xlLine | 4 | 折れ線 |
| xlLineMarkers | 65 | マーカー付き折れ線 |
| xlPie | 5 | 円 |
| xlArea | 1 | 面 |
| xlXYScatter | -4169 | 散布図 |
| xlDoughnut | -4120 | ドーナツ |
| xlRadar | -4151 | レーダー |

**注意点:**
- `AddChart` (旧メソッド) は非推奨。`AddChart2` を使用すること。
- グラフデータの編集には内部的に Excel が起動される。パフォーマンスに影響する可能性がある。
- `ChartData.Workbook` でアクセスする Excel ワークブックは使用後に閉じる必要はない（自動管理）。

---

### 2.5 Shapes.AddTable (テーブルの挿入)

```python
# 3行4列のテーブルを追加
table_shape = slide.Shapes.AddTable(
    NumRows=3,
    NumColumns=4,
    Left=50,
    Top=100,
    Width=600,   # 幅（ポイント）
    Height=200   # 高さ（ポイント）
)

# テーブルオブジェクトにアクセス
table = table_shape.Table

# セルにテキストを設定
table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "ヘッダー1"
table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "ヘッダー2"
table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "データ1"
table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "データ2"

# 行・列数の取得
num_rows = table.Rows.Count
num_cols = table.Columns.Count

# セルの結合
table.Cell(1, 1).Merge(table.Cell(1, 2))

# 列幅の設定
table.Columns(1).Width = 150

# 行の高さの設定
table.Rows(1).Height = 40

# セルの背景色
table.Cell(1, 1).Shape.Fill.ForeColor.RGB = rgb_to_bgr(0, 0, 128)
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| NumRows | Long | Yes | 行数 |
| NumColumns | Long | Yes | 列数 |
| Left | Single | No | 左端位置（デフォルト: 0） |
| Top | Single | No | 上端位置（デフォルト: 0） |
| Width | Single | No | 幅 |
| Height | Single | No | 高さ |

**注意点:**
- テーブルは Shape の中に Table オブジェクトを持つ構造。`table_shape.HasTable` で確認可能。
- セルのテキストは `Cell(Row, Col).Shape.TextFrame.TextRange.Text` でアクセス。
- Cell のインデックスは 1 始まり。
- 結合セルへのアクセスには注意が必要。

---

### 2.6 Shapes.AddMediaObject2 (動画・音声)

```python
# 動画を挿入
video = slide.Shapes.AddMediaObject2(
    FileName=r"C:\videos\sample.mp4",
    LinkToFile=False,     # msoFalse
    SaveWithDocument=True, # msoTrue
    Left=100,
    Top=100,
    Width=400,    # -1 でデフォルトサイズ
    Height=300    # -1 でデフォルトサイズ
)

# 音声を挿入
audio = slide.Shapes.AddMediaObject2(
    FileName=r"C:\audio\bgm.mp3",
    LinkToFile=False,
    SaveWithDocument=True,
    Left=50,
    Top=50,
    Width=-1,
    Height=-1
)
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| FileName | String | Yes | メディアファイルパス |
| LinkToFile | MsoTriState | No | ファイルにリンクするか |
| SaveWithDocument | MsoTriState | No | ドキュメントに保存するか |
| Left | Single | No | 左端位置 |
| Top | Single | No | 上端位置 |
| Width | Single | No | 幅（-1でデフォルト） |
| Height | Single | No | 高さ（-1でデフォルト） |

**注意点:**
- `AddMediaObject` (旧メソッド) は PowerPoint 2013 で非推奨。`AddMediaObject2` を使用。
- `LinkToFile=False` かつ `SaveWithDocument=False` はエラーになる。
- サポートされるメディア形式は PowerPoint のバージョンに依存する（mp4, wmv, avi, mp3, wav 等）。

---

### 2.7 Shapes.AddOLEObject (OLE オブジェクト)

```python
# Excel ワークシートを埋め込み
ole = slide.Shapes.AddOLEObject(
    Left=100,
    Top=100,
    Width=400,
    Height=300,
    ClassName="Excel.Sheet"
)

# ファイルから OLE オブジェクトを作成
ole_doc = slide.Shapes.AddOLEObject(
    Left=100,
    Top=100,
    Width=400,
    Height=300,
    FileName=r"C:\documents\report.xlsx"
)

# リンクとして挿入
ole_linked = slide.Shapes.AddOLEObject(
    Left=100,
    Top=100,
    Width=400,
    Height=300,
    FileName=r"C:\documents\report.xlsx",
    Link=True  # msoTrue
)

# アイコンとして表示
ole_icon = slide.Shapes.AddOLEObject(
    Left=100,
    Top=100,
    Width=100,
    Height=100,
    ClassName="Excel.Sheet",
    DisplayAsIcon=True,
    IconLabel="Excel Data"
)
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Left | Single | No | 左端位置（デフォルト: 0） |
| Top | Single | No | 上端位置（デフォルト: 0） |
| Width | Single | No | 幅（デフォルト: 自動） |
| Height | Single | No | 高さ（デフォルト: 自動） |
| ClassName | String | No | OLE クラス名 / ProgID |
| FileName | String | No | ソースファイルパス |
| DisplayAsIcon | MsoTriState | No | アイコン表示するか |
| IconFileName | String | No | アイコンファイルパス |
| IconIndex | Long | No | アイコンインデックス |
| IconLabel | String | No | アイコンの下に表示するラベル |
| Link | MsoTriState | No | リンクとして挿入するか |

**注意点:**
- `ClassName` と `FileName` のどちらか一方を指定する（両方は不可）。
- 一般的な ClassName: `"Excel.Sheet"`, `"Word.Document"`, `"Forms.CommandButton.1"`
- OLE オブジェクトはファイルサイズを大幅に増加させる可能性がある。

---

### 2.8 Shapes.AddConnector (コネクタ)

```python
# 2つのシェイプを作成
shape1 = slide.Shapes.AddShape(Type=1, Left=100, Top=100, Width=100, Height=60)
shape2 = slide.Shapes.AddShape(Type=1, Left=400, Top=300, Width=100, Height=60)

# コネクタを追加
# msoConnectorStraight = 1, msoConnectorElbow = 2, msoConnectorCurve = 3
connector = slide.Shapes.AddConnector(
    Type=2,       # msoConnectorElbow（L字コネクタ）
    BeginX=0,     # 始点X（接続後に自動調整）
    BeginY=0,     # 始点Y
    EndX=100,     # 終点X
    EndY=100      # 終点Y
)

# シェイプに接続
connector.ConnectorFormat.BeginConnect(
    ConnectedShape=shape1,
    ConnectionSite=1  # 接続ポイント番号
)
connector.ConnectorFormat.EndConnect(
    ConnectedShape=shape2,
    ConnectionSite=3
)

# 接続を最適化
connector.RerouteConnections()
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Type | MsoConnectorType | Yes | コネクタの種類 |
| BeginX | Single | Yes | 始点X座標（ポイント） |
| BeginY | Single | Yes | 始点Y座標（ポイント） |
| EndX | Single | Yes | 終点X座標（ポイント） |
| EndY | Single | Yes | 終点Y座標（ポイント） |

**MsoConnectorType 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoConnectorStraight | 1 | 直線 |
| msoConnectorElbow | 2 | L字（カギ線） |
| msoConnectorCurve | 3 | 曲線 |
| msoConnectorTypeMixed | -2 | 混合 |

**注意点:**
- コネクタの座標は、`BeginConnect`/`EndConnect` でシェイプに接続すると自動調整される。
- `ConnectionSite` はシェイプの接続ポイント番号。`Shape.ConnectionSiteCount` で確認可能。
- `RerouteConnections()` で最短経路に再配線される。

---

### 2.9 Shapes.AddSmartArt

```python
# SmartArt レイアウトを取得
# Application.SmartArtLayouts で全レイアウトにアクセス
layout = app.SmartArtLayouts(1)  # インデックスで指定

# URN で特定のレイアウトを指定
layout = app.SmartArtLayouts("urn:microsoft.com/office/officeart/2005/8/layout/orgChart1")

# SmartArt を追加
smart_art = slide.Shapes.AddSmartArt(
    Layout=layout,
    Left=50,
    Top=50,
    Width=500,
    Height=400
)

# SmartArt のノードにテキストを設定
smart_art.SmartArt.AllNodes(1).TextFrame2.TextRange.Text = "管理者"
smart_art.SmartArt.AllNodes(2).TextFrame2.TextRange.Text = "部下1"
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| Layout | SmartArtLayout | Yes | SmartArt レイアウト |
| Left | Single | No | 左端位置 |
| Top | Single | No | 上端位置 |
| Width | Single | No | 幅 |
| Height | Single | No | 高さ |

**注意点:**
- `SmartArtLayouts` コレクションは `Application` オブジェクトからアクセスする（`Presentation` ではない）。
- SmartArt の操作は複雑で、ノードの追加・削除・移動が可能。
- レイアウトの URN は Office のバージョンによって異なる可能性がある。

---

### 2.10 Shapes.AddLine (線)

```python
# 直線を追加
line = slide.Shapes.AddLine(
    BeginX=50,
    BeginY=50,
    EndX=300,
    EndY=200
)

# 線のスタイル設定
line.Line.Weight = 3.0           # 線の太さ（ポイント）
line.Line.ForeColor.RGB = rgb_to_bgr(255, 0, 0)  # 赤

# msoLineDash = 4, msoLineDashDot = 5
line.Line.DashStyle = 4          # 破線

# 矢印の設定
# msoArrowheadTriangle = 2
line.Line.BeginArrowheadStyle = 0  # msoArrowheadNone
line.Line.EndArrowheadStyle = 2    # msoArrowheadTriangle
```

**パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| BeginX | Single | Yes | 始点X座標 |
| BeginY | Single | Yes | 始点Y座標 |
| EndX | Single | Yes | 終点X座標 |
| EndY | Single | Yes | 終点Y座標 |

---

### 2.11 Shapes.AddCurve / AddPolyline (曲線・ポリライン)

#### AddCurve (ベジエ曲線)

```python
import array

# SafeArrayOfPoints として座標を渡す
# ベジエ曲線の制御点（始点、制御点1、制御点2、終点、...の繰り返し）
# ※ win32com では 2次元配列として渡す
points = ((0, 0), (50, 100), (100, 100), (150, 0))

# VBA では SafeArrayOfPoints を使うが、Python では配列として渡す
curve = slide.Shapes.AddCurve(SafeArrayOfPoints=points)
```

#### AddPolyline (ポリライン)

```python
# ポリラインの頂点座標
points = ((100, 100), (200, 50), (300, 100), (400, 50))
polyline = slide.Shapes.AddPolyline(SafeArrayOfPoints=points)
```

**注意点:**
- Python (win32com) からの SafeArrayOfPoints の渡し方は Python のタプルまたはリストのリストを使用する。
- `AddCurve` はベジエ曲線で、制御点の数は 3n+1 である必要がある（n はセグメント数）。
- `AddPolyline` は直線で接続される頂点群を指定する。

---

### 2.12 BuildFreeform / AddNodes / ConvertToShape (フリーフォーム)

```python
# フリーフォームの構築
# msoEditingCorner = 0
builder = slide.Shapes.BuildFreeform(
    EditingType=0,   # msoEditingCorner
    X1=100,
    Y1=100
)

# ノードを追加
# msoSegmentLine = 0, msoSegmentCurve = 1
# msoEditingAuto = 0, msoEditingCorner = 0
builder.AddNodes(
    SegmentType=0,    # msoSegmentLine
    EditingType=0,    # msoEditingAuto
    X1=200,
    Y1=100
)

builder.AddNodes(
    SegmentType=0,    # msoSegmentLine
    EditingType=0,
    X1=200,
    Y1=200
)

builder.AddNodes(
    SegmentType=0,    # msoSegmentLine
    EditingType=0,
    X1=100,
    Y1=200
)

# 曲線セグメントの追加（ベジエ）
builder.AddNodes(
    SegmentType=1,    # msoSegmentCurve
    EditingType=0,    # msoEditingCorner
    X1=50, Y1=180,    # 制御点1
    X2=50, Y2=120,    # 制御点2
    X3=100, Y3=100    # 終点
)

# シェイプに変換
freeform = builder.ConvertToShape()
```

**BuildFreeform パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| EditingType | MsoEditingType | Yes | 最初の頂点の編集タイプ |
| X1 | Single | Yes | 最初の頂点X座標 |
| Y1 | Single | Yes | 最初の頂点Y座標 |

**AddNodes パラメータ:**
| パラメータ | 型 | 必須 | 説明 |
|---|---|---|---|
| SegmentType | MsoSegmentType | Yes | セグメントの種類 |
| EditingType | MsoEditingType | Yes | 編集タイプ |
| X1 | Single | Yes | X座標（直線の場合は終点、曲線の場合は制御点1） |
| Y1 | Single | Yes | Y座標 |
| X2 | Single | No | 曲線の制御点2のX座標 |
| Y2 | Single | No | 曲線の制御点2のY座標 |
| X3 | Single | No | 曲線の終点X座標 |
| Y3 | Single | No | 曲線の終点Y座標 |

**MsoSegmentType 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoSegmentLine | 0 | 直線セグメント |
| msoSegmentCurve | 1 | 曲線セグメント |

**MsoEditingType 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoEditingAuto | 0 | 自動 |
| msoEditingCorner | 1 | 角 |
| msoEditingSmooth | 2 | スムーズ |
| msoEditingSymmetric | 3 | 対称 |

---

### 2.13 Shapes.AddCallout / AddLabel / AddTextEffect

#### AddCallout (吹き出し)

```python
# msoCalloutTwo = 2 (2つのセグメントを持つ吹き出し)
callout = slide.Shapes.AddCallout(
    Type=2,     # msoCalloutTwo
    Left=100,
    Top=100,
    Width=200,
    Height=100
)
callout.TextFrame.TextRange.Text = "注釈テキスト"
```

**MsoCalloutType 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoCalloutOne | 1 | 1セグメント吹き出し |
| msoCalloutTwo | 2 | 2セグメント吹き出し |
| msoCalloutThree | 3 | 3セグメント吹き出し |
| msoCalloutFour | 4 | 4セグメント吹き出し |
| msoCalloutMixed | -2 | 混合 |

#### AddLabel (ラベル)

```python
# msoTextOrientationHorizontal = 1
label = slide.Shapes.AddLabel(
    Orientation=1,
    Left=100,
    Top=100,
    Width=200,
    Height=50
)
label.TextFrame.TextRange.Text = "ラベルテキスト"
```

#### AddTextEffect (ワードアート)

```python
# msoTextEffect1 = 0
text_effect = slide.Shapes.AddTextEffect(
    PresetTextEffect=0,     # msoTextEffect1
    Text="WordArt!",
    FontName="Impact",
    FontSize=36,
    FontBold=True,          # msoTrue
    FontItalic=False,       # msoFalse
    Left=100,
    Top=100
)
```

---

### 2.14 Shapes.AddPlaceholder

プレースホルダーを追加する。通常はスライドレイアウトに既に定義されているプレースホルダーを使用する。

```python
# 通常はレイアウトからプレースホルダーにアクセス
title_ph = slide.Shapes.Placeholders(1)  # タイトル
body_ph = slide.Shapes.Placeholders(2)   # 本文

# テキストを設定
title_ph.TextFrame.TextRange.Text = "スライドタイトル"
body_ph.TextFrame.TextRange.Text = "本文テキスト"

# プレースホルダーの確認
print(f"HasTitle: {slide.Shapes.HasTitle}")
if slide.Shapes.HasTitle:
    print(f"Title: {slide.Shapes.Title.TextFrame.TextRange.Text}")
```

**注意点:**
- `AddPlaceholder` は通常のスライドではなく、スライドマスターやレイアウトの編集時に使用。
- 一般的にはプレースホルダーは `Shapes.Placeholders` コレクションからアクセスする。
- プレースホルダーの種類は `PlaceholderFormat.Type` で確認可能。

---

### 2.15 GroupItems / Group / Ungroup (グループ化)

#### シェイプのグループ化

```python
# 複数のシェイプを追加
shape1 = slide.Shapes.AddShape(Type=1, Left=100, Top=100, Width=80, Height=60)
shape1.Name = "Rect1"
shape2 = slide.Shapes.AddShape(Type=9, Left=200, Top=100, Width=80, Height=60)
shape2.Name = "Oval1"

# ShapeRange を作成してグループ化
shape_range = slide.Shapes.Range(["Rect1", "Oval1"])
group = shape_range.Group()
print(f"グループ名: {group.Name}")
```

#### グループの解除

```python
# グループを解除
ungrouped = group.Ungroup()  # ShapeRange が返される
print(f"解除されたシェイプ数: {ungrouped.Count}")
```

#### グループ内のシェイプにアクセス

```python
# GroupItems でグループ内のシェイプにアクセス
for i in range(1, group.GroupItems.Count + 1):
    item = group.GroupItems(i)
    print(f"  {item.Name}: Type={item.Type}")
    # プロパティの変更は可能
    item.Fill.ForeColor.RGB = rgb_to_bgr(255, 0, 0)
```

#### 再グループ化

```python
# 解除したシェイプを再グループ化
regrouped = ungrouped.Regroup()
```

**注意点:**
- `ShapeRange.Group()` でグループ化。対象の ShapeRange には少なくとも2つのシェイプが必要。
- `Shape.Ungroup()` でグループ解除。ネストされたグループがある場合は再帰的に解除される。
- `GroupItems` でグループ内のシェイプにアクセスできるが、グループ内のシェイプを追加・削除することはできない（プロパティ変更のみ可能）。
- `ShapeRange.Regroup()` で再グループ化が可能。

---

### 2.16 Shapes.AddComment (コメント)

```python
# ※ PowerPoint のバージョンによって動作が異なる
# 従来のコメント
comment = slide.Shapes.AddComment()
comment.TextFrame.TextRange.Text = "コメントテキスト"
```

**注意点:**
- PowerPoint 2021 / Microsoft 365 では最新のコメント機能（モダンコメント）が導入されており、従来の `AddComment` とは異なる API になっている場合がある。
- 最新バージョンでは `Slide.Comments.Add` や `Slide.Comments.Add2` を使用することが推奨される場合がある。

---

## 3. Shape 個別操作

### 3.1 位置とサイズ: Left, Top, Width, Height

```python
shape = slide.Shapes(1)

# 位置の取得
print(f"Left: {shape.Left}, Top: {shape.Top}")
print(f"Width: {shape.Width}, Height: {shape.Height}")

# 位置の設定（ポイント単位。72ポイント = 1インチ ≈ 2.54cm）
shape.Left = 100     # 左端から100ポイント
shape.Top = 50       # 上端から50ポイント
shape.Width = 200    # 幅200ポイント
shape.Height = 150   # 高さ150ポイント

# インクリメンタルな移動
shape.IncrementLeft(10)   # 右に10ポイント移動
shape.IncrementTop(-5)    # 上に5ポイント移動
```

**単位について:**
- PowerPoint COM では位置・サイズは「ポイント」単位（72ポイント = 1インチ）。
- 標準スライドサイズ（16:9）: 幅 = 960ポイント（13.333インチ）、高さ = 540ポイント（7.5インチ）。
- 標準スライドサイズ（4:3）: 幅 = 720ポイント（10インチ）、高さ = 540ポイント（7.5インチ）。

---

### 3.2 回転: Rotation

```python
# 回転角度の取得（度数、時計回り）
print(f"現在の回転: {shape.Rotation}")

# 回転角度の設定
shape.Rotation = 45.0    # 時計回りに45度

# インクリメンタルな回転
shape.IncrementRotation(15)  # さらに15度回転

# 反時計回り（負の値を指定すると内部的に正の値に変換される）
# 例: -45.0 → 315.0
shape.Rotation = -45.0   # = 315度
```

**注意点:**
- 回転は時計回りの度数で指定（0-360）。
- 負の値を指定すると自動的に正の値に変換される（例: -45 → 315）。
- 回転の中心はシェイプの中心点。

---

### 3.3 ZOrder (前面・背面移動)

```python
# 最前面に移動
shape.ZOrder(0)   # msoBringToFront = 0

# 最背面に移動
shape.ZOrder(1)   # msoSendToBack = 1

# 1つ前面に移動
shape.ZOrder(2)   # msoBringForward = 2

# 1つ背面に移動
shape.ZOrder(3)   # msoSendBackward = 3

# 現在のZ位置を取得（読み取り専用）
z_pos = shape.ZOrderPosition
print(f"Z位置: {z_pos}")
```

**MsoZOrderCmd 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoBringToFront | 0 | 最前面に移動 |
| msoSendToBack | 1 | 最背面に移動 |
| msoBringForward | 2 | 1つ前面に移動 |
| msoSendBackward | 3 | 1つ背面に移動 |
| msoBringInFrontOfText | 4 | テキストの前面（Word のみ） |
| msoSendBehindText | 5 | テキストの背面（Word のみ） |

**注意点:**
- `msoBringInFrontOfText` (4) と `msoSendBehindText` (5) は Word でのみ使用。PowerPoint では使わない。
- `ZOrderPosition` は読み取り専用。1始まりで、値が大きいほど前面。

---

### 3.4 Name (名前の設定・取得)

```python
# 名前の取得
print(f"シェイプ名: {shape.Name}")

# 名前の設定
shape.Name = "MyCustomShape"

# 名前でシェイプにアクセス
my_shape = slide.Shapes("MyCustomShape")
```

**注意点:**
- デフォルト名は自動生成（例: "Rectangle 1", "TextBox 2"）。
- 名前はスライド内で一意であるべき（重複してもエラーにはならないが、アクセス時に最初のものが返される）。
- 名前には日本語も使用可能。

---

### 3.5 Visible プロパティ

```python
# 表示/非表示の取得
print(f"表示: {shape.Visible}")

# 非表示にする
shape.Visible = False   # msoFalse = 0

# 表示する
shape.Visible = True    # msoTrue = -1
```

**注意点:**
- MsoTriState 値: `msoTrue = -1`, `msoFalse = 0`, `msoCTrue = 1`。
- Python からは `True`/`False` で設定可能（win32com が自動変換）。
- 非表示のシェイプはスライドショーでは表示されないが、編集画面では半透明で表示される。

---

### 3.6 Copy, Cut, Delete

```python
# コピー（クリップボードにコピー）
shape.Copy()

# 切り取り（クリップボードに移動）
shape.Cut()

# 削除
shape.Delete()

# ペースト（Shapes コレクションにペースト）
pasted = slide.Shapes.Paste()  # ShapeRange が返される
```

**注意点:**
- `Copy`/`Cut` はクリップボードを使用する。
- 別のスライドにペーストする場合は、対象スライドの `Shapes.Paste()` を呼ぶ。
- `PasteSpecial` で形式を指定してペーストも可能。

---

### 3.7 Duplicate

```python
# シェイプを複製
dup = shape.Duplicate()

# 複製されたシェイプの位置を調整
dup.Left = shape.Left + 50
dup.Top = shape.Top + 50
```

**注意点:**
- `Duplicate` は同じスライド内にコピーを作成する。
- 戻り値は `ShapeRange` オブジェクト。
- 複製されたシェイプは元のシェイプと同じプロパティを持つ（位置は若干ずれる）。

---

### 3.8 Flip (水平・垂直反転)

```python
# 水平反転
# msoFlipHorizontal = 0
shape.Flip(0)

# 垂直反転
# msoFlipVertical = 1
shape.Flip(1)

# 反転状態の確認（読み取り専用プロパティ）
print(f"水平反転: {shape.HorizontalFlip}")  # MsoTriState
print(f"垂直反転: {shape.VerticalFlip}")    # MsoTriState
```

**MsoFlipCmd 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| msoFlipHorizontal | 0 | 水平方向に反転 |
| msoFlipVertical | 1 | 垂直方向に反転 |

---

### 3.9 PickUp / Apply (書式のコピー＆貼り付け)

```python
# シェイプ1の書式をコピー
slide.Shapes(1).PickUp()

# シェイプ2に書式を適用
slide.Shapes(2).Apply()
```

**注意点:**
- `PickUp` はシェイプの書式（塗りつぶし、線、効果など）をコピーする（テキストの書式は含まない）。
- `Apply` で別のシェイプにコピーした書式を適用する。
- `PickUp` → `Apply` のペアで使用する。`PickUp` を呼ばずに `Apply` を呼ぶと前回の書式が適用される。

#### PickupAnimation / ApplyAnimation

```python
# アニメーション設定のコピー
slide.Shapes(1).PickupAnimation()
slide.Shapes(2).ApplyAnimation()
```

---

### 3.10 ActionSettings (ハイパーリンク・マクロ実行)

```python
# ppMouseClick = 1, ppMouseOver = 2
click_action = shape.ActionSettings(1)   # クリック時の動作
hover_action = shape.ActionSettings(2)   # マウスオーバー時の動作

# ハイパーリンクを設定
# ppActionHyperlink = 7
click_action.Action = 7
click_action.Hyperlink.Address = "https://www.example.com"

# 別のスライドへのリンク
# ppActionNamedSlideShow = 6 (ただし通常は Hyperlink で設定)
click_action.Action = 7  # ppActionHyperlink
click_action.Hyperlink.SubAddress = "3"  # スライド番号

# マクロの実行
# ppActionRunMacro = 8
click_action.Action = 8
click_action.Run = "MyMacroName"

# プログラムの実行
# ppActionRunProgram = 9
click_action.Action = 9
click_action.Run = "notepad.exe"

# 次のスライドに移動
# ppActionNextSlide = 1
click_action.Action = 1

# アニメーション付きアクション
click_action.AnimateAction = True

# サウンド再生
click_action.SoundEffect.ImportFromFile(r"C:\sound\click.wav")
```

**PpActionType 定数:**
| 定数 | 値 | 説明 |
|---|---|---|
| ppActionNone | 0 | アクションなし |
| ppActionNextSlide | 1 | 次のスライド |
| ppActionPreviousSlide | 2 | 前のスライド |
| ppActionFirstSlide | 3 | 最初のスライド |
| ppActionLastSlide | 4 | 最後のスライド |
| ppActionLastSlideViewed | 5 | 最後に表示したスライド |
| ppActionEndShow | 6 | スライドショー終了 |
| ppActionHyperlink | 7 | ハイパーリンク |
| ppActionRunMacro | 8 | マクロ実行 |
| ppActionRunProgram | 9 | プログラム実行 |
| ppActionNamedSlideShow | 10 | 名前付きスライドショー |
| ppActionOLEVerb | 11 | OLE動詞 |
| ppActionMixed | -2 | 混合 |

---

### 3.11 AnimationSettings / アニメーション効果

```python
# AnimationSettings（旧アニメーション API）
anim = shape.AnimationSettings
anim.Animate = True

# エントリーエフェクト
# ppEffectFlyFromLeft = 3844
anim.EntryEffect = 3844
anim.TextLevelEffect = 1  # ppAnimateByAllLevels

# アニメーション順序
anim.AnimationOrder = 1

# アニメーションの進行方法
# ppAdvanceOnClick = 0, ppAdvanceOnTime = 2, ppAdvanceModeMixed = -2
anim.AdvanceMode = 0
anim.AdvanceTime = 2  # 秒

# サウンド
anim.SoundEffect.ImportFromFile(r"C:\sound\fly.wav")

# テキストのアニメーション
# ppAnimateByFirstLevel = 1, ppAnimateBySecondLevel = 2, etc.
anim.TextLevelEffect = 1
# ppAnimateByAllLevels = 16
anim.TextUnitEffect = 0  # ppAnimateByParagraph
```

**AnimationSettings 主要プロパティ:**
| プロパティ | 型 | 説明 |
|---|---|---|
| Animate | MsoTriState | アニメーションを有効にするか |
| AnimateBackground | MsoTriState | 背景をアニメーションするか |
| AnimateTextInReverse | MsoTriState | テキストを逆順でアニメーション |
| AnimationOrder | Long | アニメーション順序 |
| AdvanceMode | PpAdvanceMode | 進行方法 |
| AdvanceTime | Single | 自動進行時間（秒） |
| EntryEffect | PpEntryEffect | エントリーエフェクト |
| SoundEffect | SoundEffect | サウンド効果 |
| TextLevelEffect | PpTextLevelEffect | テキストレベルエフェクト |
| TextUnitEffect | PpTextUnitEffect | テキストユニットエフェクト |

**注意点:**
- `AnimationSettings` は旧来の API で、PowerPoint 2002 以降の「カスタムアニメーション」には `TimeLine` オブジェクトを使用する。
- `EntryEffect` を設定しても `Animate = True` かつ `TextLevelEffect` が `ppAnimateLevelNone` でなければアニメーションは実行されない。
- 最新のアニメーション制御には `Slide.TimeLine.MainSequence` を使用することが推奨される。

---

### 3.12 Selection (選択状態の操作)

```python
# シェイプを選択
shape.Select()

# 現在の選択を取得
selection = app.ActiveWindow.Selection

# 選択の種類
# ppSelectionNone = 0, ppSelectionSlides = 1, ppSelectionShapes = 2, ppSelectionText = 3
sel_type = selection.Type

# 選択されたシェイプにアクセス
if selection.Type == 2:  # ppSelectionShapes
    for i in range(1, selection.ShapeRange.Count + 1):
        print(f"選択シェイプ: {selection.ShapeRange(i).Name}")

# 追加選択（Shift+クリック相当）
shape2.Select(Replace=False)  # msoFalse = 0 で追加選択
```

**注意点:**
- `Shape.Select()` はウィンドウが表示されている場合のみ動作。
- バックグラウンド処理では `Select` は使わず、直接オブジェクトを操作する方が良い。

---

### 3.13 Tags (カスタムタグの追加)

```python
# タグの追加
shape.Tags.Add("CATEGORY", "Header")
shape.Tags.Add("PRIORITY", "High")

# タグの取得（名前は自動的に大文字に変換される）
value = shape.Tags("CATEGORY")  # "Header"
# または
value = shape.Tags.Item("CATEGORY")

# タグ数の取得
count = shape.Tags.Count

# 全タグの列挙
for i in range(1, shape.Tags.Count + 1):
    name = shape.Tags.Name(i)
    value = shape.Tags.Value(i)
    print(f"  {name} = {value}")

# タグの削除
shape.Tags.Delete("CATEGORY")

# スライドにもタグを付けられる
slide.Tags.Add("SECTION", "Introduction")

# プレゼンテーションにもタグを付けられる
prs.Tags.Add("VERSION", "1.0")
```

**注意点:**
- タグ名は内部的に大文字で保存される。
- タグ値は文字列のみ（数値は文字列に変換して保存）。
- タグはユーザーからは見えない隠しメタデータ。
- MCP サーバーでシェイプの識別やメタデータ管理に非常に有用。

---

### 3.14 ScaleHeight / ScaleWidth

```python
# 現在のサイズの150%にスケール
# msoTrue = -1, msoFalse = 0
# msoScaleFromTopLeft = 0, msoScaleFromMiddle = 1, msoScaleFromBottomRight = 2
shape.ScaleHeight(Factor=1.5, RelativeToOriginalSize=False, fScale=0)
shape.ScaleWidth(Factor=1.5, RelativeToOriginalSize=False, fScale=0)

# 元のサイズに対して50%にスケール
shape.ScaleHeight(Factor=0.5, RelativeToOriginalSize=True)
```

---

### 3.15 LockAspectRatio

```python
# アスペクト比のロック
shape.LockAspectRatio = True   # msoTrue
# ロック状態で Width を変更すると Height も自動調整される
shape.Width = 300
```

---

### 3.16 Fill / Line (塗りつぶし・線)

```python
# 塗りつぶし
shape.Fill.Solid()
shape.Fill.ForeColor.RGB = rgb_to_bgr(255, 128, 0)  # オレンジ
shape.Fill.Transparency = 0.3  # 30%透明

# グラデーション
shape.Fill.TwoColorGradient(Style=1, Variant=1)
shape.Fill.ForeColor.RGB = rgb_to_bgr(255, 0, 0)
shape.Fill.BackColor.RGB = rgb_to_bgr(255, 255, 0)

# 線のスタイル
shape.Line.Weight = 2.0
shape.Line.ForeColor.RGB = rgb_to_bgr(0, 0, 0)
shape.Line.DashStyle = 1   # msoLineSolid
shape.Line.Visible = True

# 線なし
shape.Line.Visible = False
```

---

### 3.17 Shadow / Glow / SoftEdge / Reflection / ThreeD

```python
# 影
shape.Shadow.Visible = True
shape.Shadow.Type = 1    # msoShadow1
shape.Shadow.ForeColor.RGB = rgb_to_bgr(128, 128, 128)
shape.Shadow.OffsetX = 5
shape.Shadow.OffsetY = 5

# 光彩
shape.Glow.Color.RGB = rgb_to_bgr(255, 255, 0)
shape.Glow.Radius = 10
shape.Glow.Transparency = 0.5

# ぼかし
shape.SoftEdge.Type = 1  # msoSoftEdgeType1
shape.SoftEdge.Radius = 5

# 反射
shape.Reflection.Type = 1  # msoReflectionType1

# 3D 効果
shape.ThreeD.Visible = True
shape.ThreeD.BevelTopType = 1  # msoBevelCircle
```

---

### 3.18 Export (シェイプの画像エクスポート)

```python
# シェイプを画像としてエクスポート
# ppShapeFormatPNG = 2, ppShapeFormatJPG = 1, ppShapeFormatBMP = 3
shape.Export(
    PathName=r"C:\exports\shape.png",
    Filter=2  # ppShapeFormatPNG
)
```

---

### 3.19 SetShapesDefaultProperties

```python
# 現在のシェイプの書式をデフォルトに設定
shape.SetShapesDefaultProperties()
# 以降に追加されるシェイプはこの書式がデフォルトになる
```

---

## 4. Shape の種類 (MsoShapeType)

### MsoShapeType 完全一覧

| 定数 | 値 | 説明 | 備考 |
|---|---|---|---|
| msoAutoShape | 1 | オートシェイプ | 四角形、円、矢印等 |
| msoCallout | 2 | 吹き出し | AddCallout で作成 |
| msoChart | 3 | グラフ | AddChart2 で作成 |
| msoComment | 4 | コメント | AddComment で作成 |
| msoFreeform | 5 | フリーフォーム | BuildFreeform + ConvertToShape |
| msoGroup | 6 | グループ | ShapeRange.Group で作成 |
| msoEmbeddedOLEObject | 7 | 埋め込み OLE | AddOLEObject で作成 |
| msoFormControl | 8 | フォームコントロール | VBA フォーム用 |
| msoLine | 9 | 線 | AddLine で作成 |
| msoLinkedOLEObject | 10 | リンク OLE | AddOLEObject(Link=True) |
| msoLinkedPicture | 11 | リンク画像 | AddPicture(LinkToFile=True) |
| msoOLEControlObject | 12 | OLE コントロール | ActiveX コントロール |
| msoPicture | 13 | 画像 | AddPicture で作成 |
| msoPlaceholder | 14 | プレースホルダー | レイアウトで定義 |
| msoTextEffect | 15 | テキストエフェクト | ワードアート |
| msoMedia | 16 | メディア | 動画・音声 |
| msoTextBox | 17 | テキストボックス | AddTextbox で作成 |
| msoScriptAnchor | 18 | スクリプトアンカー | 非推奨 |
| msoTable | 19 | テーブル | AddTable で作成 |
| msoCanvas | 20 | キャンバス | 描画キャンバス |
| msoDiagram | 21 | ダイアグラム | 旧ダイアグラム |
| msoInk | 22 | インク | ペン入力 |
| msoInkComment | 23 | インクコメント | ペンコメント |
| msoIgxGraphic | 24 | SmartArt | AddSmartArt で作成 |
| msoSlicer | 25 | スライサー | Excel 用 |
| msoWebVideo | 26 | Web ビデオ | オンライン動画 |
| msoContentApp | 27 | コンテンツ アドイン | Office アドイン |
| msoGraphic | 28 | グラフィック | SVG 画像等 |
| msoLinkedGraphic | 29 | リンクグラフィック | リンクされた SVG |
| mso3DModel | 30 | 3D モデル | 3D オブジェクト |
| msoLinked3DModel | 31 | リンク 3D モデル | リンクされた 3D |
| msoShapeTypeMixed | -2 | 混合 | 複数選択時の戻り値 |

### Type プロパティでの判別

```python
# シェイプの種類を判別
shape = slide.Shapes(1)
shape_type = shape.Type

if shape_type == 1:    # msoAutoShape
    print(f"オートシェイプ: {shape.AutoShapeType}")
elif shape_type == 17:  # msoTextBox
    print(f"テキストボックス: {shape.TextFrame.TextRange.Text}")
elif shape_type == 13:  # msoPicture
    print("画像")
elif shape_type == 19:  # msoTable
    print(f"テーブル: {shape.Table.Rows.Count}行 x {shape.Table.Columns.Count}列")
elif shape_type == 14:  # msoPlaceholder
    print(f"プレースホルダー: {shape.PlaceholderFormat.Type}")
elif shape_type == 6:   # msoGroup
    print(f"グループ: {shape.GroupItems.Count}個のシェイプ")
elif shape_type == 3:   # msoChart
    print("グラフ")
elif shape_type == 16:  # msoMedia
    print(f"メディア: {shape.MediaType}")
```

### HasTextFrame, HasTable, HasChart での安全な確認

```python
# テキストフレームの有無を確認
if shape.HasTextFrame:
    text = shape.TextFrame.TextRange.Text
    print(f"テキスト: {text}")

# テーブルの有無を確認
if shape.HasTable:
    table = shape.Table
    print(f"テーブル: {table.Rows.Count}x{table.Columns.Count}")

# グラフの有無を確認
if shape.HasChart:
    chart = shape.Chart
    print(f"グラフタイプ: {chart.ChartType}")

# SmartArt の有無を確認
if shape.HasSmartArt:
    smart_art = shape.SmartArt
    print(f"SmartArt ノード数: {smart_art.AllNodes.Count}")
```

---

## 5. 主要定数・列挙型一覧

### MsoTriState

| 定数 | 値 | 説明 |
|---|---|---|
| msoTrue | -1 | True |
| msoFalse | 0 | False |
| msoCTrue | 1 | True（一部のメソッドで使用） |
| msoTriStateToggle | -3 | トグル（反転） |
| msoTriStateMixed | -2 | 混合 |

### PpSlideSize (スライドサイズ)

| 定数 | 値 | 説明 |
|---|---|---|
| ppSlideSizeOnScreen | 1 | 画面サイズ（4:3） |
| ppSlideSizeLetterPaper | 2 | レターサイズ |
| ppSlideSizeA4Paper | 3 | A4 |
| ppSlideSizeCustom | 7 | カスタム |
| ppSlideSizeOnScreen16x9 | 9 | 画面サイズ（16:9） |

### PpPlaceholderType (プレースホルダーの種類)

| 定数 | 値 | 説明 |
|---|---|---|
| ppPlaceholderTitle | 1 | タイトル |
| ppPlaceholderBody | 2 | 本文 |
| ppPlaceholderCenterTitle | 3 | 中央タイトル |
| ppPlaceholderSubtitle | 4 | サブタイトル |
| ppPlaceholderDate | 10 | 日付 |
| ppPlaceholderSlideNumber | 12 | スライド番号 |
| ppPlaceholderFooter | 11 | フッター |
| ppPlaceholderHeader | 13 | ヘッダー |

### MsoLineDashStyle (線の破線スタイル)

| 定数 | 値 | 説明 |
|---|---|---|
| msoLineSolid | 1 | 実線 |
| msoLineSquareDot | 2 | 正方形点線 |
| msoLineRoundDot | 3 | 丸点線 |
| msoLineDash | 4 | 破線 |
| msoLineDashDot | 5 | 一点鎖線 |
| msoLineDashDotDot | 6 | 二点鎖線 |
| msoLineLongDash | 7 | 長い破線 |
| msoLineLongDashDot | 8 | 長い一点鎖線 |

### MsoArrowheadStyle (矢印スタイル)

| 定数 | 値 | 説明 |
|---|---|---|
| msoArrowheadNone | 1 | 矢印なし |
| msoArrowheadTriangle | 2 | 三角形 |
| msoArrowheadOpen | 3 | 開いた矢印 |
| msoArrowheadStealth | 4 | ステルス |
| msoArrowheadDiamond | 5 | ひし形 |
| msoArrowheadOval | 6 | 楕円 |

---

## 6. Python win32com での定数利用方法

### 方法1: 数値を直接使用（最も簡単）

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
prs = app.Presentations.Add()

# 定数の数値を直接指定
ppLayoutBlank = 12
msoShapeRectangle = 1

slide = prs.Slides.Add(1, ppLayoutBlank)
shape = slide.Shapes.AddShape(msoShapeRectangle, 100, 100, 200, 100)
```

### 方法2: win32com.client.constants を使用

```python
import win32com.client

# EnsureDispatch を使用すると定数が利用可能になる
app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")

from win32com.client import constants

# 定数名で参照可能
slide = prs.Slides.Add(1, constants.ppLayoutBlank)
shape = slide.Shapes.AddShape(constants.msoShapeRectangle, 100, 100, 200, 100)
```

### 方法3: makepy を事前実行

```bash
# コマンドラインで実行
python -m win32com.client.makepy "Microsoft PowerPoint 16.0 Object Library"
python -m win32com.client.makepy "Microsoft Office 16.0 Object Library"
```

```python
import win32com.client
# makepy 実行後は gencache.EnsureDispatch で自動的に定数が利用可能
app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
```

### 方法4: 定数を自分で定義（MCP サーバー推奨）

```python
# MCP サーバーでは定数を明示的に定義するのが最も安定
class PPTConstants:
    """PowerPoint COM 定数"""
    # PpSlideLayout
    ppLayoutTitle = 1
    ppLayoutText = 2
    ppLayoutBlank = 12
    ppLayoutTitleOnly = 11
    ppLayoutSectionHeader = 33
    ppLayoutComparison = 34

    # MsoAutoShapeType
    msoShapeRectangle = 1
    msoShapeRoundedRectangle = 5
    msoShapeOval = 9
    msoShapeRightArrow = 33
    msoShape5pointStar = 92

    # MsoShapeType
    msoAutoShape = 1
    msoTextBox = 17
    msoPicture = 13
    msoTable = 19
    msoChart = 3
    msoGroup = 6
    msoPlaceholder = 14
    msoMedia = 16
    msoFreeform = 5
    msoLine = 9

    # MsoTriState
    msoTrue = -1
    msoFalse = 0

    # MsoConnectorType
    msoConnectorStraight = 1
    msoConnectorElbow = 2
    msoConnectorCurve = 3

    # MsoZOrderCmd
    msoBringToFront = 0
    msoSendToBack = 1
    msoBringForward = 2
    msoSendBackward = 3

    # MsoFlipCmd
    msoFlipHorizontal = 0
    msoFlipVertical = 1

    # MsoTextOrientation
    msoTextOrientationHorizontal = 1
    msoTextOrientationVertical = 5

    # PpActionType
    ppActionNone = 0
    ppActionNextSlide = 1
    ppActionHyperlink = 7
    ppActionRunMacro = 8
```

---

## 7. MCP サーバー実装への推奨事項

### 7.1 優先度: 高（必須実装）

| 機能カテゴリ | 具体的な操作 | 理由 |
|---|---|---|
| スライド管理 | `Slides.Add` / `Slides.AddSlide` | 基本操作 |
| スライド管理 | `Slide.Delete`, `Slide.Duplicate`, `Slide.MoveTo` | 基本操作 |
| スライド管理 | スライドの一覧取得（Count, Name, Index） | 基本操作 |
| スライド管理 | `NotesPage` テキストの読み書き | よく使われる |
| シェイプ追加 | `Shapes.AddShape` (オートシェイプ) | 最もよく使われる |
| シェイプ追加 | `Shapes.AddTextbox` | テキスト挿入の基本 |
| シェイプ追加 | `Shapes.AddPicture` | 画像挿入は必須 |
| シェイプ追加 | `Shapes.AddTable` | テーブルは頻繁に使用 |
| シェイプ操作 | Left, Top, Width, Height | 位置・サイズ制御 |
| シェイプ操作 | Name の取得・設定 | シェイプ識別に必須 |
| シェイプ操作 | Delete | 基本操作 |
| シェイプ操作 | テキストの読み書き (TextFrame.TextRange) | 最重要機能 |
| シェイプ操作 | Fill, Line の設定 | 書式設定の基本 |
| シェイプ情報 | Type の取得、シェイプ一覧 | 情報取得に必須 |
| プレースホルダー | Placeholders へのアクセス | テンプレート活用に必要 |

### 7.2 優先度: 中（推奨実装）

| 機能カテゴリ | 具体的な操作 | 理由 |
|---|---|---|
| スライド管理 | `Slide.Layout` / `Slide.CustomLayout` の変更 | レイアウト制御 |
| スライド管理 | `Background.Fill` の設定 | デザイン機能 |
| スライド管理 | `SlideShowTransition` の設定 | プレゼンテーション品質 |
| スライド管理 | プレゼンテーション間のスライドコピー | 便利機能 |
| シェイプ追加 | `Shapes.AddChart2` | グラフは重要 |
| シェイプ追加 | `Shapes.AddLine` | 基本図形 |
| シェイプ追加 | `Shapes.AddConnector` | フローチャート等 |
| シェイプ操作 | Rotation | 回転制御 |
| シェイプ操作 | ZOrder | レイヤー制御 |
| シェイプ操作 | Duplicate, Copy | 複製操作 |
| シェイプ操作 | Flip | 反転操作 |
| シェイプ操作 | Tags | メタデータ管理 |
| グループ | Group / Ungroup / GroupItems | グループ操作 |
| シェイプ操作 | Shadow, Glow, SoftEdge | 視覚効果 |

### 7.3 優先度: 低（将来実装）

| 機能カテゴリ | 具体的な操作 | 理由 |
|---|---|---|
| シェイプ追加 | `Shapes.AddSmartArt` | 複雑な操作 |
| シェイプ追加 | `Shapes.AddMediaObject2` | メディア挿入 |
| シェイプ追加 | `Shapes.AddOLEObject` | OLE は複雑 |
| シェイプ追加 | `BuildFreeform` / `AddCurve` / `AddPolyline` | 高度な描画 |
| シェイプ追加 | `Shapes.AddTextEffect` | ワードアート |
| シェイプ追加 | `Shapes.AddCallout` / `AddLabel` | 特殊用途 |
| シェイプ操作 | ActionSettings | ハイパーリンク等 |
| シェイプ操作 | AnimationSettings | アニメーション |
| シェイプ操作 | PickUp / Apply | 書式コピー |
| シェイプ操作 | Selection 操作 | UI 依存の操作 |
| シェイプ操作 | Export | 画像出力 |
| シェイプ操作 | ScaleHeight / ScaleWidth | スケーリング |

### 7.4 実装上の注意点

1. **単位系**: PowerPoint COM はポイント単位（72pt = 1inch）。MCP ツールのインターフェースではセンチメートルやインチも受け付けて内部変換すると使いやすい。

2. **色の指定**: COM は BGR 形式（`0xBBGGRR`）。MCP ツールでは `#RRGGBB` の HTML カラーコードを受け付けて内部変換するのが望ましい。

3. **定数管理**: `win32com.client.gencache.EnsureDispatch` に依存せず、必要な定数は自前で定義するのが安定的。

4. **エラーハンドリング**: COM 操作は例外が発生しやすい。特に以下のケースに注意:
   - 存在しないインデックスへのアクセス
   - 削除されたオブジェクトへの参照
   - HasTextFrame = False のシェイプに TextFrame をアクセス
   - ファイルパスの存在確認（AddPicture, AddMediaObject2 等）

5. **パフォーマンス**: 大量のシェイプ操作時は `Application.ScreenUpdating = False` を設定し、操作完了後に `True` に戻す。

6. **COM オブジェクトの解放**: Python の `win32com` は参照カウント方式だが、明示的に `None` を代入するか `del` で解放するのが安全。

---

## 参考資料

- [Shapes object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes)
- [Shape object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape)
- [PpSlideLayout enumeration - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout)
- [MsoShapeType enumeration - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/office.msoshapetype)
- [MsoAutoShapeType enumeration - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/office.msoautoshapetype)
- [SlideShowTransition object - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.SlideShowTransition)
- [ActionSettings object - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.actionsettings)
- [AnimationSettings object - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.AnimationSettings)
- [Shapes.AddConnector method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addconnector)
- [Shapes.AddChart2 method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2)
- [Shapes.AddTable method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTable)
- [Shapes.AddPicture method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addpicture)
- [Shapes.AddMediaObject2 method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addmediaobject2)
- [Shapes.AddOLEObject method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addoleobject)
- [Shapes.AddSmartArt method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addsmartart)
- [Slides.InsertFromFile method - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.Slides.InsertFromFile)
- [Tags object (PowerPoint) - Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.tags)
- [WIN32 automation of PowerPoint (GitHub Gist)](https://gist.github.com/dmahugh/f642607d50cd008cc752f1344e9809e6)
- [Automating PowerPoint with Python - S Anand](https://www.s-anand.net/blog/automating-powerpoint-with-python/)
- [Controlling PowerPoint w/ Python via COM32 - Medium](https://medium.com/@chasekidder/controlling-powerpoint-w-python-52f6f6bf3f2d)
