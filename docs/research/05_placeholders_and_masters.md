# PowerPoint COM 自動化: プレースホルダー、スライドマスター、レイアウトの完全リファレンス

> **対象**: PowerPoint COM (win32com) を使用した MCP サーバー構築のための技術調査
> **最終更新**: 2026-02-16
> **重要度**: 最高（ユーザーがプレースホルダーを多用するため）

---

## 目次

1. [SlideMaster（スライドマスター）](#1-slidemasterスライドマスター)
2. [CustomLayout（カスタムレイアウト）](#2-customlayoutカスタムレイアウト)
3. [Placeholders（プレースホルダー）](#3-placeholdersプレースホルダー)
4. [継承構造 (Master → Layout → Slide)](#4-継承構造-master--layout--slide)
5. [HeadersFooters（ヘッダー・フッター）](#5-headersfootersヘッダーフッター)
6. [実践的な操作パターン](#6-実践的な操作パターン)
7. [定数・列挙型一覧](#7-定数列挙型一覧)

---

## 1. SlideMaster（スライドマスター）

### 1.1 概要

スライドマスター（Master オブジェクト）は、プレゼンテーション内の全スライドの外観を統一的に制御するための基盤オブジェクトである。マスターには以下の種類がある:

- **SlideMaster**: 通常のスライドマスター
- **TitleMaster**: タイトルスライドマスター（レガシー）
- **NotesMaster**: ノートマスター
- **HandoutMaster**: 配布資料マスター

COM オブジェクトモデルでは、これらはすべて `Master` オブジェクトとして返される。

### 1.2 Presentation.SlideMaster

**パス**: `Presentation.SlideMaster`
**戻り値**: `Master` オブジェクト
**説明**: プレゼンテーションのデフォルトのスライドマスターを返す。複数のデザインがある場合は、最初のデザインのマスターが返される。

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
prs = app.Presentations.Open(r"C:\path\to\presentation.pptx")

# スライドマスターを取得
master = prs.SlideMaster
print(f"マスター名: {master.Name}")
print(f"幅: {master.Width} pt, 高さ: {master.Height} pt")
```

**注意点**:
- `Presentation.SlideMaster` は読み取り専用プロパティ
- 複数のマスターがある場合、最初のデザインのマスターのみ返す
- 複数マスターを扱うには `Presentation.Designs` コレクションを使用する

### 1.3 Presentation.Designs コレクション

**パス**: `Presentation.Designs`
**戻り値**: `Designs` コレクション
**説明**: プレゼンテーション内のすべてのデザインテンプレート（スライドマスター）のコレクション。PowerPoint 2002 以降、1つのプレゼンテーションに複数のデザインを持つことが可能。

```python
# 全デザインを列挙
for i in range(1, prs.Designs.Count + 1):
    design = prs.Designs.Item(i)
    print(f"Design {i}: {design.Name}")
    print(f"  SlideMaster: {design.SlideMaster.Name}")
    print(f"  HasTitleMaster: {design.HasTitleMaster}")
```

#### Designs コレクションのメソッド

| メソッド | 説明 | 構文 |
|---------|------|------|
| `Add` | 新しいデザインを追加 | `Designs.Add(designName)` |
| `Clone` | 既存デザインを複製 | `Designs.Clone(pOriginal, [Index])` |
| `Item` | インデックスまたは名前でデザインを取得 | `Designs.Item(Index)` |
| `Load` | テンプレートファイルからデザインを読み込み | `Designs.Load(TemplateName, [Index])` |

```python
# 新しいデザインを追加
new_design = prs.Designs.Add(designName="MyNewDesign")

# テンプレートからデザインを読み込み
# prs.Designs.Load(r"C:\path\to\template.potx")

# デザインを複製
original_design = prs.Designs(1)
cloned = prs.Designs.Clone(pOriginal=original_design, Index=2)

# デザインの名前で取得
# design = prs.Designs.Item("MyDesign")
```

#### Designs コレクションのプロパティ

| プロパティ | 説明 |
|-----------|------|
| `Count` | デザインの数 |
| `Application` | Application オブジェクト |
| `Parent` | 親オブジェクト |

### 1.4 Design.SlideMaster

**パス**: `Design.SlideMaster`
**戻り値**: `Master` オブジェクト
**説明**: 各 Design オブジェクトから対応するスライドマスターを取得する。

```python
# 特定のデザインからマスターを取得
design = prs.Designs(1)
master = design.SlideMaster

# そのマスターの背景を設定
# msoGradientHorizontal = 1, msoGradientBrass = 20
master.Background.Fill.PresetGradient(1, 1, 20)
```

**Design オブジェクトのプロパティ**:

| プロパティ | 説明 |
|-----------|------|
| `Name` | デザイン名 |
| `Index` | コレクション内のインデックス |
| `SlideMaster` | スライドマスター（Master オブジェクト） |
| `HasTitleMaster` | タイトルマスターの有無 |
| `Preserved` | テーマ変更時にデザインを保持するか |
| `Application` | Application オブジェクト |

**Design オブジェクトのメソッド**:

| メソッド | 説明 |
|---------|------|
| `AddTitleMaster` | タイトルマスターを追加 |
| `Delete` | デザインを削除 |
| `MoveTo(toPos)` | デザインの位置を移動 |

### 1.5 Master オブジェクトの全プロパティ

**パス**: `Master.*`

| プロパティ | 型 | 説明 |
|-----------|-----|------|
| `Application` | Application | Application オブジェクト |
| `Background` | ShapeRange | 背景のシェイプ情報 |
| `BackgroundStyle` | MsoBackgroundStyleIndex | 背景スタイル |
| `ColorScheme` | ColorScheme | カラースキーム（レガシー） |
| `CustomerData` | CustomerData | カスタムデータ |
| `CustomLayouts` | CustomLayouts | カスタムレイアウトコレクション |
| `Design` | Design | 親デザインオブジェクト |
| `Guides` | Guides | ガイド線 |
| `HeadersFooters` | HeadersFooters | ヘッダー・フッター |
| `Height` | Single | 高さ（ポイント） |
| `Hyperlinks` | Hyperlinks | ハイパーリンク |
| `Name` | String | マスター名 |
| `Parent` | Object | 親オブジェクト |
| `Shapes` | Shapes | マスター上のシェイプ |
| `SlideShowTransition` | SlideShowTransition | スライドショー切り替え効果 |
| `TextStyles` | TextStyles | テキストスタイル（タイトル/本文/デフォルト） |
| `Theme` | Theme | テーマオブジェクト |
| `TimeLine` | TimeLine | アニメーションタイムライン |
| `Width` | Single | 幅（ポイント） |

**Master オブジェクトのメソッド**:

| メソッド | 説明 |
|---------|------|
| `ApplyTheme(themeName)` | テーマファイル（.thmx）を適用 |
| `Delete` | マスターを削除 |

### 1.6 SlideMaster.Shapes（マスター上のシェイプ）

マスター上の Shapes コレクションには、マスターに配置された背景画像、ロゴ、プレースホルダーなどが含まれる。

```python
master = prs.SlideMaster

# マスター上の全シェイプを列挙
for i in range(1, master.Shapes.Count + 1):
    shp = master.Shapes(i)
    print(f"Shape {i}: Name={shp.Name}, Type={shp.Type}")

    # プレースホルダーかどうか確認
    # msoPlaceholder = 14
    if shp.Type == 14:
        pf = shp.PlaceholderFormat
        print(f"  Placeholder Type: {pf.Type}")

# マスターのタイトルシェイプ（存在する場合）
try:
    title = master.Shapes.Title
    print(f"Title: {title.Name}")
except:
    print("マスターにタイトルなし")
```

### 1.7 SlideMaster.CustomLayouts コレクション

**パス**: `Master.CustomLayouts`
**戻り値**: `CustomLayouts` コレクション
**説明**: マスターに関連付けられたすべてのカスタムレイアウトを返す。

```python
master = prs.SlideMaster

# レイアウト一覧を取得
print(f"レイアウト数: {master.CustomLayouts.Count}")
for i in range(1, master.CustomLayouts.Count + 1):
    layout = master.CustomLayouts(i)
    print(f"  Layout {i}: Name='{layout.Name}', Index={layout.Index}")

    # レイアウト上のプレースホルダーを列挙
    placeholders = layout.Shapes.Placeholders
    for j in range(1, placeholders.Count + 1):
        ph = placeholders(j)
        print(f"    Placeholder {j}: Type={ph.PlaceholderFormat.Type}, Name='{ph.Name}'")
```

### 1.8 SlideMaster.Theme

**パス**: `Master.Theme`
**戻り値**: `Theme` オブジェクト
**説明**: マスターに適用されているテーマを返す。テーマにはカラースキーム、フォント、エフェクトが含まれる。

```python
master = prs.SlideMaster
theme = master.Theme

# テーマカラースキーム
color_scheme = theme.ThemeColorScheme
for i in range(1, color_scheme.Count + 1):
    color = color_scheme(i)
    print(f"  Color {i}: RGB={color.RGB}")

# テーマをファイルから適用
# master.ApplyTheme(r"C:\path\to\theme.thmx")
```

**テーマカラーのインデックス定数** (`MsoThemeColorIndex`):

| 定数 | 値 | 説明 |
|------|-----|------|
| `msoThemeColorDark1` | 1 | 濃色1（通常は黒系） |
| `msoThemeColorLight1` | 2 | 淡色1（通常は白系） |
| `msoThemeColorDark2` | 3 | 濃色2 |
| `msoThemeColorLight2` | 4 | 淡色2 |
| `msoThemeColorAccent1` | 5 | アクセント1 |
| `msoThemeColorAccent2` | 6 | アクセント2 |
| `msoThemeColorAccent3` | 7 | アクセント3 |
| `msoThemeColorAccent4` | 8 | アクセント4 |
| `msoThemeColorAccent5` | 9 | アクセント5 |
| `msoThemeColorAccent6` | 10 | アクセント6 |
| `msoThemeColorHyperlink` | 11 | ハイパーリンク |
| `msoThemeColorFollowedHyperlink` | 12 | 表示済みハイパーリンク |

### 1.9 SlideMaster.Background

**パス**: `Master.Background`
**戻り値**: `ShapeRange` オブジェクト
**説明**: マスターの背景を返す。Fill プロパティで塗りつぶし設定を変更できる。

```python
master = prs.SlideMaster
bg = master.Background

# 単色塗りつぶし
bg.Fill.Solid()
bg.Fill.ForeColor.RGB = 0xFFFFFF  # 白

# グラデーション
# msoGradientHorizontal = 1
bg.Fill.PresetGradient(1, 1, 20)

# 画像で背景を設定
# bg.Fill.UserPicture(r"C:\path\to\image.jpg")

# 背景のパターン
# bg.Fill.Patterned(msoPatternDarkHorizontal)
```

### 1.10 SlideMaster.HeadersFooters

**パス**: `Master.HeadersFooters`
**戻り値**: `HeadersFooters` オブジェクト
**説明**: マスターのヘッダー・フッター設定。マスターで設定した内容は、個別設定がないスライドに継承される。

```python
master = prs.SlideMaster
hf = master.HeadersFooters

# フッターの設定
hf.Footer.Visible = True  # -1 = msoTrue
hf.Footer.Text = "社外秘"

# スライド番号の表示
hf.SlideNumber.Visible = True  # -1 = msoTrue

# 日時の設定
hf.DateAndTime.Visible = True  # -1 = msoTrue
hf.DateAndTime.UseFormat = True  # -1 = msoTrue（自動更新）
# ppDateTimeMdyy = 1
hf.DateAndTime.Format = 1
```

### 1.11 複数マスターの管理

PowerPoint では1つのプレゼンテーション内に複数のスライドマスター（デザイン）を持つことができる。これは `Designs` コレクションを通じて管理する。

```python
# 複数マスターの完全な管理例

# 全デザイン（マスター）の列挙
for i in range(1, prs.Designs.Count + 1):
    design = prs.Designs(i)
    master = design.SlideMaster
    print(f"Design {i}: '{design.Name}'")
    print(f"  Master: '{master.Name}'")
    print(f"  Layouts: {master.CustomLayouts.Count}")

# 特定のスライドがどのデザイン（マスター）を使っているか
slide = prs.Slides(1)
slide_design = slide.Design
print(f"Slide 1 uses design: {slide_design.Name}")

# スライドに別のデザインを割り当て
# slide.Design = prs.Designs(2)

# 各スライドのマスターを取得
slide_master = slide.Master  # Slide.Master プロパティ
print(f"Slide 1 master: {slide_master.Name}")
```

**注意点**:
- `Presentation.SlideMaster` は最初のデザインのマスターのみ返す
- 複数マスターを扱うには必ず `Presentation.Designs` を使う
- `Slide.Design` でスライドごとのデザインを取得・設定可能
- `Slide.Master` でスライドに適用されているマスターを直接取得可能
- デザインの最小数は1（最後のデザインは削除できない）

### 1.12 SlideMaster.TextStyles

**パス**: `Master.TextStyles`
**戻り値**: `TextStyles` コレクション
**説明**: マスターの3種類のテキストスタイル（タイトル、本文、デフォルト）を管理する。これがプレースホルダーの書式継承の根幹となる。

```python
master = prs.SlideMaster
styles = master.TextStyles

# ppTitleStyle = 1, ppBodyStyle = 2, ppDefaultStyle = 3
# タイトルスタイルの設定
title_style = styles(1)  # ppTitleStyle
# 本文スタイルの設定
body_style = styles(2)   # ppBodyStyle
# デフォルトスタイルの設定
default_style = styles(3)  # ppDefaultStyle

# 各スタイルにはレベル（1〜5）がある
# タイトルスタイルの第1レベルのフォント設定
level1 = title_style.Levels(1)
level1.Font.Name = "Yu Gothic UI"
level1.Font.Size = 36
level1.Font.Bold = True  # -1 = msoTrue

# 本文スタイルの各レベル設定
for lvl in range(1, 6):
    body_level = body_style.Levels(lvl)
    body_level.Font.Size = 28 - (lvl - 1) * 2  # レベルごとにサイズ減少
    print(f"  Body Level {lvl}: Size={body_level.Font.Size}")
```

**TextStyle の構造**:
- `TextStyle.TextFrame` - テキストフレーム（テキスト配置の設定）
- `TextStyle.Levels` - `TextStyleLevels` コレクション（レベル1〜5）
  - `TextStyleLevel.Font` - フォント設定
  - `TextStyleLevel.ParagraphFormat` - 段落書式

---

## 2. CustomLayout（カスタムレイアウト）

### 2.1 概要

CustomLayout オブジェクトは、スライドマスターに関連付けられた個々のレイアウトを表す。各レイアウトは、スライドで使用可能なプレースホルダーの種類と配置を定義する。

### 2.2 CustomLayouts コレクション

**パス**: `Master.CustomLayouts`
**戻り値**: `CustomLayouts` コレクション

#### メソッド

| メソッド | 説明 | 構文 |
|---------|------|------|
| `Add(Index)` | 新しいカスタムレイアウトを追加 | `CustomLayouts.Add(Index)` |
| `Item(Index)` | インデックスでレイアウトを取得 | `CustomLayouts.Item(Index)` |
| `Paste([Index])` | クリップボードからレイアウトを貼り付け | `CustomLayouts.Paste([Index])` |

#### プロパティ

| プロパティ | 型 | 説明 |
|-----------|-----|------|
| `Count` | Long | レイアウト数 |
| `Application` | Application | Application オブジェクト |
| `Parent` | Object | 親オブジェクト |

```python
master = prs.SlideMaster
layouts = master.CustomLayouts

# レイアウトの数を取得
print(f"レイアウト数: {layouts.Count}")

# 新しい空のレイアウトを先頭に追加
new_layout = layouts.Add(1)
new_layout.Name = "カスタムレイアウト"
```

### 2.3 CustomLayout オブジェクトのプロパティ

| プロパティ | 型 | 説明 |
|-----------|-----|------|
| `Name` | String | レイアウト名（読み書き可能） |
| `Index` | Long | コレクション内のインデックス |
| `MatchingName` | String | マッチング名（テーマ間の対応付けに使用） |
| `Shapes` | Shapes | レイアウト上のシェイプコレクション |
| `Background` | ShapeRange | 背景 |
| `Design` | Design | 親デザイン |
| `DisplayMasterShapes` | MsoTriState | マスターシェイプを表示するか |
| `FollowMasterBackground` | MsoTriState | マスター背景に従うか |
| `Guides` | Guides | ガイド線 |
| `HeadersFooters` | HeadersFooters | ヘッダー・フッター |
| `Height` | Single | 高さ（ポイント） |
| `Width` | Single | 幅（ポイント） |
| `Hyperlinks` | Hyperlinks | ハイパーリンク |
| `Preserved` | MsoTriState | レイアウトを保持するか（使用されていなくても削除しない） |
| `SlideShowTransition` | SlideShowTransition | 切り替え効果 |
| `ThemeColorScheme` | ThemeColorScheme | テーマカラースキーム |
| `TimeLine` | TimeLine | タイムライン |
| `CustomerData` | CustomerData | カスタムデータ |

### 2.4 CustomLayout オブジェクトのメソッド

| メソッド | 説明 |
|---------|------|
| `Copy` | レイアウトをクリップボードにコピー |
| `Cut` | レイアウトを切り取り |
| `Delete` | レイアウトを削除 |
| `Duplicate` | レイアウトを複製 |
| `MoveTo(toPos)` | レイアウトの位置を移動 |
| `Select` | レイアウトを選択（UI上） |

```python
master = prs.SlideMaster

# レイアウトの複製
original_layout = master.CustomLayouts(1)
duplicated = original_layout.Duplicate()
duplicated.Name = "コピーされたレイアウト"

# レイアウトの削除
# 注意: 使用中のレイアウトは削除できない場合がある
# master.CustomLayouts(3).Delete()

# レイアウトの移動（位置を変更）
# master.CustomLayouts(3).MoveTo(1)
```

### 2.5 CustomLayout.Name（レイアウト名の取得・設定）

```python
master = prs.SlideMaster

# 全レイアウト名を取得
layout_names = []
for i in range(1, master.CustomLayouts.Count + 1):
    layout = master.CustomLayouts(i)
    layout_names.append(layout.Name)
    print(f"  {i}: '{layout.Name}' (MatchingName: '{layout.MatchingName}')")

# レイアウト名を変更
master.CustomLayouts(1).Name = "新しいレイアウト名"
```

**標準レイアウト名の例**（テーマにより異なる）:
- タイトル スライド
- タイトルとコンテンツ
- セクション見出し
- 2つのコンテンツ
- 比較
- タイトルのみ
- 白紙
- タイトル付きのコンテンツ
- タイトル付きの図

### 2.6 CustomLayout.Placeholders

レイアウト上のプレースホルダーは `CustomLayout.Shapes.Placeholders` で取得する。

```python
master = prs.SlideMaster

# 「タイトルとコンテンツ」レイアウトのプレースホルダーを調べる
for i in range(1, master.CustomLayouts.Count + 1):
    layout = master.CustomLayouts(i)
    if layout.Name == "タイトルとコンテンツ":
        phs = layout.Shapes.Placeholders
        print(f"レイアウト '{layout.Name}' のプレースホルダー数: {phs.Count}")
        for j in range(1, phs.Count + 1):
            ph = phs(j)
            pf = ph.PlaceholderFormat
            print(f"  [{j}] Type={pf.Type}, Name='{ph.Name}', "
                  f"Left={ph.Left:.1f}, Top={ph.Top:.1f}, "
                  f"Width={ph.Width:.1f}, Height={ph.Height:.1f}")
        break
```

### 2.7 スライドへのレイアウト適用 (Slide.CustomLayout)

**パス**: `Slide.CustomLayout`
**説明**: スライドに適用されているカスタムレイアウトを取得・設定する。

```python
# 現在のスライドのレイアウトを確認
slide = prs.Slides(1)
current_layout = slide.CustomLayout
print(f"スライド1のレイアウト: '{current_layout.Name}'")

# レイアウトを変更する
# 方法1: マスターのレイアウトコレクションから指定
target_layout = prs.SlideMaster.CustomLayouts(2)  # 2番目のレイアウト
slide.CustomLayout = target_layout

# 方法2: レイアウト名で検索して適用
for i in range(1, prs.SlideMaster.CustomLayouts.Count + 1):
    layout = prs.SlideMaster.CustomLayouts(i)
    if layout.Name == "タイトルとコンテンツ":
        slide.CustomLayout = layout
        break
```

**注意点**:
- `Slide.CustomLayout` は読み書き可能（ドキュメントによっては読み取り専用と記載されるが、実際には設定可能）
- レイアウト変更時、既存のコンテンツは可能な限り保持される
- 異なるマスターのレイアウトを適用すると、スライドのデザインも変更される

### 2.8 レイアウトの一覧取得方法（実用的な関数）

```python
def get_all_layouts(presentation):
    """プレゼンテーション内の全デザインの全レイアウトを取得する"""
    result = []
    for d_idx in range(1, presentation.Designs.Count + 1):
        design = presentation.Designs(d_idx)
        master = design.SlideMaster
        for l_idx in range(1, master.CustomLayouts.Count + 1):
            layout = master.CustomLayouts(l_idx)
            ph_info = []
            phs = layout.Shapes.Placeholders
            for p_idx in range(1, phs.Count + 1):
                ph = phs(p_idx)
                ph_info.append({
                    "index": p_idx,
                    "type": ph.PlaceholderFormat.Type,
                    "name": ph.Name,
                })
            result.append({
                "design_index": d_idx,
                "design_name": design.Name,
                "layout_index": l_idx,
                "layout_name": layout.Name,
                "placeholders": ph_info,
            })
    return result

# 使用例
layouts = get_all_layouts(prs)
for layout in layouts:
    print(f"[Design '{layout['design_name']}'] Layout '{layout['layout_name']}'")
    for ph in layout["placeholders"]:
        print(f"  Placeholder: type={ph['type']}, name='{ph['name']}'")
```

### 2.9 レイアウトにプレースホルダーを追加

**重要**: COM API では、レイアウトにプレースホルダーを新規追加する直接的な方法は限られている。`Shapes.AddPlaceholder` メソッドは、**削除されたプレースホルダーの復元**に使用されるものであり、新規作成ではない。

レイアウトへのプレースホルダー追加は、主にマスター/レイアウトの `Shapes.AddPlaceholder` を通じて行う。

```python
master = prs.SlideMaster

# マスターに削除されたプレースホルダーを復元
# ppPlaceholderTitle = 1
# restored = master.Shapes.AddPlaceholder(Type=1)

# レイアウトに削除されたプレースホルダーを復元
layout = master.CustomLayouts(1)
# ppPlaceholderBody = 2
# restored = layout.Shapes.AddPlaceholder(Type=2, Left=100, Top=200, Width=400, Height=300)
```

**制限事項**:
- `AddPlaceholder` は削除済みプレースホルダーの復元のみ可能
- スライドが作成時に持っていた数を超えてプレースホルダーを追加することはできない
- プレースホルダーの数を変更するには `Slide.Layout` プロパティを変更する

---

## 3. Placeholders（プレースホルダー） -- 最重要

### 3.1 概要

プレースホルダーは PowerPoint の最も重要な概念の1つである。スライドマスター → レイアウト → スライド の継承チェーンの中核を担い、テキスト、画像、表、グラフ、SmartArt など、さまざまなコンテンツを配置するための「枠」として機能する。

### 3.2 Placeholders コレクション

**パス**: `Slide.Shapes.Placeholders` / `CustomLayout.Shapes.Placeholders`
**戻り値**: `Placeholders` コレクション
**説明**: スライドまたはレイアウト上の全プレースホルダーのコレクション。各メンバーは `Shape` オブジェクト。

```python
slide = prs.Slides(1)
placeholders = slide.Shapes.Placeholders

print(f"プレースホルダー数: {placeholders.Count}")

# 全プレースホルダーを列挙
for i in range(1, placeholders.Count + 1):
    ph = placeholders(i)
    pf = ph.PlaceholderFormat
    print(f"Placeholder {i}:")
    print(f"  Name: '{ph.Name}'")
    print(f"  Type: {pf.Type}")
    print(f"  ContainedType: {pf.ContainedType}")
    print(f"  Position: Left={ph.Left:.1f}, Top={ph.Top:.1f}")
    print(f"  Size: Width={ph.Width:.1f}, Height={ph.Height:.1f}")
    print(f"  HasTextFrame: {ph.HasTextFrame}")
```

#### Placeholders コレクションのメソッド

| メソッド | 説明 | 構文 |
|---------|------|------|
| `Item(Index)` | インデックス（Long）でプレースホルダーを取得 | `Placeholders.Item(index)` |
| `FindByName(Index)` | インデックスまたは名前で検索 | `Placeholders.FindByName(nameOrIndex)` |

#### プロパティ

| プロパティ | 型 | 説明 |
|-----------|-----|------|
| `Count` | Long | プレースホルダー数 |
| `Application` | Application | Application オブジェクト |
| `Parent` | Object | 親オブジェクト |

### 3.3 Placeholders(index) でのアクセス方法

**重要**: Placeholders コレクションのインデックスは、シェイプの追加順序に基づく**プレースホルダー固有のインデックス番号**であり、Shapes コレクションのインデックスとは異なる。

```python
slide = prs.Slides(1)
phs = slide.Shapes.Placeholders

# インデックスでアクセス（1から始まる）
# タイトルがあるスライドでは、Placeholders(1) がタイトル
title_ph = phs(1)

# Shapes.Title と Placeholders(1) は同等
# assert slide.Shapes.Title.Name == phs(1).Name

# コンテンツプレースホルダーは通常 Placeholders(2)
if phs.Count >= 2:
    content_ph = phs(2)
```

#### FindByName を使ったアクセス

`FindByName` メソッドは `Item` メソッドと異なり、**名前（文字列）** でもアクセスできる。

```python
# 名前でプレースホルダーを検索
title_ph = phs.FindByName("Title 1")

# インデックスでも使用可能（Variantを受け取る）
ph = phs.FindByName(1)
```

**注意点**:
- `Item(Index)` は `Long` 型のみ受け付ける（名前指定不可）
- `FindByName(Index)` は `Variant` 型を受け付ける（名前でもインデックスでも可能）
- プレースホルダーが削除された場合、インデックスに欠番が生じる可能性がある
- インデックスはレイアウトで定義された順序に基づく

### 3.4 PlaceholderFormat オブジェクト

**パス**: `Shape.PlaceholderFormat`
**戻り値**: `PlaceholderFormat` オブジェクト
**説明**: プレースホルダー固有のプロパティを含むオブジェクト。Shape.Type が `msoPlaceholder (14)` の場合のみアクセス可能。

| プロパティ | 型 | 説明 |
|-----------|-----|------|
| `Type` | PpPlaceholderType | プレースホルダーの種類 |
| `ContainedType` | MsoShapeType | 含まれているコンテンツの種類 |
| `Name` | String | プレースホルダー名 |
| `Application` | Application | Application オブジェクト |
| `Parent` | Object | 親オブジェクト |

### 3.5 PpPlaceholderType の全種類（完全版）

| 定数名 | 値 | 説明 | 用途 |
|--------|-----|------|------|
| `ppPlaceholderMixed` | -2 | 混合（複数選択時） | 複数プレースホルダー選択時に返される |
| `ppPlaceholderTitle` | 1 | タイトル | スライドのメインタイトル |
| `ppPlaceholderBody` | 2 | 本文/コンテンツ | テキスト本文、コンテンツエリア |
| `ppPlaceholderCenterTitle` | 3 | 中央タイトル | タイトルスライドの中央配置タイトル |
| `ppPlaceholderSubtitle` | 4 | サブタイトル | タイトルスライドのサブタイトル |
| `ppPlaceholderVerticalTitle` | 5 | 縦書きタイトル | 縦書きレイアウトのタイトル |
| `ppPlaceholderVerticalBody` | 6 | 縦書き本文 | 縦書きレイアウトの本文 |
| `ppPlaceholderObject` | 7 | オブジェクト | OLEオブジェクト用 |
| `ppPlaceholderChart` | 8 | グラフ | グラフ専用 |
| `ppPlaceholderBitmap` | 9 | ビットマップ | ビットマップ画像 |
| `ppPlaceholderMediaClip` | 10 | メディアクリップ | 動画・音声 |
| `ppPlaceholderOrgChart` | 11 | 組織図 | 組織図 |
| `ppPlaceholderTable` | 12 | 表 | 表専用 |
| `ppPlaceholderSlideNumber` | 13 | スライド番号 | スライド番号表示 |
| `ppPlaceholderHeader` | 14 | ヘッダー | ヘッダー（ノート・配布資料のみ） |
| `ppPlaceholderFooter` | 15 | フッター | フッターテキスト |
| `ppPlaceholderDate` | 16 | 日付 | 日付表示 |
| `ppPlaceholderVerticalObject` | 17 | 縦書きオブジェクト | 縦書きオブジェクト |
| `ppPlaceholderPicture` | 18 | 画像 | 画像専用プレースホルダー |
| `ppPlaceholderCameo` | 19 | カメオ | カメオ（PowerPoint 365のライブカメラ機能） |

### 3.6 PlaceholderFormat.ContainedType

**パス**: `Shape.PlaceholderFormat.ContainedType`
**戻り値**: `MsoShapeType`
**説明**: プレースホルダー内に実際に含まれているコンテンツの種類を返す。

```python
slide = prs.Slides(1)
for i in range(1, slide.Shapes.Placeholders.Count + 1):
    ph = slide.Shapes.Placeholders(i)
    pf = ph.PlaceholderFormat
    print(f"Placeholder {i}: Type={pf.Type}, ContainedType={pf.ContainedType}")
    # ContainedType の例:
    # msoAutoShape = 1       (空のプレースホルダー/テキスト)
    # msoTable = 19          (表が挿入されている)
    # msoChart = 3           (グラフが挿入されている)
    # msoSmartArt = 24       (SmartArt が挿入されている)
    # msoPicture = 13        (画像が挿入されている)
    # msoLinkedPicture = 11  (リンク画像)
    # msoPlaceholder = 14    (空のプレースホルダー)
```

**ContainedType で返される主な MsoShapeType 値**:

| 定数名 | 値 | 説明 |
|--------|-----|------|
| `msoAutoShape` | 1 | オートシェイプ（テキスト含む） |
| `msoChart` | 3 | グラフ |
| `msoLinkedPicture` | 11 | リンク画像 |
| `msoPicture` | 13 | 画像 |
| `msoPlaceholder` | 14 | プレースホルダー（空） |
| `msoTable` | 19 | 表 |
| `msoSmartArt` | 24 | SmartArt |
| `msoMedia` | 16 | メディア |

### 3.7 プレースホルダーのテキスト操作

#### 基本的なテキスト操作

```python
slide = prs.Slides(1)
phs = slide.Shapes.Placeholders

# タイトルにテキストを設定
title = phs(1)
if title.HasTextFrame:
    title.TextFrame.TextRange.Text = "プレゼンテーションタイトル"

# コンテンツ（本文）にテキストを設定
if phs.Count >= 2:
    body = phs(2)
    if body.HasTextFrame:
        body.TextFrame.TextRange.Text = "箇条書き1\r箇条書き2\r箇条書き3"
```

#### HasTextFrame と TextFrame.TextRange の詳細

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(1)

# テキストフレームの存在確認
# msoTrue = -1, msoFalse = 0
if ph.HasTextFrame:
    tf = ph.TextFrame
    tr = tf.TextRange

    # テキスト全体を取得
    full_text = tr.Text
    print(f"テキスト: {full_text}")

    # テキストがあるか確認
    has_text = tf.HasText  # msoTrue (-1) or msoFalse (0)

    # テキストを設定
    tr.Text = "新しいテキスト"

    # 段落（パラグラフ）操作
    paragraphs = tr.Paragraphs()
    print(f"段落数: {paragraphs.Count}")

    # 特定の段落にアクセス
    para1 = tr.Paragraphs(1)
    print(f"第1段落: {para1.Text}")

    # 段落範囲を指定
    # Paragraphs(Start, Length) - Start段落目からLength段落分
    paras = tr.Paragraphs(2, 1)  # 2番目の段落のみ

    # 文字範囲を指定
    # Characters(Start, Length) - Start文字目からLength文字分
    chars = tr.Characters(1, 5)  # 最初の5文字

    # ラン（書式が同一の連続部分）
    runs = tr.Runs()
    for r_idx in range(1, runs.Count + 1):
        run = runs(r_idx)
        print(f"  Run {r_idx}: '{run.Text}', Font.Size={run.Font.Size}")
```

#### テキストの書式設定

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(1)
tr = ph.TextFrame.TextRange

# フォント設定
tr.Font.Name = "Yu Gothic UI"
tr.Font.Size = 28
tr.Font.Bold = True       # msoTrue = -1
tr.Font.Italic = False     # msoFalse = 0
tr.Font.Underline = False
tr.Font.Color.RGB = 0x000000  # 黒（BGR形式: 0xBBGGRR）

# テーマカラーを使用
# msoThemeColorAccent1 = 5
tr.Font.Color.ObjectThemeColor = 5

# 段落書式
para_fmt = tr.ParagraphFormat
# ppAlignLeft=1, ppAlignCenter=2, ppAlignRight=3, ppAlignJustify=4
para_fmt.Alignment = 2  # 中央揃え

# インデントレベル（0〜4）
para_fmt.IndentLevel = 1

# 行間
para_fmt.SpaceAfter = 6   # 段落後の間隔（ポイント）
para_fmt.SpaceBefore = 0  # 段落前の間隔
para_fmt.SpaceWithin = 1.2  # 行間（倍率）

# 箇条書きの設定
para_fmt.Bullet.Type = 1  # ppBulletUnnumbered
para_fmt.Bullet.Character = 8226  # Unicode の黒丸
para_fmt.Bullet.Font.Color.RGB = 0xFF0000  # 赤色の箇条書き記号

# 特定の段落のみ書式変更
para2 = tr.Paragraphs(2)
para2.Font.Size = 20
para2.Font.Color.RGB = 0x808080  # グレー
```

#### テキストの挿入と操作

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(2)
tr = ph.TextFrame.TextRange

# テキスト全体を置換
tr.Text = "新しい内容"

# テキストの末尾に追加（改行込み）
tr.InsertAfter("\r追加テキスト")

# テキストの先頭に挿入
tr.InsertBefore("先頭テキスト\r")

# 日時の挿入
# ppDateTimeMdyy = 1
# tr.InsertDateTime(DateTimeFormat=1, InsertAsField=True)

# スライド番号の挿入
# tr.InsertSlideNumber()
```

### 3.8 プレースホルダー内のコンテンツ種類

プレースホルダーは以下の種類のコンテンツを保持できる:

#### 3.8.1 テキスト

前述のとおり `TextFrame.TextRange` で操作。

#### 3.8.2 画像

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(2)  # コンテンツプレースホルダー

# 方法1: Fill.UserPicture を使用（プレースホルダーの塗りつぶしとして画像を設定）
# 注意: この方法は画像がストレッチされ、アスペクト比が保持されない
# ph.Fill.UserPicture(r"C:\path\to\image.jpg")

# 方法2: 画像プレースホルダー（ppPlaceholderPicture = 18）の場合
# Shape の位置・サイズを取得して AddPicture で配置
# msoTrue = -1, msoFalse = 0
left = ph.Left
top = ph.Top
width = ph.Width
height = ph.Height

# プレースホルダーを削除して同じ位置に画像を配置
# ph.Delete()
# new_pic = slide.Shapes.AddPicture(
#     FileName=r"C:\path\to\image.jpg",
#     LinkToFile=0,      # msoFalse
#     SaveWithDocument=-1,  # msoTrue
#     Left=left,
#     Top=top,
#     Width=width,
#     Height=height
# )
```

**画像プレースホルダーの制限事項**:
- COM API では、プレースホルダーに直接画像を「挿入」する（UI のように）メソッドがない
- `Fill.UserPicture` は背景塗りつぶしとして画像を設定するため、アスペクト比が保持されない
- 実用的なアプローチは、プレースホルダーの位置・サイズを取得し、`Shapes.AddPicture` で同じ位置に配置する方法

#### 3.8.3 表

```python
slide = prs.Slides(1)

# プレースホルダーの位置に表を追加
ph = slide.Shapes.Placeholders(2)
left, top, width, height = ph.Left, ph.Top, ph.Width, ph.Height

# 表の追加（プレースホルダーを置き換える形で）
table_shape = slide.Shapes.AddTable(
    NumRows=4,
    NumColumns=3,
    Left=left,
    Top=top,
    Width=width,
    Height=height
)

# 表のセルにデータを設定
table = table_shape.Table
table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "ヘッダー1"
table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "ヘッダー2"
table.Cell(1, 3).Shape.TextFrame.TextRange.Text = "ヘッダー3"
table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "データ1"
```

#### 3.8.4 グラフ

```python
slide = prs.Slides(1)

# グラフの追加
# AddChart2(Style, Type, Left, Top, Width, Height)
# xlColumnClustered = 51
ph = slide.Shapes.Placeholders(2)
chart_shape = slide.Shapes.AddChart2(
    Style=201,
    Type=51,  # xlColumnClustered
    Left=ph.Left,
    Top=ph.Top,
    Width=ph.Width,
    Height=ph.Height
)

# グラフにデータを設定（Excel ワークシート経由）
chart = chart_shape.Chart
chart_data = chart.ChartData
chart_data.Activate()
# ワークシートにデータを書き込み...
```

#### 3.8.5 SmartArt

```python
slide = prs.Slides(1)

# SmartArt の追加
# SmartArt レイアウトの取得
# smartart_layout = app.SmartArtLayouts(1)  # SmartArt レイアウトを指定

# slide.Shapes.AddSmartArt(
#     Layout=smartart_layout,
#     Left=ph.Left,
#     Top=ph.Top,
#     Width=ph.Width,
#     Height=ph.Height
# )
```

### 3.9 プレースホルダーの書式設定（マスターからの継承と上書き）

プレースホルダーの書式は、マスター → レイアウト → スライド の順に継承される。スライドレベルで書式を変更すると「上書き」となり、以降マスター/レイアウトの変更が反映されなくなる。

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(1)

# 書式を上書き（マスターの継承が切れる）
ph.TextFrame.TextRange.Font.Size = 40
ph.TextFrame.TextRange.Font.Color.RGB = 0xFF0000  # 赤

# マスター側の書式を変更（全スライドに影響）
master = prs.SlideMaster
# タイトルスタイルのフォントサイズを変更
master.TextStyles(1).Levels(1).Font.Size = 44  # ppTitleStyle

# 注意: スライドレベルで上書きされた書式は、
# マスター変更の影響を受けない
```

### 3.10 プレースホルダーの位置・サイズ変更

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(1)

# 位置を変更（ポイント単位）
ph.Left = 72    # 1インチ = 72ポイント
ph.Top = 36

# サイズを変更
ph.Width = 648  # 9インチ
ph.Height = 72  # 1インチ

# ポイント ↔ インチ変換
# 1 inch = 72 points
# 1 cm = 28.35 points

# ロック（位置固定）
ph.LockAspectRatio = True   # msoTrue = -1（アスペクト比固定）
ph.LockAspectRatio = False  # msoFalse = 0

# 回転
ph.Rotation = 0  # 度単位
```

### 3.11 プレースホルダーの判定と検索

プレースホルダーかどうかの判定、および特定タイプのプレースホルダーを検索するユーティリティ関数:

```python
def is_placeholder(shape):
    """シェイプがプレースホルダーかどうか判定"""
    # msoPlaceholder = 14
    return shape.Type == 14


def find_placeholder_by_type(slide, placeholder_type):
    """特定のタイプのプレースホルダーを検索して返す

    Args:
        slide: Slide オブジェクト
        placeholder_type: PpPlaceholderType の値
            1=Title, 2=Body, 3=CenterTitle, 4=Subtitle,
            13=SlideNumber, 14=Header, 15=Footer, 16=Date,
            18=Picture

    Returns:
        Shape オブジェクト（見つからない場合は None）
    """
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        if ph.PlaceholderFormat.Type == placeholder_type:
            return ph
    return None


def find_all_placeholders_by_type(slide, placeholder_type):
    """特定のタイプの全プレースホルダーをリストで返す"""
    result = []
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        if ph.PlaceholderFormat.Type == placeholder_type:
            result.append(ph)
    return result


# 使用例
slide = prs.Slides(1)

# タイトルプレースホルダーを取得
title_ph = find_placeholder_by_type(slide, 1)  # ppPlaceholderTitle
if title_ph:
    title_ph.TextFrame.TextRange.Text = "タイトル"

# 本文プレースホルダーを取得
body_ph = find_placeholder_by_type(slide, 2)  # ppPlaceholderBody
if body_ph:
    body_ph.TextFrame.TextRange.Text = "本文テキスト"

# フッタープレースホルダーを取得
footer_ph = find_placeholder_by_type(slide, 15)  # ppPlaceholderFooter
```

---

## 4. 継承構造 (Master → Layout → Slide)

### 4.1 継承の全体図

PowerPoint の書式継承は以下の階層で行われる:

```
Theme (テーマ)
  ├── フォントテーマ (Heading / Body)
  ├── カラーテーマ (12色)
  └── エフェクトテーマ
      │
      ▼
SlideMaster (スライドマスター)
  ├── Background (背景)
  ├── TextStyles (テキストスタイル: Title/Body/Default × 5レベル)
  ├── Shapes (マスター上のシェイプ: ロゴ、装飾等)
  ├── HeadersFooters (ヘッダー・フッター設定)
  └── Placeholders (マスタープレースホルダー)
      │
      ▼
CustomLayout (カスタムレイアウト)
  ├── Background (独自 or マスターに従う: FollowMasterBackground)
  ├── Shapes (レイアウト固有のシェイプ)
  ├── DisplayMasterShapes (マスターシェイプ表示: True/False)
  ├── HeadersFooters (レイアウトレベル設定)
  └── Placeholders (レイアウトプレースホルダー - 種類・位置を定義)
      │
      ▼
Slide (スライド)
  ├── Background (独自 or レイアウト/マスターに従う: FollowMasterBackground)
  ├── DisplayMasterShapes (マスターシェイプ表示: True/False)
  ├── HeadersFooters (スライドレベル設定)
  └── Placeholders (スライドプレースホルダー - 実際のコンテンツ)
```

### 4.2 マスター上のプレースホルダーがレイアウト・スライドに継承される仕組み

1. **マスターのプレースホルダー**: マスターには基本的なプレースホルダー（タイトル、本文、日付、フッター、スライド番号）が定義される
2. **レイアウトへの継承**: 各レイアウトはマスターのプレースホルダーを継承し、さらに独自のプレースホルダーを追加できる。レイアウトでプレースホルダーの位置・サイズ・書式を変更すると、そのレイアウトレベルで「上書き」される
3. **スライドへの継承**: スライドを作成すると、使用するレイアウトのプレースホルダーがスライドに複製される。スライドレベルでの変更は「上書き」となる

```python
# 継承の確認例
master = prs.SlideMaster

# マスターのタイトルスタイルを確認
master_title_font = master.TextStyles(1).Levels(1).Font  # ppTitleStyle
print(f"マスター タイトルフォント: {master_title_font.Name}, Size={master_title_font.Size}")

# マスターの本文スタイルを確認
master_body_font = master.TextStyles(2).Levels(1).Font  # ppBodyStyle
print(f"マスター 本文フォント: {master_body_font.Name}, Size={master_body_font.Size}")

# スライドのプレースホルダーが継承しているか確認
slide = prs.Slides(1)
if slide.Shapes.Placeholders.Count > 0:
    title_ph = slide.Shapes.Placeholders(1)
    if title_ph.HasTextFrame:
        slide_font = title_ph.TextFrame.TextRange.Font
        print(f"スライド タイトルフォント: {slide_font.Name}, Size={slide_font.Size}")
```

### 4.3 書式の継承と上書きの仕組み

#### 継承のルール

1. **テーマフォント**: テーマで定義されたフォント（Heading/Body）はマスター→レイアウト→スライドに自動継承される。フォント名に「(本文)」や「(見出し)」が付いている場合、テーマフォントが使用されている
2. **書式の継承**: マスターのプレースホルダーで設定した書式（フォントサイズ、色、箇条書き等）は、レイアウト→スライドに継承される
3. **上書きの発生**: レイアウトまたはスライドレベルで書式を変更すると、その特定の属性について「上書き」が発生し、上位レベルの変更が反映されなくなる

#### 上書きの確認と解除

```python
slide = prs.Slides(1)

# スライドの背景がマスターに従っているか確認
print(f"FollowMasterBackground: {slide.FollowMasterBackground}")
# msoTrue = -1: マスター背景に従う
# msoFalse = 0: 独自の背景

# マスターシェイプが表示されているか確認
print(f"DisplayMasterShapes: {slide.DisplayMasterShapes}")

# マスター背景への追従を有効化
slide.FollowMasterBackground = True  # msoTrue = -1

# マスターシェイプの表示を有効化
slide.DisplayMasterShapes = True  # msoTrue = -1

# ヘッダー・フッターの個別設定をクリア（マスター設定に戻す）
slide.HeadersFooters.Clear()
```

### 4.4 テーマフォント・テーマカラーの継承

```python
master = prs.SlideMaster
theme = master.Theme

# テーマカラースキームの確認
tcs = theme.ThemeColorScheme
print("テーマカラー:")
color_names = [
    "Dark1", "Light1", "Dark2", "Light2",
    "Accent1", "Accent2", "Accent3", "Accent4",
    "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"
]
for i in range(1, min(tcs.Count + 1, 13)):
    print(f"  {color_names[i-1]}: RGB={tcs(i).RGB}")

# テーマカラーの変更（全スライドに影響）
# tcs(5).RGB = 0xFF3300  # Accent1 を変更

# テーマフォントの確認（ThemeFontScheme 経由）
# 注意: COM経由でのテーマフォントの直接取得は制限がある場合がある
# 代替手段として TextStyles から確認
title_font = master.TextStyles(1).Levels(1).Font  # ppTitleStyle
body_font = master.TextStyles(2).Levels(1).Font   # ppBodyStyle
print(f"テーマ見出しフォント: {title_font.Name}")
print(f"テーマ本文フォント: {body_font.Name}")
```

### 4.5 Background の継承

```python
# === マスターの背景設定 ===
master = prs.SlideMaster
master_bg = master.Background.Fill
master_bg.Solid()
master_bg.ForeColor.RGB = 0xF0F0F0  # 薄いグレー

# === レイアウトの背景設定 ===
layout = master.CustomLayouts(1)

# レイアウトがマスター背景に従うか確認
print(f"Layout FollowMasterBackground: {layout.FollowMasterBackground}")

# レイアウト独自の背景を設定する場合
layout.FollowMasterBackground = False  # msoFalse = 0
layout.Background.Fill.Solid()
layout.Background.Fill.ForeColor.RGB = 0xFFFFFF  # 白

# マスター背景に戻す
layout.FollowMasterBackground = True  # msoTrue = -1

# === スライドの背景設定 ===
slide = prs.Slides(1)

# スライドがマスター背景に従うか確認
print(f"Slide FollowMasterBackground: {slide.FollowMasterBackground}")

# スライド独自の背景を設定
slide.FollowMasterBackground = False  # msoFalse = 0
slide.Background.Fill.Solid()
slide.Background.Fill.ForeColor.RGB = 0x003366  # ダークブルー

# マスター背景に戻す
slide.FollowMasterBackground = True  # msoTrue = -1
```

### 4.6 DisplayMasterShapes の制御

マスターに配置されたロゴや装飾などの「マスターシェイプ」の表示・非表示を制御する。

```python
# レイアウトレベル
layout = prs.SlideMaster.CustomLayouts(1)
layout.DisplayMasterShapes = True   # マスターシェイプを表示
# layout.DisplayMasterShapes = False  # マスターシェイプを非表示

# スライドレベル
slide = prs.Slides(1)
slide.DisplayMasterShapes = True   # マスターシェイプを表示
# slide.DisplayMasterShapes = False  # マスターシェイプを非表示
```

---

## 5. HeadersFooters（ヘッダー・フッター）

### 5.1 概要

HeadersFooters オブジェクトは、スライド、ノートページ、配布資料、マスター上のヘッダー・フッター・日時・スライド番号を管理する。

**パス**: `Slide.HeadersFooters` / `Master.HeadersFooters` / `CustomLayout.HeadersFooters`
**戻り値**: `HeadersFooters` オブジェクト

### 5.2 HeadersFooters オブジェクトのプロパティ

| プロパティ | 戻り値 | 説明 |
|-----------|--------|------|
| `DateAndTime` | HeaderFooter | 日付と時刻のプレースホルダー |
| `Footer` | HeaderFooter | フッターのプレースホルダー |
| `Header` | HeaderFooter | ヘッダーのプレースホルダー（ノート・配布資料のみ） |
| `SlideNumber` | HeaderFooter | スライド番号のプレースホルダー |
| `DisplayOnTitleSlide` | MsoTriState | タイトルスライドに表示するか |
| `Application` | Application | Application オブジェクト |
| `Parent` | Object | 親オブジェクト |

**メソッド**:

| メソッド | 説明 |
|---------|------|
| `Clear` | 個別設定をクリアしてマスター設定に戻す |

### 5.3 HeaderFooter オブジェクトのプロパティ

各 HeaderFooter オブジェクト（Footer, Header, DateAndTime, SlideNumber）は以下のプロパティを持つ:

| プロパティ | 型 | 説明 |
|-----------|-----|------|
| `Text` | String | テキスト内容 |
| `Visible` | MsoTriState | 表示・非表示（True=-1, False=0） |
| `Format` | PpDateTimeFormat | 日付形式（DateAndTime のみ） |
| `UseFormat` | MsoTriState | 自動更新形式を使用するか（DateAndTime のみ） |
| `Application` | Application | Application オブジェクト |
| `Parent` | Object | 親オブジェクト |

### 5.4 HeadersFooters.Footer（フッターテキスト）

```python
# === マスターレベルでフッターを設定 ===
master = prs.SlideMaster
hf = master.HeadersFooters

hf.Footer.Visible = True  # msoTrue = -1
hf.Footer.Text = "社外秘 - Confidential"

# === スライドレベルで個別設定 ===
slide = prs.Slides(1)
slide.HeadersFooters.Footer.Visible = True
slide.HeadersFooters.Footer.Text = "特別会議用"

# === 個別設定をクリア（マスター設定に戻す） ===
slide.HeadersFooters.Clear()
```

### 5.5 HeadersFooters.Header（ヘッダーテキスト）

**重要**: ヘッダーはスライドでは利用できない。ノートマスターと配布資料マスターでのみ使用可能。

```python
# ノートマスターのヘッダーを設定
notes_master = prs.NotesMaster
notes_master.HeadersFooters.Header.Visible = True
notes_master.HeadersFooters.Header.Text = "会議名: 定例会議"

# 配布資料マスターのヘッダーを設定
handout_master = prs.HandoutMaster
handout_master.HeadersFooters.Header.Visible = True
handout_master.HeadersFooters.Header.Text = "配布資料ヘッダー"
```

### 5.6 HeadersFooters.DateAndTime

```python
hf = prs.SlideMaster.HeadersFooters

# 日時を表示
hf.DateAndTime.Visible = True  # msoTrue = -1

# 方法1: 自動更新（プレゼンテーション表示時の日時）
hf.DateAndTime.UseFormat = True  # msoTrue = -1
# ppDateTimeMdyy = 1 (例: 2/16/2026)
hf.DateAndTime.Format = 1

# 方法2: 固定テキスト
hf.DateAndTime.UseFormat = False  # msoFalse = 0
hf.DateAndTime.Text = "2026年2月16日"
```

### 5.7 HeadersFooters.SlideNumber

```python
# スライド番号を表示
master = prs.SlideMaster
master.HeadersFooters.SlideNumber.Visible = True  # msoTrue = -1

# 個別スライドでの制御
slide = prs.Slides(1)
slide.HeadersFooters.SlideNumber.Visible = True
```

### 5.8 DisplayOnTitleSlide

タイトルスライド（最初のスライドや「タイトルスライド」レイアウト使用時）にヘッダー・フッターを表示するかどうかを制御する。

```python
# タイトルスライドでの表示を制御
master = prs.SlideMaster
hf = master.HeadersFooters

# タイトルスライドにフッター等を表示しない
hf.DisplayOnTitleSlide = False  # msoFalse = 0

# タイトルスライドにも表示する
hf.DisplayOnTitleSlide = True  # msoTrue = -1
```

### 5.9 HeadersFooters の継承

```python
# 全スライドのヘッダー・フッターをマスター設定に統一する
master = prs.SlideMaster

# マスターの設定
master.HeadersFooters.Footer.Visible = True
master.HeadersFooters.Footer.Text = "Regional Sales"
master.HeadersFooters.SlideNumber.Visible = True
master.HeadersFooters.DateAndTime.Visible = True
master.HeadersFooters.DateAndTime.UseFormat = True
master.HeadersFooters.DateAndTime.Format = 1  # ppDateTimeMdyy

# 全スライドの個別設定をクリアしてマスターに従わせる
for i in range(1, prs.Slides.Count + 1):
    slide = prs.Slides(i)
    slide.DisplayMasterShapes = True  # マスターシェイプを表示
    slide.HeadersFooters.Clear()     # 個別設定をクリア
```

### 5.10 PpDateTimeFormat 全定数

| 定数名 | 値 | 説明 | 表示例 |
|--------|-----|------|--------|
| `ppDateTimeFormatMixed` | -2 | 混合形式 | - |
| `ppDateTimeMdyy` | 1 | 月/日/年（短） | 2/16/26 |
| `ppDateTimeddddMMMMddyyyy` | 2 | 曜日 月 日 年 | Monday, February 16, 2026 |
| `ppDateTimedMMMMyyyy` | 3 | 日 月 年 | 16 February 2026 |
| `ppDateTimeMMMMdyyyy` | 4 | 月 日 年 | February 16, 2026 |
| `ppDateTimedMMMyy` | 5 | 日 月 年（短） | 16-Feb-26 |
| `ppDateTimeMMMMyy` | 6 | 月 年 | February 26 |
| `ppDateTimeMMyy` | 7 | 月/年 | 2/26 |
| `ppDateTimeMMddyyHmm` | 8 | 月/日/年 時:分 | 2/16/26 14:30 |
| `ppDateTimeMMddyyhmmAMPM` | 9 | 月/日/年 時:分 AM/PM | 2/16/26 2:30 PM |
| `ppDateTimeHmm` | 10 | 時:分（24時間） | 14:30 |
| `ppDateTimeHmmss` | 11 | 時:分:秒（24時間） | 14:30:45 |
| `ppDateTimehmmAMPM` | 12 | 時:分 AM/PM | 2:30 PM |
| `ppDateTimehmmssAMPM` | 13 | 時:分:秒 AM/PM | 2:30:45 PM |
| `ppDateTimeFigureOut` | 14 | 自動判別 | - |
| `ppDateTimeUAQ1` ~ `ppDateTimeUAQ7` | 15-21 | ユーザー定義形式 1-7 | （ロケール依存） |

**注意**: `ppDateTimeUAQ1` 〜 `ppDateTimeUAQ7` はロケール（言語設定）によって表示形式が異なる。日本語環境では和暦や日本式日付形式になる場合がある。

---

## 6. 実践的な操作パターン

### 6.1 特定のレイアウトを持つスライドを追加する

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
prs = app.Presentations.Open(r"C:\path\to\presentation.pptx")

# 方法1: Slides.AddSlide（推奨）
# 特定のレイアウトを指定してスライドを追加
master = prs.SlideMaster
layout = master.CustomLayouts(2)  # 2番目のレイアウト（例: タイトルとコンテンツ）
new_slide = prs.Slides.AddSlide(
    Index=prs.Slides.Count + 1,  # 末尾に追加
    pCustomLayout=layout
)

# 方法2: レイアウト名で検索して追加
def add_slide_by_layout_name(presentation, layout_name, position=None):
    """レイアウト名を指定してスライドを追加する"""
    if position is None:
        position = presentation.Slides.Count + 1

    master = presentation.SlideMaster
    for i in range(1, master.CustomLayouts.Count + 1):
        layout = master.CustomLayouts(i)
        if layout.Name == layout_name:
            return presentation.Slides.AddSlide(position, layout)

    raise ValueError(f"レイアウト '{layout_name}' が見つかりません")

# 使用例
new_slide = add_slide_by_layout_name(prs, "タイトルとコンテンツ")

# 方法3: 既存スライドと同じレイアウトで追加
existing_layout = prs.Slides(1).CustomLayout
new_slide = prs.Slides.AddSlide(prs.Slides.Count + 1, existing_layout)
```

### 6.2 プレースホルダーのタイプで検索してアクセスする

```python
# PpPlaceholderType 定数
PP_PLACEHOLDER_TITLE = 1
PP_PLACEHOLDER_BODY = 2
PP_PLACEHOLDER_CENTER_TITLE = 3
PP_PLACEHOLDER_SUBTITLE = 4
PP_PLACEHOLDER_SLIDE_NUMBER = 13
PP_PLACEHOLDER_FOOTER = 15
PP_PLACEHOLDER_DATE = 16
PP_PLACEHOLDER_PICTURE = 18


def get_placeholder_by_type(slide, ph_type):
    """プレースホルダーをタイプで検索"""
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        if ph.PlaceholderFormat.Type == ph_type:
            return ph
    return None


def get_placeholder_info(slide):
    """スライドの全プレースホルダー情報を取得"""
    result = []
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        pf = ph.PlaceholderFormat
        info = {
            "index": i,
            "name": ph.Name,
            "type": pf.Type,
            "contained_type": pf.ContainedType,
            "left": ph.Left,
            "top": ph.Top,
            "width": ph.Width,
            "height": ph.Height,
            "has_text_frame": bool(ph.HasTextFrame),
        }
        if ph.HasTextFrame:
            info["has_text"] = bool(ph.TextFrame.HasText)
            if ph.TextFrame.HasText:
                info["text"] = ph.TextFrame.TextRange.Text
        result.append(info)
    return result


# 使用例
slide = prs.Slides(1)

# タイトルを取得
title = get_placeholder_by_type(slide, PP_PLACEHOLDER_TITLE)
if title is None:
    # タイトルスライドの中央タイトルも試す
    title = get_placeholder_by_type(slide, PP_PLACEHOLDER_CENTER_TITLE)

if title and title.HasTextFrame:
    title.TextFrame.TextRange.Text = "調査結果レポート"

# 本文を取得
body = get_placeholder_by_type(slide, PP_PLACEHOLDER_BODY)
if body and body.HasTextFrame:
    body.TextFrame.TextRange.Text = "項目1\r項目2\r項目3"

# 全プレースホルダー情報を出力
info = get_placeholder_info(slide)
for item in info:
    print(f"  [{item['index']}] {item['name']}: type={item['type']}")
```

### 6.3 コンテンツプレースホルダーにテキスト・画像・表を入れる

#### テキストの設定

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(2)  # コンテンツプレースホルダー

if ph.HasTextFrame:
    tr = ph.TextFrame.TextRange

    # シンプルなテキスト
    tr.Text = "テキストコンテンツ"

    # 複数段落（箇条書き）
    tr.Text = "第1項目\r第2項目\r第3項目"

    # 段落ごとにインデントレベルを設定
    tr.Paragraphs(1).IndentLevel = 1
    tr.Paragraphs(2).IndentLevel = 2
    tr.Paragraphs(3).IndentLevel = 2

    # 段落ごとに書式を設定
    tr.Paragraphs(1).Font.Bold = True
    tr.Paragraphs(1).Font.Size = 24
```

#### 画像の配置

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(2)  # コンテンツプレースホルダー

# プレースホルダーの位置・サイズを取得
left = ph.Left
top = ph.Top
width = ph.Width
height = ph.Height

# 方法1: プレースホルダーを残して、Fill.UserPicture で設定
# （ストレッチされるため注意）
# ph.Fill.UserPicture(r"C:\path\to\image.jpg")

# 方法2: プレースホルダーの位置に画像を追加（推奨）
# プレースホルダーを削除（必要に応じて）
# ph.Delete()

# 画像を追加
pic = slide.Shapes.AddPicture(
    FileName=r"C:\path\to\image.jpg",
    LinkToFile=0,       # msoFalse
    SaveWithDocument=-1,  # msoTrue
    Left=left,
    Top=top,
    Width=width,
    Height=height
)

# アスペクト比を維持してフィットさせる
pic.LockAspectRatio = True  # msoTrue = -1
# 必要に応じて幅・高さを調整
```

#### 表の配置

```python
slide = prs.Slides(1)
ph = slide.Shapes.Placeholders(2)

# プレースホルダーの位置に表を作成
table_shape = slide.Shapes.AddTable(
    NumRows=4,
    NumColumns=3,
    Left=ph.Left,
    Top=ph.Top,
    Width=ph.Width,
    Height=ph.Height
)

# 表のデータを設定
table = table_shape.Table

# ヘッダー行
headers = ["項目", "数量", "金額"]
for col_idx, header in enumerate(headers, 1):
    cell = table.Cell(1, col_idx)
    cell.Shape.TextFrame.TextRange.Text = header
    cell.Shape.TextFrame.TextRange.Font.Bold = True

# データ行
data = [
    ["商品A", "100", "10,000"],
    ["商品B", "200", "20,000"],
    ["合計", "300", "30,000"],
]
for row_idx, row_data in enumerate(data, 2):
    for col_idx, value in enumerate(row_data, 1):
        table.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text = value

# 必要に応じてプレースホルダーを削除
# ph.Delete()
```

### 6.4 マスターの書式を変更して全スライドに反映する

```python
master = prs.SlideMaster

# === テキストスタイルの変更 ===
# タイトルのフォントを変更（全スライドのタイトルに影響）
title_style = master.TextStyles(1)  # ppTitleStyle
title_style.Levels(1).Font.Name = "Yu Gothic UI Semibold"
title_style.Levels(1).Font.Size = 36
title_style.Levels(1).Font.Color.RGB = 0x333333

# 本文のフォントを変更（全スライドの本文に影響）
body_style = master.TextStyles(2)  # ppBodyStyle
for level in range(1, 6):
    body_style.Levels(level).Font.Name = "Yu Gothic UI"
body_style.Levels(1).Font.Size = 24
body_style.Levels(2).Font.Size = 20
body_style.Levels(3).Font.Size = 18
body_style.Levels(4).Font.Size = 16
body_style.Levels(5).Font.Size = 14

# === 背景の変更 ===
master.Background.Fill.Solid()
master.Background.Fill.ForeColor.RGB = 0xFAFAFA

# === フッターの設定 ===
master.HeadersFooters.Footer.Visible = True
master.HeadersFooters.Footer.Text = "社名 - 社外秘"
master.HeadersFooters.SlideNumber.Visible = True

# === マスターにロゴを追加 ===
logo = master.Shapes.AddPicture(
    FileName=r"C:\path\to\logo.png",
    LinkToFile=0,
    SaveWithDocument=-1,
    Left=master.Width - 100,  # 右上に配置
    Top=10,
    Width=80,
    Height=30
)

# === 全スライドにマスター設定を強制適用 ===
for i in range(1, prs.Slides.Count + 1):
    slide = prs.Slides(i)
    slide.FollowMasterBackground = True
    slide.DisplayMasterShapes = True
    slide.HeadersFooters.Clear()
```

**注意点**:
- マスターの変更は、スライドレベルで「上書き」されていない属性にのみ反映される
- スライドレベルの上書きをリセットするには、`HeadersFooters.Clear()` や `FollowMasterBackground = True` を使用する
- テキストの書式上書きをリセットするには、プレースホルダーを削除して再追加する（`Delete` → `AddPlaceholder`）が最も確実

### 6.5 レイアウトを切り替える

```python
slide = prs.Slides(1)

# 現在のレイアウトを確認
print(f"現在のレイアウト: '{slide.CustomLayout.Name}'")

# レイアウトをインデックスで切り替え
new_layout = prs.SlideMaster.CustomLayouts(3)  # 3番目のレイアウト
slide.CustomLayout = new_layout
print(f"変更後のレイアウト: '{slide.CustomLayout.Name}'")

# レイアウトを名前で切り替え
def change_layout_by_name(slide, layout_name):
    """スライドのレイアウトを名前で変更する"""
    master = slide.Master  # スライドのマスターを取得
    for i in range(1, master.CustomLayouts.Count + 1):
        layout = master.CustomLayouts(i)
        if layout.Name == layout_name:
            slide.CustomLayout = layout
            return True
    return False

# 使用例
success = change_layout_by_name(prs.Slides(1), "白紙")
if success:
    print("レイアウト変更成功")
else:
    print("指定されたレイアウトが見つかりません")

# 全スライドのレイアウトを変更
def change_all_slides_layout(presentation, old_layout_name, new_layout_name):
    """特定レイアウトを使用している全スライドのレイアウトを変更"""
    master = presentation.SlideMaster
    new_layout = None
    for i in range(1, master.CustomLayouts.Count + 1):
        if master.CustomLayouts(i).Name == new_layout_name:
            new_layout = master.CustomLayouts(i)
            break

    if new_layout is None:
        raise ValueError(f"レイアウト '{new_layout_name}' が見つかりません")

    changed_count = 0
    for i in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(i)
        if slide.CustomLayout.Name == old_layout_name:
            slide.CustomLayout = new_layout
            changed_count += 1

    return changed_count
```

### 6.6 プレゼンテーション全体の構造を解析する

```python
def analyze_presentation(presentation):
    """プレゼンテーションの構造を完全に解析する"""
    result = {
        "designs": [],
        "slides": [],
    }

    # デザイン（マスター）情報
    for d_idx in range(1, presentation.Designs.Count + 1):
        design = presentation.Designs(d_idx)
        master = design.SlideMaster

        design_info = {
            "index": d_idx,
            "name": design.Name,
            "layouts": [],
        }

        for l_idx in range(1, master.CustomLayouts.Count + 1):
            layout = master.CustomLayouts(l_idx)
            layout_info = {
                "index": l_idx,
                "name": layout.Name,
                "placeholders": [],
            }

            phs = layout.Shapes.Placeholders
            for p_idx in range(1, phs.Count + 1):
                ph = phs(p_idx)
                layout_info["placeholders"].append({
                    "index": p_idx,
                    "type": ph.PlaceholderFormat.Type,
                    "name": ph.Name,
                })

            design_info["layouts"].append(layout_info)

        result["designs"].append(design_info)

    # スライド情報
    for s_idx in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(s_idx)
        slide_info = {
            "index": s_idx,
            "layout_name": slide.CustomLayout.Name,
            "design_name": slide.Design.Name,
            "follow_master_bg": bool(slide.FollowMasterBackground),
            "display_master_shapes": bool(slide.DisplayMasterShapes),
            "placeholders": [],
        }

        phs = slide.Shapes.Placeholders
        for p_idx in range(1, phs.Count + 1):
            ph = phs(p_idx)
            ph_info = {
                "index": p_idx,
                "type": ph.PlaceholderFormat.Type,
                "name": ph.Name,
                "has_text_frame": bool(ph.HasTextFrame),
            }
            if ph.HasTextFrame and ph.TextFrame.HasText:
                ph_info["text"] = ph.TextFrame.TextRange.Text[:100]  # 先頭100文字

            slide_info["placeholders"].append(ph_info)

        result["slides"].append(slide_info)

    return result


# 使用例
analysis = analyze_presentation(prs)

print("=== デザイン（マスター） ===")
for design in analysis["designs"]:
    print(f"Design: '{design['name']}'")
    for layout in design["layouts"]:
        ph_types = [str(ph['type']) for ph in layout['placeholders']]
        print(f"  Layout: '{layout['name']}' - PH types: [{', '.join(ph_types)}]")

print("\n=== スライド ===")
for slide in analysis["slides"]:
    print(f"Slide {slide['index']}: layout='{slide['layout_name']}', "
          f"design='{slide['design_name']}'")
    for ph in slide["placeholders"]:
        text_preview = ph.get('text', '(empty)')[:50]
        print(f"  PH[{ph['index']}] type={ph['type']} name='{ph['name']}': {text_preview}")
```

### 6.7 テンプレートベースのスライド生成

```python
def create_title_slide(presentation, title, subtitle=None):
    """タイトルスライドを作成"""
    # タイトルスライドレイアウトを検索
    master = presentation.SlideMaster
    layout = None
    for i in range(1, master.CustomLayouts.Count + 1):
        l = master.CustomLayouts(i)
        # タイトルスライドは通常最初のレイアウト
        # ppPlaceholderCenterTitle(3) を持つレイアウトを探す
        phs = l.Shapes.Placeholders
        for j in range(1, phs.Count + 1):
            if phs(j).PlaceholderFormat.Type == 3:  # ppPlaceholderCenterTitle
                layout = l
                break
        if layout:
            break

    if layout is None:
        layout = master.CustomLayouts(1)  # フォールバック

    slide = presentation.Slides.AddSlide(
        presentation.Slides.Count + 1, layout
    )

    # タイトルを設定
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        ph_type = ph.PlaceholderFormat.Type
        if ph_type in (1, 3):  # Title or CenterTitle
            ph.TextFrame.TextRange.Text = title
        elif ph_type == 4 and subtitle:  # Subtitle
            ph.TextFrame.TextRange.Text = subtitle

    return slide


def create_content_slide(presentation, title, content_lines):
    """タイトルとコンテンツのスライドを作成"""
    master = presentation.SlideMaster
    layout = None

    # 「タイトルとコンテンツ」レイアウトを検索
    for i in range(1, master.CustomLayouts.Count + 1):
        l = master.CustomLayouts(i)
        phs = l.Shapes.Placeholders
        has_title = False
        has_body = False
        for j in range(1, phs.Count + 1):
            pt = phs(j).PlaceholderFormat.Type
            if pt == 1:  # ppPlaceholderTitle
                has_title = True
            elif pt == 2:  # ppPlaceholderBody
                has_body = True
        if has_title and has_body:
            layout = l
            break

    if layout is None:
        layout = master.CustomLayouts(2)  # フォールバック

    slide = presentation.Slides.AddSlide(
        presentation.Slides.Count + 1, layout
    )

    # プレースホルダーに内容を設定
    phs = slide.Shapes.Placeholders
    for i in range(1, phs.Count + 1):
        ph = phs(i)
        ph_type = ph.PlaceholderFormat.Type
        if ph_type == 1:  # Title
            ph.TextFrame.TextRange.Text = title
        elif ph_type == 2:  # Body
            ph.TextFrame.TextRange.Text = "\r".join(content_lines)

    return slide


# 使用例
create_title_slide(prs, "四半期レポート", "2026年第1四半期")
create_content_slide(prs, "売上概要", [
    "売上高: 10億円（前年比120%）",
    "営業利益: 2億円（前年比115%）",
    "新規顧客数: 50社",
])
```

---

## 7. 定数・列挙型一覧

### 7.1 PpPlaceholderType（プレースホルダータイプ）

```python
# PpPlaceholderType
ppPlaceholderMixed = -2
ppPlaceholderTitle = 1
ppPlaceholderBody = 2
ppPlaceholderCenterTitle = 3
ppPlaceholderSubtitle = 4
ppPlaceholderVerticalTitle = 5
ppPlaceholderVerticalBody = 6
ppPlaceholderObject = 7
ppPlaceholderChart = 8
ppPlaceholderBitmap = 9
ppPlaceholderMediaClip = 10
ppPlaceholderOrgChart = 11
ppPlaceholderTable = 12
ppPlaceholderSlideNumber = 13
ppPlaceholderHeader = 14
ppPlaceholderFooter = 15
ppPlaceholderDate = 16
ppPlaceholderVerticalObject = 17
ppPlaceholderPicture = 18
ppPlaceholderCameo = 19
```

### 7.2 PpDateTimeFormat（日時形式）

```python
# PpDateTimeFormat
ppDateTimeFormatMixed = -2
ppDateTimeMdyy = 1
ppDateTimeddddMMMMddyyyy = 2
ppDateTimedMMMMyyyy = 3
ppDateTimeMMMMdyyyy = 4
ppDateTimedMMMyy = 5
ppDateTimeMMMMyy = 6
ppDateTimeMMyy = 7
ppDateTimeMMddyyHmm = 8
ppDateTimeMMddyyhmmAMPM = 9
ppDateTimeHmm = 10
ppDateTimeHmmss = 11
ppDateTimehmmAMPM = 12
ppDateTimehmmssAMPM = 13
ppDateTimeFigureOut = 14
ppDateTimeUAQ1 = 15  # ユーザー定義1（ロケール依存）
ppDateTimeUAQ2 = 16
ppDateTimeUAQ3 = 17
ppDateTimeUAQ4 = 18
ppDateTimeUAQ5 = 19
ppDateTimeUAQ6 = 20
ppDateTimeUAQ7 = 21
```

### 7.3 MsoShapeType（シェイプタイプ、ContainedType で使用）

```python
# MsoShapeType（主要なもの）
msoAutoShape = 1
msoCallout = 2
msoChart = 3
msoComment = 4
msoFreeform = 5
msoGroup = 6
msoEmbeddedOLEObject = 7
msoFormControl = 8
msoLine = 9
msoLinkedOLEObject = 10
msoLinkedPicture = 11
msoOLEControlObject = 12
msoPicture = 13
msoPlaceholder = 14
msoTextEffect = 15
msoMedia = 16
msoTextBox = 17
msoTable = 19
msoCanvas = 20
msoDiagram = 21
msoInk = 22
msoInkComment = 23
msoSmartArt = 24
msoSlicer = 25
msoWebVideo = 26
msoContentApp = 27
msoGraphic = 28
msoLinkedGraphic = 29
mso3DModel = 30
msoLinked3DModel = 31
```

### 7.4 PpTextStyleType（テキストスタイルタイプ）

```python
# PpTextStyleType
ppDefaultStyle = 1
ppTitleStyle = 1   # 注意: Microsoft ドキュメントでは ppTitleStyle=1
ppBodyStyle = 2
ppDefaultStyle = 3
# 実際のVBA定数:
# ppTitleStyle = 1
# ppBodyStyle = 2
# ppDefaultStyle = 3
```

**注意**: win32com で定数を使用する場合は `win32com.client.constants` を使用するか、数値を直接指定する。

```python
# 定数の使用方法
import win32com.client
app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")

# EnsureDispatch を使うと定数にアクセス可能
from win32com.client import constants
# constants.ppPlaceholderTitle  # 1
# constants.ppPlaceholderBody   # 2

# または数値を直接使用（推奨: 確実に動作する）
PP_PLACEHOLDER_TITLE = 1
PP_PLACEHOLDER_BODY = 2
```

### 7.5 MsoTriState（True/False 三値）

```python
# MsoTriState
msoTrue = -1
msoFalse = 0
msoCTrue = 1
msoTriStateToggle = -3
msoTriStateMixed = -2
```

### 7.6 PpSlideLayout（レガシーレイアウト定数）

旧来の `Slides.Add` メソッドで使用される定数。`Slides.AddSlide` では `CustomLayout` を使用するため、通常は不要。

```python
# PpSlideLayout（参考）
ppLayoutBlank = 12
ppLayoutChart = 8
ppLayoutChartAndText = 6
ppLayoutClipartAndText = 10
ppLayoutClipArtAndVerticalText = 26
ppLayoutCustom = 32
ppLayoutFourObjects = 24
ppLayoutLargeObject = 15
ppLayoutMediaClipAndText = 18
ppLayoutMixed = -2
ppLayoutObject = 16
ppLayoutObjectAndText = 14
ppLayoutObjectOverText = 19
ppLayoutOrgchart = 7
ppLayoutTable = 4
ppLayoutText = 2
ppLayoutTextAndChart = 5
ppLayoutTextAndClipart = 9
ppLayoutTextAndMediaClip = 17
ppLayoutTextAndObject = 13
ppLayoutTextAndTwoObjects = 21
ppLayoutTextOverObject = 20
ppLayoutTitle = 1
ppLayoutTitleOnly = 11
ppLayoutTwoColumnText = 3
ppLayoutTwoObjects = 29
ppLayoutTwoObjectsAndObject = 30
ppLayoutTwoObjectsAndText = 22
ppLayoutTwoObjectsOverText = 23
ppLayoutVerticalText = 25
ppLayoutVerticalTitleAndText = 27
ppLayoutVerticalTitleAndTextOverChart = 28
```

---

## 重要な注意事項とゴッチャ (Gotchas)

### インデックスの扱い

1. **Placeholders(index) のインデックス**: プレースホルダーのインデックスは1から始まり、レイアウトで定義された順序に基づく。プレースホルダーを削除してもインデックスに欠番が生じる可能性がある。
2. **Shapes コレクションとの違い**: `Shapes(i)` と `Shapes.Placeholders(i)` のインデックスは異なる。`Shapes.Title` は `Shapes.Placeholders(1)` と同等。
3. **FindByName の使用**: インデックスの不確実性を避けるため、`Placeholders.FindByName` メソッドを使用するのが安全。

### COM オブジェクトの解放

```python
import pythoncom
import gc

# COM オブジェクトの明示的解放
def release_com_object(obj):
    """COM オブジェクトを安全に解放"""
    try:
        import ctypes
        if obj is not None:
            ctypes.windll.ole32.CoReleaseMarshalData
            del obj
            gc.collect()
    except:
        pass

# 使用後は必ず解放
# release_com_object(prs)
# release_com_object(app)
```

### MsoTriState の扱い

PowerPoint COM では Boolean 値の代わりに `MsoTriState` を使用する場合がある。Python の `True`/`False` が正しく変換されないことがあるため、数値（`-1`/`0`）を使用するのが安全。

```python
# 推奨: 数値を直接使用
shape.Visible = -1   # msoTrue
shape.Visible = 0    # msoFalse

# Python の True/False も多くの場合動作するが、保証はない
shape.Visible = True   # 通常は動作する
```

### RGB カラー値

PowerPoint COM の RGB 値は BGR 形式（0xBBGGRR）である点に注意:

```python
# 赤色: R=255, G=0, B=0
red = 0x0000FF  # BGR形式では 0x0000FF
# または
red = 255  # 赤のみの場合

# 正確な RGB → BGR 変換
def rgb_to_bgr(r, g, b):
    return r + (g << 8) + (b << 16)

# 使用例
color = rgb_to_bgr(255, 0, 0)  # 赤
```

**注意**: 実際には PowerPoint COM の `RGB` プロパティは標準的な `RGB(R, G, B)` 関数と同じ形式（内部的には同じバイト配列）を使用する。Python からは `r + (g * 256) + (b * 65536)` で計算する。

### プレースホルダーへの画像挿入の制限

COM API には、プレースホルダーに直接画像を「挿入」する（UIの「画像を挿入」アイコンと同等の）メソッドが存在しない。代替アプローチ:

1. **Fill.UserPicture**: 背景として画像を設定（アスペクト比は保持されない）
2. **Shapes.AddPicture**: プレースホルダーの位置・サイズを取得して、同じ場所に画像を配置

### AddPlaceholder の制限

`Shapes.AddPlaceholder` は、削除されたプレースホルダーの**復元**にのみ使用可能。新規プレースホルダーの追加には使用できない。スライドが元々持っていた数を超えてプレースホルダーを追加することはできない。

### レガシー API との共存

- `Slides.Add(Index, Layout)`: 旧 API（PpSlideLayout 定数を使用）
- `Slides.AddSlide(Index, CustomLayout)`: 新 API（CustomLayout オブジェクトを使用）
- 新 API の `AddSlide` の使用を推奨

---

## 参考リンク

- [Master object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.master)
- [CustomLayouts object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customlayouts)
- [CustomLayout object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.customlayout)
- [Placeholders object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.placeholders)
- [PlaceholderFormat object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.placeholderformat)
- [PpPlaceholderType enumeration (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype)
- [HeadersFooters object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/PowerPoint.HeadersFooters)
- [HeaderFooter object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.headerfooter)
- [Design object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.design)
- [Designs object (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.designs)
- [Presentation.Designs property (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.designs)
- [Slides.AddSlide method (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.addslide)
- [Shapes.AddPlaceholder method (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addplaceholder)
- [Placeholders.FindByName method (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.placeholders.findbyname)
- [PpDateTimeFormat enumeration (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat)
- [Slide.DisplayMasterShapes property (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.displaymastershapes)
- [Slide.FollowMasterBackground property (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.followmasterbackground)
- [Shapes.AddTable method (PowerPoint) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addtable)
