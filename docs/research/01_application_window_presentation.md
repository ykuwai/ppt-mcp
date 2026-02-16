# PowerPoint COM オートメーション調査レポート
# Application / Window管理 / Presentation管理

> 調査日: 2026-02-16
> 目的: MCP Server実装のための PowerPoint COM API 調査

---

## 目次

1. [Application Object (PowerPoint.Application)](#1-application-object)
2. [Window管理 (DocumentWindow / SlideShowWindow)](#2-window管理)
3. [Presentation管理](#3-presentation管理)
4. [SlideShow制御](#4-slideshow制御)
5. [PrintOptions / Export](#5-printoptions--export)
6. [定数一覧 (Enumerations)](#6-定数一覧)
7. [制限事項・注意点](#7-制限事項注意点)

---

## 1. Application Object

### 1.1 Application の起動・接続

PowerPoint COM オートメーションのエントリーポイント。Python では `win32com.client` モジュールを使用する。

#### 新規インスタンスの作成 (Dispatch / CreateObject相当)

```python
import win32com.client

# 新しい PowerPoint インスタンスを起動
app = win32com.client.Dispatch("PowerPoint.Application")

# EnsureDispatch を使うと型情報がキャッシュされ、定数にアクセスしやすくなる
app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
```

**注意**: `Dispatch` は既存のインスタンスがあればそれに接続し、なければ新規作成する場合がある。PowerPoint は複数インスタンスを同時に実行できないため、既に起動中であればそのインスタンスへの参照が返される。

#### 既存インスタンスへの接続 (GetActiveObject / GetObject相当)

```python
import win32com.client

# 既に起動している PowerPoint インスタンスに接続
# PowerPoint が起動していない場合は例外が発生する
try:
    app = win32com.client.GetActiveObject("PowerPoint.Application")
except Exception as e:
    print(f"PowerPoint が起動していません: {e}")
    # 起動していなければ新規作成
    app = win32com.client.Dispatch("PowerPoint.Application")
```

#### DispatchWithEvents (イベント対応)

```python
import win32com.client
import pythoncom

class PowerPointEvents:
    """PowerPoint Application イベントハンドラ"""

    def OnPresentationOpen(self, Pres):
        print(f"プレゼンテーションが開かれました: {Pres.Name}")

    def OnPresentationClose(self, Pres):
        print(f"プレゼンテーションが閉じられました: {Pres.Name}")

    def OnSlideShowBegin(self, Wn):
        print("スライドショーが開始されました")

    def OnSlideShowEnd(self, Pres):
        print("スライドショーが終了しました")

    def OnNewPresentation(self, Pres):
        print(f"新規プレゼンテーション: {Pres.Name}")

    def OnPresentationSave(self, Pres):
        print(f"保存されました: {Pres.Name}")

    def OnWindowActivate(self, Pres, Wn):
        print(f"ウィンドウがアクティブに: {Pres.Name}")

# イベント付きで Dispatch
app = win32com.client.DispatchWithEvents("PowerPoint.Application", PowerPointEvents)

# イベントループ（メッセージポンプ）
while True:
    pythoncom.PumpWaitingMessages()
```

**重要**: `DispatchWithEvents` のイベントハンドラメソッド名は `On` + イベント名 の形式にする。

### 1.2 主要プロパティ

| プロパティ | 型 | R/W | 説明 |
|---|---|---|---|
| `Visible` | MsoTriState | R/W | アプリケーションウィンドウの表示/非表示 |
| `WindowState` | PpWindowState | R/W | ウィンドウの状態（最大化、最小化、通常） |
| `ActiveWindow` | DocumentWindow | R | アクティブなドキュメントウィンドウを返す |
| `ActivePresentation` | Presentation | R | アクティブなプレゼンテーションを返す |
| `Presentations` | Presentations | R | 開いている全プレゼンテーションのコレクション |
| `Windows` | DocumentWindows | R | 全ドキュメントウィンドウのコレクション |
| `SlideShowWindows` | SlideShowWindows | R | 全スライドショーウィンドウのコレクション |
| `Caption` | String | R/W | タイトルバーに表示されるテキスト |
| `Path` | String | R | PowerPoint.exe のパス |
| `Version` | String | R | PowerPoint のバージョン文字列 |
| `ProductCode` | String | R | プロダクトコード |
| `Name` | String | R | "Microsoft PowerPoint" |
| `OperatingSystem` | String | R | OS情報 |
| `Left` | Single | R/W | アプリケーションウィンドウの左端位置（ポイント） |
| `Top` | Single | R/W | アプリケーションウィンドウの上端位置（ポイント） |
| `Width` | Single | R/W | アプリケーションウィンドウの幅（ポイント） |
| `Height` | Single | R/W | アプリケーションウィンドウの高さ（ポイント） |

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")

# アプリケーションを表示する
app.Visible = True  # msoTrue = -1

# ウィンドウを最大化
# ppWindowMaximized = 3, ppWindowMinimized = 2, ppWindowNormal = 1
app.WindowState = 3  # ppWindowMaximized

# バージョン確認
print(f"Version: {app.Version}")
print(f"Path: {app.Path}")
print(f"Name: {app.Name}")

# アクティブプレゼンテーション
if app.Presentations.Count > 0:
    pres = app.ActivePresentation
    print(f"Active: {pres.Name}")

# アクティブウィンドウ
if app.Windows.Count > 0:
    win = app.ActiveWindow
    print(f"Window: {win.Caption}")
```

### 1.3 主要メソッド

| メソッド | 説明 |
|---|---|
| `Quit()` | PowerPoint アプリケーションを終了する |
| `Activate()` | PowerPoint ウィンドウをアクティブにする（前面に） |
| `Help(HelpFile, ContextID)` | ヘルプを表示 |
| `Run(MacroName, ...)` | VBA マクロを実行する |

```python
# アプリケーションをアクティブにする
app.Activate()

# VBA マクロの実行
# app.Run("Module1.MyMacro", arg1, arg2)

# アプリケーションの終了
app.Quit()
app = None  # COM参照を解放
```

**Quit() の注意点**: 未保存のプレゼンテーションがある場合、保存ダイアログが表示される可能性がある。自動化環境では事前に `DisplayAlerts` を設定するか、すべてのプレゼンテーションを明示的に保存/閉じてから `Quit()` を呼ぶ必要がある。

### 1.4 Application Events 一覧

PowerPoint Application オブジェクトが発火するイベントの一覧。Python では `DispatchWithEvents` で `On` プレフィックスを付けたメソッド名で受信する。

#### プレゼンテーション関連イベント

| イベント名 | パラメータ | 説明 |
|---|---|---|
| `NewPresentation` | Pres | 新規プレゼンテーション作成時 |
| `AfterNewPresentation` | Pres | 新規プレゼンテーション作成後 |
| `PresentationOpen` | Pres | プレゼンテーションを開いた時 |
| `PresentationSave` | Pres | 保存時 |
| `PresentationBeforeSave` | Pres, Cancel | 保存前（Cancel=Trueで保存をキャンセル可能） |
| `PresentationBeforeClose` | Pres, Cancel | 閉じる前（Cancel=Trueで閉じることをキャンセル可能） |
| `PresentationClose` | Pres | プレゼンテーションを閉じた時 |
| `PresentationCloseFinal` | Pres | プレゼンテーション完全クローズ時 |
| `PresentationNewSlide` | Sld | 新しいスライドが追加された時 |
| `PresentationPrint` | Pres | 印刷時 |
| `PresentationSync` | Pres, SyncType | 同期時 |

#### スライドショー関連イベント

| イベント名 | パラメータ | 説明 |
|---|---|---|
| `SlideShowBegin` | Wn (SlideShowWindow) | スライドショー開始時 |
| `SlideShowEnd` | Pres | スライドショー終了時 |
| `SlideShowNextBuild` | Wn | 次のアニメーション/ビルド時 |
| `SlideShowNextSlide` | Wn | 次のスライドに進んだ時 |
| `SlideShowNextClick` | Wn, nEffect | クリック時 |
| `SlideShowOnNext` | Wn | 「次へ」操作時 |
| `SlideShowOnPrevious` | Wn | 「前へ」操作時 |

#### ウィンドウ関連イベント

| イベント名 | パラメータ | 説明 |
|---|---|---|
| `WindowActivate` | Pres, Wn | ウィンドウがアクティブになった時 |
| `WindowDeactivate` | Pres, Wn | ウィンドウが非アクティブになった時 |
| `WindowBeforeDoubleClick` | Sel, Cancel | ダブルクリック前 |
| `WindowBeforeRightClick` | Sel, Cancel | 右クリック前 |
| `WindowSelectionChange` | Sel | 選択が変更された時 |
| `SlideSelectionChanged` | SldRange | スライド選択が変更された時 |

**制限事項**: `SlideShowNextSlide` イベントは、自動再生（タイマーによるスライド切り替え）では発火しない場合がある。ポーリングによる代替手段が必要になることがある。

---

## 2. Window管理

### 2.1 DocumentWindow オブジェクト

DocumentWindow はプレゼンテーションの編集ウィンドウを表す。`Application.Windows` コレクションまたは `Application.ActiveWindow` でアクセスする。

**重要**: `DocumentWindows` コレクションにはスライドショーウィンドウは含まれない。スライドショーウィンドウは `SlideShowWindows` コレクションで管理される。

#### プロパティ

| プロパティ | 型 | R/W | 説明 |
|---|---|---|---|
| `Active` | MsoTriState | R | ウィンドウがアクティブかどうか |
| `ActivePane` | Pane | R | アクティブなペイン |
| `Caption` | String | R | ウィンドウのキャプション |
| `Height` | Single | R/W | ウィンドウの高さ（ポイント） |
| `Width` | Single | R/W | ウィンドウの幅（ポイント） |
| `Left` | Single | R/W | ウィンドウの左端位置（ポイント） |
| `Top` | Single | R/W | ウィンドウの上端位置（ポイント） |
| `WindowState` | PpWindowState | R/W | ウィンドウの状態 |
| `Presentation` | Presentation | R | このウィンドウで表示中のプレゼンテーション |
| `Selection` | Selection | R | 現在の選択範囲 |
| `View` | View | R | ウィンドウのビューオブジェクト |
| `ViewType` | PpViewType | R/W | ビューの種類 |
| `Panes` | Panes | R | ペインのコレクション |
| `SplitHorizontal` | Long | R/W | アウトラインペインの幅（画面幅の%） |
| `SplitVertical` | Long | R/W | スライドペインの高さ（画面高さの%） |
| `BlackAndWhite` | MsoTriState | R/W | 白黒表示かどうか |

#### メソッド

| メソッド | 説明 |
|---|---|
| `Activate()` | ウィンドウをアクティブにする |
| `Close()` | ウィンドウを閉じる |
| `FitToPage()` | スライドをウィンドウに合わせて表示 |
| `NewWindow()` | 同じプレゼンテーションの新しいウィンドウを開く |
| `LargeScroll(Down, Up, ToRight, ToLeft)` | 大きくスクロール |
| `SmallScroll(Down, Up, ToRight, ToLeft)` | 小さくスクロール |
| `ScrollIntoView(Left, Top, Width, Height, Start)` | 指定領域までスクロール |
| `ExpandSection(sectionIndex, Expand)` | セクションを展開/折りたたみ |
| `IsSectionExpanded(sectionIndex)` | セクションが展開されているか |
| `PointsToScreenPixelsX(Points)` | ポイントをスクリーンピクセルX座標に変換 |
| `PointsToScreenPixelsY(Points)` | ポイントをスクリーンピクセルY座標に変換 |
| `RangeFromPoint(x, y)` | スクリーン座標からShapeRange/Slideを取得 |

```python
import win32com.client

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True

# アクティブウィンドウの取得
win = app.ActiveWindow

# ウィンドウの位置とサイズ
print(f"位置: ({win.Left}, {win.Top})")
print(f"サイズ: {win.Width} x {win.Height}")

# ウィンドウの位置・サイズ変更
win.Left = 100
win.Top = 100
win.Width = 800
win.Height = 600

# ウィンドウ状態
# ppWindowNormal = 1, ppWindowMinimized = 2, ppWindowMaximized = 3
win.WindowState = 3  # 最大化

# ウィンドウの複製（同じプレゼンテーションを別ウィンドウで表示）
new_win = win.NewWindow()

# スクロール
win.SmallScroll(Down=3)  # 下に3回分スクロール
win.LargeScroll(Down=1)  # 1ページ分下にスクロール

# ページに合わせて表示
win.FitToPage()

# ウィンドウを閉じる
# new_win.Close()
```

### 2.2 ビューの種類 (PpViewType)

`DocumentWindow.ViewType` プロパティで取得・設定できるビューの種類。

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppViewSlide` | 1 | スライドビュー |
| `ppViewSlideMaster` | 2 | スライドマスタービュー |
| `ppViewNotesPage` | 3 | ノートページビュー |
| `ppViewHandoutMaster` | 4 | 配布資料マスタービュー |
| `ppViewNotesMaster` | 5 | ノートマスタービュー |
| `ppViewOutline` | 6 | アウトラインビュー |
| `ppViewSlideSorter` | 7 | スライド一覧ビュー |
| `ppViewTitleMaster` | 8 | タイトルマスタービュー |
| `ppViewNormal` | 9 | 標準ビュー |
| `ppViewPrintPreview` | 10 | 印刷プレビュー |
| `ppViewThumbnails` | 11 | サムネイルビュー |
| `ppViewMasterThumbnails` | 12 | マスターサムネイルビュー |

```python
# ビューの切り替え
win = app.ActiveWindow

# 現在のビューを確認
current_view = win.ViewType
print(f"現在のビュー: {current_view}")

# 標準ビューに切り替え
win.ViewType = 9  # ppViewNormal

# スライド一覧ビューに切り替え
win.ViewType = 7  # ppViewSlideSorter

# アウトラインビューに切り替え
win.ViewType = 6  # ppViewOutline

# ノートページビューに切り替え
win.ViewType = 3  # ppViewNotesPage

# スライドマスタービューに切り替え
win.ViewType = 2  # ppViewSlideMaster
```

### 2.3 View オブジェクト（ズーム・スライド移動）

`DocumentWindow.View` プロパティで取得できるビューオブジェクト。

```python
view = app.ActiveWindow.View

# ズーム率の取得・設定（パーセント）
print(f"ズーム: {view.Zoom}%")
view.Zoom = 100  # 100% に設定
view.Zoom = 150  # 150% に設定

# 特定のスライドに移動（標準ビュー / スライドビューの場合）
view.GotoSlide(3)  # スライド3に移動

# 貼り付け
# view.Paste()
# view.PasteSpecial(DataType)
```

### 2.4 Pane 操作

標準ビュー (Normal View) では3つのペインが存在する:
- Pane(1): アウトラインペイン（サムネイルペイン）
- Pane(2): スライドペイン
- Pane(3): ノートペイン

```python
win = app.ActiveWindow

# ペイン数の確認
print(f"ペイン数: {win.Panes.Count}")

# アクティブペインの確認
active_pane = win.ActivePane
print(f"アクティブペインのViewType: {active_pane.ViewType}")

# 特定のペインをアクティブにする
# アウトラインペインをアクティブにする場合
if win.ActivePane.ViewType != 1:  # ppViewOutline = 6 ではなく ViewType=1
    win.Panes(1).Activate()

# ノートペインをアクティブにする
win.Panes(3).Activate()

# スライドペインをアクティブにする
win.Panes(2).Activate()

# ペインの分割比率を調整
win.SplitHorizontal = 20  # アウトラインペインを画面幅の20%に
win.SplitVertical = 80    # スライドペインを画面高さの80%に
```

### 2.5 SlideShowWindow オブジェクト

スライドショー実行中のウィンドウを表す。`Application.SlideShowWindows` コレクションでアクセス。

| プロパティ | 型 | R/W | 説明 |
|---|---|---|---|
| `View` | SlideShowView | R | スライドショービュー |
| `Presentation` | Presentation | R | 実行中のプレゼンテーション |
| `IsFullScreen` | Boolean | R | フルスクリーンかどうか |
| `Height` | Single | R/W | ウィンドウの高さ |
| `Width` | Single | R/W | ウィンドウの幅 |
| `Left` | Single | R/W | ウィンドウの左端位置 |
| `Top` | Single | R/W | ウィンドウの上端位置 |
| `Active` | MsoTriState | R | アクティブかどうか |

```python
# スライドショーウィンドウの取得
if app.SlideShowWindows.Count > 0:
    ssw = app.SlideShowWindows(1)
    print(f"フルスクリーン: {ssw.IsFullScreen}")
    print(f"サイズ: {ssw.Width} x {ssw.Height}")

    # スライドショービューへのアクセス
    view = ssw.View
    print(f"現在のスライド位置: {view.CurrentShowPosition}")
```

### 2.6 複数ウィンドウの管理

```python
# 全ドキュメントウィンドウの列挙
for i in range(1, app.Windows.Count + 1):
    win = app.Windows(i)
    print(f"Window {i}: {win.Caption} (Active: {win.Active})")
    print(f"  Presentation: {win.Presentation.Name}")
    print(f"  ViewType: {win.ViewType}")
    print(f"  State: {win.WindowState}")

# 全スライドショーウィンドウの列挙
for i in range(1, app.SlideShowWindows.Count + 1):
    ssw = app.SlideShowWindows(i)
    print(f"SlideShow {i}: {ssw.Presentation.Name}")

# Windows(1) は常にアクティブウィンドウ = ActiveWindow と同じ
assert app.Windows(1) == app.ActiveWindow  # (概念的に)
```

---

## 3. Presentation管理

### 3.1 新規作成

```python
# 新規プレゼンテーションの作成
pres = app.Presentations.Add()
# Add() は WithWindow パラメータを受け取る (デフォルト: msoTrue)
# pres = app.Presentations.Add(WithWindow=False)  # ウィンドウなしで作成

print(f"スライド数: {pres.Slides.Count}")  # 0 (空のプレゼンテーション)

# 空白のスライドを追加
# ppLayoutBlank = 12, ppLayoutTitle = 1, ppLayoutText = 2
slide = pres.Slides.Add(1, 12)  # 位置1に空白レイアウトを追加
```

### 3.2 開く (Presentations.Open)

```python
# 基本的なファイルを開く
pres = app.Presentations.Open(
    FileName=r"C:\path\to\presentation.pptx"
)

# 読み取り専用で開く
pres = app.Presentations.Open(
    FileName=r"C:\path\to\presentation.pptx",
    ReadOnly=True  # msoTrue = -1
)

# タイトルなし（コピーとして開く）
pres = app.Presentations.Open(
    FileName=r"C:\path\to\presentation.pptx",
    Untitled=True  # msoTrue
)

# ウィンドウなしで開く（バックグラウンド処理用）
pres = app.Presentations.Open(
    FileName=r"C:\path\to\presentation.pptx",
    WithWindow=False  # msoFalse = 0
)
```

**Presentations.Open パラメータ一覧:**

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| `FileName` | Yes | String | ファイルパス |
| `ReadOnly` | No | MsoTriState | 読み取り専用で開く（デフォルト: msoFalse） |
| `Untitled` | No | MsoTriState | タイトルなしで開く（コピーとして） |
| `WithWindow` | No | MsoTriState | ウィンドウの表示（デフォルト: msoTrue） |

**対応ファイル形式**: .pptx, .pptm, .ppt, .potx, .potm, .pot, .ppsx, .ppsm, .pps, .ppam, .ppa, .thmx, .htm, .html, .mhtml, .rtf, .txt, .doc, .docx, .docm, .xls, .xlsx, .xlsm, .odp 他

**注意**: パスワード保護されたファイルを開くための直接的なパスワードパラメータは `Presentations.Open` にはない。パスワード保護されたファイルはダイアログが表示される。自動化の場合は `Presentations.Open2007` メソッドを使用するか、別の手段が必要。

### 3.3 保存 (Save / SaveAs / SaveCopyAs)

#### 上書き保存

```python
# 上書き保存
pres.Save()
```

#### 名前を付けて保存 (SaveAs)

```python
# PPTX形式で保存
pres.SaveAs(
    FileName=r"C:\output\presentation.pptx",
    FileFormat=24  # ppSaveAsOpenXMLPresentation
)

# PDF形式で保存
pres.SaveAs(
    FileName=r"C:\output\presentation.pdf",
    FileFormat=32  # ppSaveAsPDF
)

# PNG形式で保存（各スライドが個別の画像ファイルに）
pres.SaveAs(
    FileName=r"C:\output\slides",
    FileFormat=18  # ppSaveAsPNG
)

# JPG形式で保存
pres.SaveAs(
    FileName=r"C:\output\slides",
    FileFormat=17  # ppSaveAsJPG
)

# フォントを埋め込んで保存
pres.SaveAs(
    FileName=r"C:\output\with_fonts.pptx",
    FileFormat=24,
    EmbedFonts=True  # msoTrue
)
```

**SaveAs パラメータ:**

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| `FileName` | Yes | String | 保存先パス |
| `FileFormat` | No | PpSaveAsFileType | ファイル形式（デフォルト: ppSaveAsDefault = 11） |
| `EmbedFonts` | No | MsoTriState | フォント埋め込み |

#### PpSaveAsFileType 主要定数

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppSaveAsPresentation` | 1 | PPT形式（旧形式） |
| `ppSaveAsTemplate` | 5 | POT形式（旧テンプレート） |
| `ppSaveAsRTF` | 6 | RTF形式 |
| `ppSaveAsShow` | 7 | PPS形式（旧スライドショー） |
| `ppSaveAsAddIn` | 8 | PPA形式（アドイン） |
| `ppSaveAsDefault` | 11 | デフォルト形式 |
| `ppSaveAsMetaFile` | 15 | WMF形式 |
| `ppSaveAsGIF` | 16 | GIF形式 |
| `ppSaveAsJPG` | 17 | JPEG形式 |
| `ppSaveAsPNG` | 18 | PNG形式 |
| `ppSaveAsBMP` | 19 | BMP形式 |
| `ppSaveAsTIF` | 21 | TIFF形式 |
| `ppSaveAsEMF` | 23 | EMF形式 |
| `ppSaveAsOpenXMLPresentation` | 24 | PPTX形式 |
| `ppSaveAsOpenXMLPresentationMacroEnabled` | 25 | PPTM形式 |
| `ppSaveAsOpenXMLTemplate` | 26 | POTX形式 |
| `ppSaveAsOpenXMLTemplateMacroEnabled` | 27 | POTM形式 |
| `ppSaveAsOpenXMLShow` | 28 | PPSX形式 |
| `ppSaveAsOpenXMLShowMacroEnabled` | 29 | PPSM形式 |
| `ppSaveAsOpenXMLAddin` | 30 | PPAM形式 |
| `ppSaveAsOpenXMLTheme` | 31 | THMX形式（テーマ） |
| `ppSaveAsPDF` | 32 | PDF形式 |
| `ppSaveAsXPS` | 33 | XPS形式 |
| `ppSaveAsXMLPresentation` | 34 | XML形式 |
| `ppSaveAsOpenDocumentPresentation` | 35 | ODP形式 |
| `ppSaveAsOpenXMLPicturePresentation` | 36 | 画像プレゼンテーション |
| `ppSaveAsWMV` | 37 | WMV形式（動画） |
| `ppSaveAsStrictOpenXMLPresentation` | 38 | Strict Open XML |
| `ppSaveAsMP4` | 39 | MP4形式（動画） |
| `ppSaveAsAnimatedGIF` | 40 | アニメーションGIF |

#### コピーとして保存

```python
# 現在のファイル名はそのままに、別名でコピーを保存
pres.SaveCopyAs(r"C:\output\backup_copy.pptx")
```

### 3.4 閉じる

```python
# プレゼンテーションを閉じる
pres.Close()
# Close() は変更がある場合、保存ダイアログを表示する可能性がある

# 保存せずに閉じたい場合は、事前に Saved プロパティを True に設定
pres.Saved = True  # 「変更なし」とマーク
pres.Close()

# または保存してから閉じる
pres.Save()
pres.Close()
```

### 3.5 プレゼンテーションのプロパティ (BuiltInDocumentProperties)

```python
# BuiltInDocumentProperties へのアクセス
props = pres.BuiltInDocumentProperties

# プロパティの読み取り
title = props("Title").Value
author = props("Author").Value
subject = props("Subject").Value
keywords = props("Keywords").Value
comments = props("Comments").Value
category = props("Category").Value
company = props("Company").Value
manager = props("Manager").Value

print(f"タイトル: {title}")
print(f"作成者: {author}")
print(f"件名: {subject}")
print(f"キーワード: {keywords}")

# プロパティの設定
props("Title").Value = "新しいタイトル"
props("Author").Value = "太郎"
props("Subject").Value = "テスト件名"
props("Keywords").Value = "PowerPoint, 自動化, COM"
props("Category").Value = "レポート"
props("Comments").Value = "自動生成されたプレゼンテーション"

# CustomDocumentProperties（カスタムプロパティ）
custom_props = pres.CustomDocumentProperties
# カスタムプロパティの追加
# custom_props.Add(Name="MyProperty", LinkToContent=False,
#                  Type=4, Value="MyValue")
# Type: 1=msoPropertyTypeBoolean, 2=msoPropertyTypeNumber,
#        3=msoPropertyTypeDate, 4=msoPropertyTypeString
```

**主要な BuiltInDocumentProperties 一覧:**

| プロパティ名 | 説明 |
|---|---|
| `Title` | タイトル |
| `Subject` | 件名 |
| `Author` | 作成者 |
| `Keywords` | キーワード |
| `Comments` | コメント |
| `Template` | テンプレート |
| `Last Author` | 最終更新者 |
| `Revision Number` | リビジョン番号 |
| `Application Name` | アプリケーション名 |
| `Last Print Date` | 最終印刷日 |
| `Creation Date` | 作成日 |
| `Last Save Time` | 最終保存日時 |
| `Total Editing Time` | 合計編集時間 |
| `Number of Pages` | ページ数（=スライド数） |
| `Number of Words` | 単語数 |
| `Number of Characters` | 文字数 |
| `Security` | セキュリティ |
| `Category` | カテゴリ |
| `Manager` | 管理者 |
| `Company` | 会社名 |

### 3.6 テンプレート適用

```python
# テンプレートの適用（デザインテーマ）
pres.ApplyTemplate(r"C:\path\to\template.potx")

# テーマの適用（.thmx ファイル）
pres.ApplyTheme(r"C:\path\to\theme.thmx")

# プレゼンテーションの TemplateName を確認
print(f"テンプレート: {pres.TemplateName}")
```

### 3.7 PageSetup (スライドサイズ・向き)

```python
# PageSetup オブジェクトの取得
page = pres.PageSetup

# スライドサイズの取得（ポイント単位: 1インチ = 72ポイント）
print(f"幅: {page.SlideWidth} pt ({page.SlideWidth / 72} inch)")
print(f"高さ: {page.SlideHeight} pt ({page.SlideHeight / 72} inch)")
print(f"最初のスライド番号: {page.FirstSlideNumber}")

# 標準 16:9 (13.333" x 7.5")
page.SlideWidth = 13.333 * 72   # 960 pt
page.SlideHeight = 7.5 * 72     # 540 pt

# 標準 4:3 (10" x 7.5")
page.SlideWidth = 10 * 72    # 720 pt
page.SlideHeight = 7.5 * 72  # 540 pt

# A4 横向き (11.693" x 8.268")
page.SlideWidth = 11.693 * 72
page.SlideHeight = 8.268 * 72

# A4 縦向き (7.5" x 10")
page.SlideWidth = 7.5 * 72
page.SlideHeight = 10 * 72

# 最初のスライド番号を設定
page.FirstSlideNumber = 1

# 向きの設定
# ppOrientationHorizontal = 1 (横), ppOrientationVertical = 2 (縦)
# ※ SlideOrientation プロパティ は PowerPoint 2003 以前。
# 現在は SlideWidth と SlideHeight で直接制御するのが推奨。

# ノート/配布資料の向き
page.NotesOrientation = 2  # ppOrientationVertical (縦)
```

**PageSetup プロパティ一覧:**

| プロパティ | 型 | R/W | 説明 |
|---|---|---|---|
| `SlideWidth` | Single | R/W | スライドの幅（ポイント） |
| `SlideHeight` | Single | R/W | スライドの高さ（ポイント） |
| `FirstSlideNumber` | Long | R/W | 最初のスライド番号 |
| `NotesOrientation` | MsoOrientation | R/W | ノートの向き |
| `SlideOrientation` | MsoOrientation | R/W | スライドの向き（非推奨） |
| `SlideSize` | PpSlideSizeType | R/W | スライドサイズの種類 |

### 3.8 SectionProperties (セクション管理)

プレゼンテーション内のセクションを管理するオブジェクト。

```python
# SectionProperties の取得
sections = pres.SectionProperties

# セクション数の確認
print(f"セクション数: {sections.Count}")

# セクションの追加
# AddSection(sectionIndex, sectionName)
# sectionIndex: 挿入位置（既存セクションの前に挿入）
new_section_idx = sections.AddSection(1, "はじめに")

# 末尾にセクションを追加（Count+1 を指定可能。最大512セクション）
new_section_idx = sections.AddSection(sections.Count + 1, "まとめ")

# セクション名の取得
for i in range(1, sections.Count + 1):
    name = sections.Name(i)
    first_slide = sections.FirstSlide(i)
    slide_count = sections.SlidesCount(i)
    section_id = sections.SectionID(i)
    print(f"Section {i}: '{name}' (First Slide: {first_slide}, Slides: {slide_count})")

# セクション名の変更（Rename）
sections.Rename(1, "新しい名前")  # Rename(sectionIndex, sectionName)

# セクションの移動
sections.Move(1, 3)  # Move(sectionIndex, toPos) - セクション1を位置3に移動

# セクションの削除
sections.Delete(1, True)  # Delete(sectionIndex, deleteSlides)
# deleteSlides=True: セクション内のスライドも削除
# deleteSlides=False: セクション区切りのみ削除（スライドは残る）
```

**SectionProperties メソッド一覧:**

| メソッド | パラメータ | 説明 |
|---|---|---|
| `AddSection` | sectionIndex, sectionName | セクションを追加。追加したセクションのインデックスを返す |
| `AddBeforeSlide` | slideIndex, sectionName | 指定スライドの直前にセクションを追加 |
| `Count` | (なし) | セクション数を返す |
| `Delete` | sectionIndex, deleteSlides | セクションを削除 |
| `FirstSlide` | sectionIndex | セクションの最初のスライド番号を返す |
| `Move` | sectionIndex, toPos | セクションを移動 |
| `Name` | sectionIndex | セクション名を返す |
| `Rename` | sectionIndex, sectionName | セクション名を変更 |
| `SectionID` | sectionIndex | セクションのIDを返す |
| `SlidesCount` | sectionIndex | セクション内のスライド数を返す |

**注意**: セクションのインデックスは1始まり。最大512セクションまで作成可能。

### 3.9 その他のプレゼンテーションプロパティ

```python
# プレゼンテーションの基本情報
print(f"名前: {pres.Name}")
print(f"フルパス: {pres.FullName}")
print(f"パス: {pres.Path}")
print(f"スライド数: {pres.Slides.Count}")
print(f"読み取り専用: {pres.ReadOnly}")
print(f"保存済み: {pres.Saved}")
print(f"最終保存者: {pres.BuiltInDocumentProperties('Last Author').Value}")

# ファイル形式の確認（Presentation が保存されている形式）
# pres.FileFormat  # 定数はなく、数値で返る

# パスワード保護
# pres.Password = "secret"           # 開くパスワード
# pres.WritePassword = "writesecret" # 書き込みパスワード

# ファイナル（最終版としてマーク）
# pres.Final = True
```

---

## 4. SlideShow制御

### 4.1 スライドショーの開始

```python
# SlideShowSettings の取得と設定
ss = pres.SlideShowSettings

# 全スライドを表示（デフォルト）
ss.RangeType = 1  # ppShowAll

# 特定範囲のスライドを表示
ss.RangeType = 2  # ppShowSlideRange
ss.StartingSlide = 2
ss.EndingSlide = 5

# ループ再生
ss.LoopUntilStopped = True  # msoTrue

# アドバンスモード
# ppSlideShowManualAdvance = 1 (手動)
# ppSlideShowUseSlideTimings = 2 (タイミング使用)
# ppSlideShowRehearseNewTimings = 3 (リハーサル)
ss.AdvanceMode = 1  # 手動

# ナレーション付き
ss.ShowWithNarration = True  # msoTrue

# アニメーション付き
ss.ShowWithAnimation = True  # msoTrue

# 発表者ツール表示
ss.ShowPresenterView = True

# スクロールバー表示
ss.ShowScrollbar = True

# メディアコントロール表示
ss.ShowMediaControls = True

# ショータイプ
# ppShowTypeSpeaker = 1 (発表者として、フルスクリーン)
# ppShowTypeWindow = 2 (ウィンドウ表示)
# ppShowTypeKiosk = 3 (キオスク、フルスクリーン)
ss.ShowType = 1

# ポインターの色
ss.PointerColor.RGB = 0x0000FF  # 赤色 (BGR形式: Blue, Green, Red)

# スライドショーを実行
slide_show_window = ss.Run()

# Run() は SlideShowWindow オブジェクトを返す
```

### 4.2 SlideShowSettings プロパティ一覧

| プロパティ | 型 | R/W | 説明 |
|---|---|---|---|
| `AdvanceMode` | PpSlideShowAdvanceMode | R/W | スライドの進め方 |
| `EndingSlide` | Long | R/W | 終了スライド番号 |
| `StartingSlide` | Long | R/W | 開始スライド番号 |
| `LoopUntilStopped` | MsoTriState | R/W | ループ再生 |
| `NamedSlideShows` | NamedSlideShows | R | 名前付きスライドショーのコレクション |
| `PointerColor` | ColorFormat | R | ポインターの色 |
| `RangeType` | PpSlideShowRangeType | R/W | 表示範囲の種類 |
| `ShowMediaControls` | MsoTriState | R/W | メディアコントロール表示 |
| `ShowPresenterView` | MsoTriState | R/W | 発表者ツール表示 |
| `ShowScrollbar` | MsoTriState | R/W | スクロールバー表示 |
| `ShowType` | PpSlideShowType | R/W | ショータイプ |
| `ShowWithAnimation` | MsoTriState | R/W | アニメーション表示 |
| `ShowWithNarration` | MsoTriState | R/W | ナレーション再生 |
| `SlideShowName` | String | R/W | カスタムスライドショー名 |

### 4.3 SlideShowView の操作

```python
# スライドショー実行中のビューの取得
ssw = app.SlideShowWindows(1)
view = ssw.View

# ナビゲーション
view.First()      # 最初のスライドに移動
view.Last()       # 最後のスライドに移動
view.Next()       # 次のスライドに移動
view.Previous()   # 前のスライドに移動
view.GotoSlide(5) # スライド5に移動
# GotoSlide(Index, ResetSlide)
# ResetSlide=True (デフォルト): アニメーションをリセット
# ResetSlide=False: アニメーション済みの状態を維持

# カスタムスライドショーへの切り替え
# view.GotoNamedShow("MyCustomShow")
# view.EndNamedShow()  # カスタムショーから元のショーに戻る

# 現在のスライド情報
current_position = view.CurrentShowPosition  # スライドショー内の位置
current_slide = view.Slide                    # 現在のSlideオブジェクト
print(f"現在のスライド: {current_slide.SlideIndex}")
print(f"ショー内位置: {current_position}")

# スライドショーの状態
# ppSlideShowRunning = 1
# ppSlideShowPaused = 2
# ppSlideShowBlackScreen = 3
# ppSlideShowWhiteScreen = 4
# ppSlideShowDone = 5
state = view.State
print(f"状態: {state}")

# 一時停止・黒画面・白画面
view.State = 2  # 一時停止
view.State = 3  # 黒画面
view.State = 4  # 白画面
view.State = 1  # 再開

# 時間情報
print(f"プレゼンテーション経過時間: {view.PresentationElapsedTime}秒")
print(f"現在のスライド経過時間: {view.SlideElapsedTime}秒")
view.ResetSlideTime()  # 現在のスライドのタイマーをリセット

# ズーム
print(f"ズーム: {view.Zoom}%")

# アドバンスモードの確認
print(f"アドバンスモード: {view.AdvanceMode}")

# クリック情報
click_count = view.GetClickCount()
click_index = view.GetClickIndex()
print(f"クリック数: {click_count}, 現在のクリックインデックス: {click_index}")

# 特定のクリックに移動
# view.GotoClick(3)  # 3番目のクリックに移動

# ペン描画
# view.DrawLine(x1, y1, x2, y2)  # 線を描画
# view.EraseDrawing()             # 描画を消去

# スライドショーの終了
view.Exit()
```

### 4.4 ポインタの種類変更

```python
view = app.SlideShowWindows(1).View

# ポインタの種類を設定
# ppSlideShowPointerNone = 0        ポインタなし
# ppSlideShowPointerArrow = 1       矢印
# ppSlideShowPointerPen = 2         ペン
# ppSlideShowPointerAlwaysHidden = 3 常に非表示
# ppSlideShowPointerAutoArrow = 4   自動矢印
# ppSlideShowPointerEraser = 5      消しゴム

view.PointerType = 2  # ペンに変更

# ポインタの色を設定
view.PointerColor.RGB = 0x0000FF  # 赤 (BGR)

# レーザーポインター
print(f"レーザーポインター有効: {view.LaserPointerEnabled}")
view.LaserPointerEnabled = True
```

### 4.5 SlideShowView プロパティ・メソッド一覧

#### メソッド

| メソッド | パラメータ | 説明 |
|---|---|---|
| `First` | (なし) | 最初のスライドに移動 |
| `Last` | (なし) | 最後のスライドに移動 |
| `Next` | (なし) | 次のスライドに移動 |
| `Previous` | (なし) | 前のスライドに移動 |
| `GotoSlide` | Index, ResetSlide | 指定スライドに移動 |
| `GotoClick` | Index | 指定クリックに移動 |
| `GotoNamedShow` | SlideShowName | カスタムスライドショーに移動 |
| `EndNamedShow` | (なし) | カスタムスライドショーから戻る |
| `Exit` | (なし) | スライドショーを終了 |
| `DrawLine` | BeginX, BeginY, EndX, EndY | 線を描画 |
| `EraseDrawing` | (なし) | 描画を消去 |
| `GetClickCount` | (なし) | クリック数を取得 |
| `GetClickIndex` | (なし) | 現在のクリックインデックスを取得 |
| `FirstAnimationIsAutomatic` | (なし) | 最初のアニメーションが自動かどうか |
| `Player` | ShapeIndex | メディアプレーヤーを取得 |
| `ResetSlideTime` | (なし) | スライドタイマーをリセット |

#### プロパティ

| プロパティ | 型 | R/W | 説明 |
|---|---|---|---|
| `AcceleratorsEnabled` | Boolean | R/W | ショートカットキーの有効/無効 |
| `AdvanceMode` | PpSlideShowAdvanceMode | R | アドバンスモード |
| `CurrentShowPosition` | Long | R | スライドショー内の現在位置 |
| `IsNamedShow` | Boolean | R | カスタムスライドショーかどうか |
| `LaserPointerEnabled` | Boolean | R/W | レーザーポインター有効/無効 |
| `LastSlideViewed` | Slide | R | 直前に表示したスライド |
| `PointerColor` | ColorFormat | R | ポインターの色 |
| `PointerType` | PpSlideShowPointerType | R/W | ポインターの種類 |
| `PresentationElapsedTime` | Single | R/W | 開始からの経過時間（秒） |
| `Slide` | Slide | R | 現在のスライド |
| `SlideElapsedTime` | Single | R/W | 現在のスライドの経過時間（秒） |
| `SlideShowName` | String | R | カスタムスライドショー名 |
| `State` | PpSlideShowState | R/W | スライドショーの状態 |
| `Zoom` | Integer | R | ズーム率 |
| `MediaControlsHeight` | Single | R | メディアコントロールの高さ |
| `MediaControlsLeft` | Single | R | メディアコントロールの左位置 |
| `MediaControlsTop` | Single | R | メディアコントロールの上位置 |
| `MediaControlsVisible` | Boolean | R | メディアコントロールの可視状態 |
| `MediaControlsWidth` | Single | R | メディアコントロールの幅 |

---

## 5. PrintOptions / Export

### 5.1 印刷設定と実行

#### PrintOptions オブジェクト

```python
# PrintOptions の取得
print_opts = pres.PrintOptions

# 印刷設定
print_opts.NumberOfCopies = 2           # 部数
print_opts.Collate = True               # 部単位で印刷
print_opts.PrintColorType = 1           # ppPrintColor=1, ppPrintBlackAndWhite=2, ppPrintPureBlackAndWhite=3
print_opts.PrintHiddenSlides = True     # 非表示スライドも印刷
print_opts.FitToPage = True             # ページに合わせる
print_opts.FrameSlides = True           # スライドを枠で囲む

# 出力タイプ
# ppPrintOutputSlides = 1          スライド
# ppPrintOutputTwoSlideHandouts = 2   2スライド/ページ
# ppPrintOutputThreeSlideHandouts = 3  3スライド/ページ
# ppPrintOutputSixSlideHandouts = 4   6スライド/ページ
# ppPrintOutputNotesPages = 5      ノートページ
# ppPrintOutputOutline = 6         アウトライン
# ppPrintOutputBuildSlides = 7     ビルドスライド
# ppPrintOutputFourSlideHandouts = 8  4スライド/ページ
# ppPrintOutputNineSlideHandouts = 9  9スライド/ページ
# ppPrintOutputOneSlideHandouts = 10  1スライド/ページ
print_opts.OutputType = 1  # スライド

# 印刷範囲の設定
# ppPrintAll = 1       全スライド
# ppPrintSelection = 2  選択範囲
# ppPrintCurrent = 3    現在のスライド
# ppPrintSlideRange = 4 スライド範囲
# ppPrintNamedSlideShow = 5 名前付きスライドショー
print_opts.RangeType = 4  # ppPrintSlideRange
print_opts.Ranges.ClearAll()
print_opts.Ranges.Add(1, 1)   # スライド1
print_opts.Ranges.Add(3, 5)   # スライド3-5
print_opts.Ranges.Add(8, 9)   # スライド8-9

# ActivePrinter の設定
# print_opts.ActivePrinter = "Microsoft Print to PDF"
```

#### PrintOut メソッド

```python
# 基本的な印刷
pres.PrintOut()

# パラメータ指定の印刷
pres.PrintOut(
    From=2,        # 開始スライド
    To=5,          # 終了スライド
    Copies=2,      # 部数
    Collate=False  # 部単位印刷しない
)

# ファイルに出力
pres.PrintOut(
    PrintToFile="C:\\output\\print_output.prn"
)

# 非表示スライドを含めて印刷
pres.PrintOptions.PrintHiddenSlides = True
pres.PrintOut(From=1, To=10, Copies=1)
```

**PrintOut パラメータ:**

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| `From` | No | Integer | 開始ページ番号 |
| `To` | No | Integer | 終了ページ番号 |
| `PrintToFile` | No | String | 出力先ファイル名（指定するとプリンタではなくファイルに出力） |
| `Copies` | No | Integer | 部数（デフォルト: 1） |
| `Collate` | No | MsoTriState | 部単位印刷 |

**PrintOptions プロパティ一覧:**

| プロパティ | 型 | R/W | 説明 |
|---|---|---|---|
| `ActivePrinter` | String | R/W | アクティブプリンター名 |
| `Collate` | MsoTriState | R/W | 部単位印刷 |
| `FitToPage` | MsoTriState | R/W | ページに合わせる |
| `FrameSlides` | MsoTriState | R/W | スライドの枠表示 |
| `HighQuality` | MsoTriState | R/W | 高品質印刷 |
| `NumberOfCopies` | Long | R/W | 部数 |
| `OutputType` | PpPrintOutputType | R/W | 出力タイプ |
| `PrintColorType` | PpPrintColorType | R/W | 色の種類 |
| `PrintComments` | MsoTriState | R/W | コメント印刷 |
| `PrintFontsAsGraphics` | MsoTriState | R/W | フォントをグラフィックとして印刷 |
| `PrintHiddenSlides` | MsoTriState | R/W | 非表示スライドを印刷 |
| `PrintInBackground` | MsoTriState | R/W | バックグラウンド印刷 |
| `RangeType` | PpPrintRangeType | R/W | 印刷範囲の種類 |
| `Ranges` | PrintRanges | R | 印刷範囲のコレクション |
| `SlideShowName` | String | R/W | スライドショー名 |

### 5.2 PDF出力

PDF出力には主に2つの方法がある。

#### 方法1: SaveAs で PDF

```python
# SaveAs でPDF出力（最もシンプル）
pres.SaveAs(
    FileName=r"C:\output\presentation.pdf",
    FileFormat=32  # ppSaveAsPDF
)
```

#### 方法2: ExportAsFixedFormat でPDF（詳細制御可能）

```python
# ExportAsFixedFormat でPDF出力（より細かい制御が可能）
pres.ExportAsFixedFormat(
    Path=r"C:\output\presentation_detailed.pdf",
    FixedFormatType=2,         # ppFixedFormatTypePDF = 2
    Intent=2,                   # ppFixedFormatIntentPrint = 2 (印刷品質)
                                # ppFixedFormatIntentScreen = 1 (画面表示品質)
    FrameSlides=False,          # msoFalse: スライド枠なし
    HandoutOrder=2,             # ppPrintHandoutVerticalFirst = 2
    OutputType=1,               # ppPrintOutputSlides = 1
    PrintHiddenSlides=False,    # msoFalse: 非表示スライドを含めない
    PrintRange=None,            # None = 全スライド
    RangeType=1,                # ppPrintAll = 1
    SlideShowName="",
    IncludeDocProperties=True,  # ドキュメントプロパティを含める
    KeepIRMSettings=True,       # IRM設定を保持
    DocStructureTags=True,      # 構造タグを含める（アクセシビリティ）
    BitmapMissingFonts=True,    # 不足フォントをビットマップ化
    UseISO19005_1=False         # PDF/A準拠にするか
)
```

**ExportAsFixedFormat パラメータ:**

| パラメータ | 必須 | 型 | 説明 |
|---|---|---|---|
| `Path` | Yes | String | 出力先パス |
| `FixedFormatType` | Yes | PpFixedFormatType | 形式 (PDF=2, XPS=1) |
| `Intent` | No | PpFixedFormatIntent | 用途 (Screen=1, Print=2) |
| `FrameSlides` | No | MsoTriState | スライド枠 |
| `HandoutOrder` | No | PpPrintHandoutOrder | 配布資料の順序 |
| `OutputType` | No | PpPrintOutputType | 出力タイプ |
| `PrintHiddenSlides` | No | MsoTriState | 非表示スライド含む |
| `PrintRange` | Yes* | PrintRange | 印刷範囲（None可） |
| `RangeType` | No | PpPrintRangeType | 範囲の種類 |
| `SlideShowName` | No | String | スライドショー名 |
| `IncludeDocProperties` | No | Boolean | ドキュメントプロパティ含む |
| `KeepIRMSettings` | No | Boolean | IRM設定保持 |
| `DocStructureTags` | No | Boolean | 構造タグ含む |
| `BitmapMissingFonts` | No | Boolean | 不足フォントビットマップ化 |
| `UseISO19005_1` | No | Boolean | PDF/A準拠 |
| `ExternalExporter` | No | Variant | 外部エクスポーター |

### 5.3 画像出力 (各スライドを画像として)

#### 方法1: SaveAs で全スライドを画像出力

```python
import os

# PNG形式で全スライドを画像出力
# 指定したディレクトリ内にフォルダが作成され、各スライドが個別のPNGファイルになる
output_dir = r"C:\output\slides_png"
pres.SaveAs(output_dir, 18)  # ppSaveAsPNG = 18

# JPG形式で全スライドを画像出力
output_dir = r"C:\output\slides_jpg"
pres.SaveAs(output_dir, 17)  # ppSaveAsJPG = 17

# BMP形式で全スライドを画像出力
output_dir = r"C:\output\slides_bmp"
pres.SaveAs(output_dir, 19)  # ppSaveAsBMP = 19

# TIFF形式で全スライドを画像出力
output_dir = r"C:\output\slides_tif"
pres.SaveAs(output_dir, 21)  # ppSaveAsTIF = 21

# GIF形式で全スライドを画像出力
output_dir = r"C:\output\slides_gif"
pres.SaveAs(output_dir, 16)  # ppSaveAsGIF = 16

# EMF形式で全スライドを画像出力
output_dir = r"C:\output\slides_emf"
pres.SaveAs(output_dir, 23)  # ppSaveAsEMF = 23
```

**注意**: `SaveAs` で画像形式を指定すると、指定パスの名前でフォルダが作成され、その中に `スライド1.PNG`, `スライド2.PNG`, ... のようなファイルが生成される。フォルダ名は指定したパスのファイル名部分が使用される。

#### 方法2: Export メソッドで個別スライドを画像出力

```python
# Presentation.Export メソッドで画像出力
# Export(Path, FilterName, ScaleWidth, ScaleHeight)
pres.Export(
    Path=r"C:\output\all_slides",
    FilterName="PNG",
    ScaleWidth=1920,   # 幅（ピクセル）
    ScaleHeight=1080   # 高さ（ピクセル）
)

# 個別スライドの Export
for i in range(1, pres.Slides.Count + 1):
    slide = pres.Slides(i)
    output_path = rf"C:\output\slide_{i}.png"
    slide.Export(output_path, "PNG", 1920, 1080)
    print(f"スライド {i} を出力: {output_path}")
```

**Export メソッドのフィルター名:**

| フィルター名 | 説明 |
|---|---|
| `"PNG"` | PNG形式 |
| `"JPG"` | JPEG形式 |
| `"BMP"` | BMP形式 |
| `"TIF"` | TIFF形式 |
| `"GIF"` | GIF形式 |
| `"EMF"` | EMF形式 |
| `"WMF"` | WMF形式 |

**Export メソッドの解像度**: `ScaleWidth` と `ScaleHeight` で出力画像の解像度を指定できる。省略した場合はデフォルトの解像度（96dpi相当）が使用される。

### 5.4 動画出力

```python
# WMV形式で動画出力
pres.SaveAs(r"C:\output\presentation.wmv", 37)  # ppSaveAsWMV = 37

# MP4形式で動画出力
pres.SaveAs(r"C:\output\presentation.mp4", 39)  # ppSaveAsMP4 = 39

# アニメーションGIF形式で出力
pres.SaveAs(r"C:\output\presentation.gif", 40)  # ppSaveAsAnimatedGIF = 40
```

**注意**: 動画出力は時間がかかる場合があり、また PowerPoint のバージョンによってサポート状況が異なる。MP4/WMV出力にはプレゼンテーションのアニメーション設定やトランジション、タイミングが反映される。

---

## 6. 定数一覧

### 6.1 PpWindowState

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppWindowNormal` | 1 | 通常 |
| `ppWindowMinimized` | 2 | 最小化 |
| `ppWindowMaximized` | 3 | 最大化 |

### 6.2 PpSlideShowState

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppSlideShowRunning` | 1 | 実行中 |
| `ppSlideShowPaused` | 2 | 一時停止 |
| `ppSlideShowBlackScreen` | 3 | 黒画面 |
| `ppSlideShowWhiteScreen` | 4 | 白画面 |
| `ppSlideShowDone` | 5 | 完了 |

### 6.3 PpSlideShowType

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppShowTypeSpeaker` | 1 | 発表者として（フルスクリーン） |
| `ppShowTypeWindow` | 2 | ウィンドウ表示 |
| `ppShowTypeKiosk` | 3 | キオスクモード（自動フルスクリーン） |

### 6.4 PpSlideShowAdvanceMode

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppSlideShowManualAdvance` | 1 | 手動 |
| `ppSlideShowUseSlideTimings` | 2 | スライドのタイミングを使用 |
| `ppSlideShowRehearseNewTimings` | 3 | リハーサル |

### 6.5 PpSlideShowRangeType

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppShowAll` | 1 | 全スライド |
| `ppShowSlideRange` | 2 | スライド範囲 |
| `ppShowNamedSlideShow` | 3 | 名前付きスライドショー |

### 6.6 PpSlideShowPointerType

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppSlideShowPointerNone` | 0 | ポインタなし |
| `ppSlideShowPointerArrow` | 1 | 矢印 |
| `ppSlideShowPointerPen` | 2 | ペン |
| `ppSlideShowPointerAlwaysHidden` | 3 | 常に非表示 |
| `ppSlideShowPointerAutoArrow` | 4 | 自動矢印 |
| `ppSlideShowPointerEraser` | 5 | 消しゴム |

### 6.7 PpPrintRangeType

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppPrintAll` | 1 | 全スライド |
| `ppPrintSelection` | 2 | 選択範囲 |
| `ppPrintCurrent` | 3 | 現在のスライド |
| `ppPrintSlideRange` | 4 | スライド範囲 |
| `ppPrintNamedSlideShow` | 5 | 名前付きスライドショー |

### 6.8 PpPrintOutputType

| 定数名 | 値 | 説明 |
|---|---|---|
| `ppPrintOutputSlides` | 1 | スライド |
| `ppPrintOutputTwoSlideHandouts` | 2 | 2スライド/ページ |
| `ppPrintOutputThreeSlideHandouts` | 3 | 3スライド/ページ |
| `ppPrintOutputSixSlideHandouts` | 4 | 6スライド/ページ |
| `ppPrintOutputNotesPages` | 5 | ノートページ |
| `ppPrintOutputOutline` | 6 | アウトライン |
| `ppPrintOutputBuildSlides` | 7 | ビルドスライド |
| `ppPrintOutputFourSlideHandouts` | 8 | 4スライド/ページ |
| `ppPrintOutputNineSlideHandouts` | 9 | 9スライド/ページ |
| `ppPrintOutputOneSlideHandouts` | 10 | 1スライド/ページ |

### 6.9 MsoTriState

| 定数名 | 値 | 説明 |
|---|---|---|
| `msoTrue` | -1 | True |
| `msoFalse` | 0 | False |
| `msoCTrue` | 1 | True（旧互換） |
| `msoTriStateToggle` | -3 | トグル |
| `msoTriStateMixed` | -2 | 混在 |

---

## 7. 制限事項・注意点

### 7.1 COM スレッディング・アパートメントモデル

```python
import pythoncom

# COM はアパートメントスレッディングモデルを使用
# メインスレッド以外で COM を使用する場合は明示的な初期化が必要
pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

try:
    app = win32com.client.Dispatch("PowerPoint.Application")
    # ... 操作 ...
finally:
    pythoncom.CoUninitialize()
```

**重要な注意事項:**
- COM オブジェクトはスレッドをまたいで使用できない
- MCP Server のように複数リクエストを処理する場合、COM オブジェクトへのアクセスは単一スレッドに限定する必要がある
- `pythoncom.PumpWaitingMessages()` を定期的に呼ばないとイベントが処理されない

### 7.2 PowerPoint の単一インスタンス制限

- PowerPoint は同時に1つのインスタンスしか実行できない（Word/Excel とは異なる）
- `Dispatch("PowerPoint.Application")` を複数回呼んでも、同じインスタンスへの参照が返される
- サーバーベースで使用する場合、排他制御が必要

### 7.3 Visible プロパティの制限

```python
# PowerPoint.Application を非表示にすると一部の操作が制限される場合がある
app.Visible = True  # 表示状態で操作するのが安全

# WithWindow=False で開いた場合、ウィンドウ関連の操作は使用不可
pres = app.Presentations.Open(r"C:\test.pptx", WithWindow=False)
# app.ActiveWindow  # エラーの可能性がある
```

### 7.4 SaveAs の注意点

- `SaveAs` で画像形式を指定すると、フォルダが作成されその中にスライドごとの画像ファイルが生成される
- PDF 出力時に既存ファイルがある場合、上書きされる（確認ダイアログなし）
- `SaveAs` を呼ぶとプレゼンテーションのファイル名が変更される。元の名前を保持したい場合は `SaveCopyAs` を使用する
- 画像形式で `SaveAs` すると、プレゼンテーションの `FullName` が変わるため注意

### 7.5 イベントの制限

- `SlideShowNextSlide` イベントは、自動再生（タイマーベースのスライド切り替え）では発火しない場合がある
- イベントを受信するには `pythoncom.PumpWaitingMessages()` を定期的に呼ぶ必要がある
- `DispatchWithEvents` と `GetActiveObject` は組み合わせにくい。イベントを受信するには `DispatchWithEvents` でインスタンスを作成する必要がある

### 7.6 COM オブジェクトの解放

```python
import gc

# COM オブジェクトは明示的に解放するのが望ましい
pres.Close()
pres = None

app.Quit()
app = None

# ガベージコレクションを強制実行
gc.collect()

# 必要に応じて COM を解放
# pythoncom.CoUninitialize()
```

### 7.7 エラーハンドリング

```python
import win32com.client
import pywintypes

try:
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(r"C:\nonexistent.pptx")
except pywintypes.com_error as e:
    # COM エラー
    hr = e.hresult        # HRESULT エラーコード
    msg = e.strerror      # エラーメッセージ
    excepinfo = e.excepinfo  # 詳細情報のタプル
    print(f"COM Error: {hr:#x} - {msg}")
    if excepinfo:
        source = excepinfo[1]       # エラーソース
        description = excepinfo[2]  # 説明
        print(f"Source: {source}, Description: {description}")
except Exception as e:
    print(f"Error: {e}")
```

### 7.8 MCP Server 実装への推奨事項

1. **接続管理**: `GetActiveObject` で既存インスタンスへの接続を試み、失敗した場合のみ `Dispatch` で新規作成する
2. **スレッド安全性**: COM 操作は単一スレッドで行い、MCP リクエストからの呼び出しはキューイングする
3. **エラーリカバリ**: COM 接続が切れた場合（PowerPoint がクラッシュした場合等）の再接続ロジックを実装する
4. **リソース管理**: 使用後の COM オブジェクト参照は確実に解放する
5. **タイムアウト**: 長時間の操作（PDF出力、動画出力等）にはタイムアウトを設定する
6. **WithWindow パラメータ**: バックグラウンド処理では `WithWindow=False` を活用するが、ウィンドウ関連操作が制限される点に注意
7. **SaveAs vs SaveCopyAs**: ファイル名を変えずにエクスポートしたい場合は `SaveCopyAs` を使用する。ただし `SaveCopyAs` はファイル形式の変換をサポートしない場合がある

---

## 参考リソース

- [Microsoft Learn - PowerPoint VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint)
- [DocumentWindow object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow)
- [SlideShowSettings object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowsettings)
- [SlideShowView object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowview)
- [Presentation.SaveAs method (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas)
- [Presentation.ExportAsFixedFormat method (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat)
- [PpSaveAsFileType enumeration (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype)
- [PpViewType enumeration (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype)
- [PageSetup object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup)
- [SectionProperties object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sectionproperties)
- [PrintOptions object (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.printoptions)
- [Presentations.Open method (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)
- [Presentation.PrintOut method (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.printout)
- [Handling PowerPoint Slide Show Events from Python (Mamezou Developer Portal)](https://developer.mamezou-tech.com/en/blogs/2024/09/02/monitor-pptx-py/)
- [Tim Golden's Python Stuff: Attach to a running instance of a COM application](https://timgolden.me.uk/python/win32_how_do_i/attach-to-a-com-instance.html)
- [WIN32 automation of PowerPoint (GitHub Gist)](https://gist.github.com/dmahugh/f642607d50cd008cc752f1344e9809e6)
- [Controlling PowerPoint with Python via COM32 (Medium)](https://medium.com/@chasekidder/controlling-powerpoint-w-python-52f6f6bf3f2d)
- [SlideShowView class reference (CodeVBA)](https://www.codevba.com/powerpoint/slideshowview.htm)
