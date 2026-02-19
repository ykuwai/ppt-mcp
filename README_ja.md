<p align="center">
  <img src="https://raw.githubusercontent.com/ykuwai/ppt-mcp/main/assets/ppt-mcp-logo-letter.png" alt="PowerPoint MCP" width="480">
</p>

<p align="center">
  <a href="README.md">English version</a>
</p>

<p align="center">
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License"></a>
  <img src="https://img.shields.io/badge/MCP_Tools-131-orange.svg" alt="Tools">
  <a href="https://modelcontextprotocol.io/"><img src="https://img.shields.io/badge/MCP-1.0+-purple.svg" alt="MCP"></a>
  <img src="https://img.shields.io/badge/Platform-Windows-0078d4.svg" alt="Platform">
</p>

<p align="center">
  <strong>COM自動化によるPowerPointのリアルタイム制御 —<br>AIエージェントと開発者のための131ツールを備えたMCPサーバー</strong>
</p>

---

PowerPointをCOM自動化で完全に制御するMCP（Model Context Protocol）サーバーです。python-pptxのようなファイルベースのライブラリとは異なり、起動中のPowerPointアプリケーションと直接やり取りし、リアルタイムの視覚的フィードバックとPowerPoint APIへの完全なアクセスを提供します。

## ⚙️ MCPクライアント設定

### Claude Code

```bash
claude mcp add powerpoint uvx ppt-mcp
```

### `.mcp.json`（手動設定）

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["ppt-mcp"]
    }
  }
}
```

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json` を編集：

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["ppt-mcp"]
    }
  }
}
```

### ソースから実行（開発用）

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uv",
      "args": [
        "--directory",
        "C:\\path\\to\\ppt-mcp",
        "run",
        "mcp",
        "run",
        "src/server.py"
      ]
    }
  }
}
```

## ✨ なぜこれが特別なのか

### 🎮 ファイル操作ではなく、PowerPointの完全制御

ファイルベースのライブラリは `.pptx` ファイルの読み書きしかできません。COM自動化により、PowerPointのすべての機能にアクセスできます：

- スライドショーの起動・制御・ナビゲーション
- アニメーション効果のリアルタイム追加・編集
- ビデオ・オーディオメディアの埋め込み
- SmartArtグラフィックの作成・編集
- 元に戻す / やり直し操作
- ビュー制御（標準、アウトライン、ノート、マスター表示）
- コメント機能（共同作業）

### 🤖 AIエージェントのために設計

- **21カテゴリ・131ツール** — スライド操作からアニメーション、SmartArt、アイコン検索まで
- **プレゼン操作対象の明示指定** — `ppt_activate_presentation` で対象ファイルを固定。AIエージェントが誤って別のプレゼンを編集することを防止
- **リアルタイム視覚フィードバック** — 編集対象のスライドに自動ナビゲーション。変更がその場で見える
- **テンプレート対応** — 個人用テンプレートフォルダを自動検出し、任意のテンプレートからプレゼンを作成
- **Material Symbolsアイコン** — 2,500以上のGoogle Material Symbolsアイコンをキーワード検索し、テーマカラーでSVG挿入
- **テーマカラー連携** — RGB値のハードコードではなく、`accent1` や `accent2` などテーマカラー名で指定
- **テキスト精密制御** — `\n` で改段落（Enter）、`\v` で改行（Shift+Enter）— テキストフローを完全にコントロール
- **STAスレッド安全性** — すべてのCOM操作を専用のSTAワーカースレッドで実行し、信頼性を確保

## 🛠️ ツール一覧

| カテゴリ | ツール数 | 主な機能 |
|---------|-------:|---------|
| **アプリケーション** | 4 | PowerPoint接続、アプリ情報、ウィンドウ状態、プレゼン一覧 |
| **プレゼンテーション** | 8 | 作成（テンプレート対応）、開く、保存、閉じる、情報取得、操作対象指定、テンプレート一覧 |
| **スライド** | 9 | 追加、削除、複製、移動、一覧、情報取得、ノート、ナビゲーション |
| **シェイプ** | 10 | 図形/テキストボックス/画像/線の追加、一覧、情報取得、更新、削除、Z順序 |
| **テキスト** | 8 | テキスト設定/取得、書式設定、段落書式、箇条書き、検索置換、テキストフレーム |
| **プレースホルダー** | 5 | 一覧、情報取得、テキスト設定 |
| **書式設定** | 3 | 塗りつぶし、線、影 |
| **テーブル** | 9 | テーブル追加、セル取得/設定、セル結合、行/列の追加/削除、スタイル |
| **エクスポート** | 2 | PDF、画像 |
| **スライドショー** | 6 | 開始、停止、次へ、前へ、スライド移動、状態取得 |
| **グラフ** | 6 | グラフ追加、データ設定/取得、書式設定、系列設定、種類変更 |
| **アニメーション** | 5 | トランジション、アニメーション追加/一覧/削除/全削除 |
| **テーマ** | 3 | テーマ適用、テーマカラー取得、ヘッダー/フッター設定 |
| **グループ** | 3 | グループ化、グループ解除、グループ項目取得 |
| **コネクタ** | 2 | 追加、書式設定 |
| **ハイパーリンク** | 3 | 追加、取得、削除 |
| **セクション** | 3 | 追加、一覧、管理 |
| **プロパティ** | 2 | プレゼンテーションメタデータの設定/取得 |
| **メディア** | 3 | ビデオ、オーディオ、メディア設定 |
| **SmartArt** | 3 | 追加、編集、レイアウト一覧 |
| **編集操作** | 6 | 元に戻す、やり直し、スライド間シェイプ/書式コピー |
| **レイアウト** | 7 | 整列、分散配置、スライドサイズ、背景、反転、シェイプ結合 |
| **視覚効果** | 3 | グロー、反射、ぼかし |
| **コメント** | 3 | 追加、一覧、削除 |
| **高度な操作** | 16 | タグ、フォント一括設定/置換、トリミング、シェイプエクスポート、表示/非表示、選択、ビュー、アニメーションコピー、URL画像、SVGアイコン、アイコン検索、縦横比ロック |
| | **131** | |

## 📦 セットアップ

**動作環境: Windows 10/11、Microsoft PowerPoint、Python 3.10以上、[uv](https://docs.astral.sh/uv/getting-started/installation/)**

```bash
git clone https://github.com/ykuwai/ppt-mcp.git
cd ppt-mcp
uv sync
```

## 🚀 使用例

```python
# 1. 操作対象のプレゼンテーションを指定（誤編集を防止）
ppt_list_presentations()
ppt_activate_presentation(presentation_name="demo.pptx")

# 2. テンプレートからプレゼンテーションを作成
ppt_list_templates()
ppt_create_presentation(template_path="C:\\...\\MyTemplate.potx")

# 3. スライド追加とコンテンツ設定
ppt_add_slide(layout_index=2)
ppt_set_text(slide_index=1, shape_name_or_index="Title 1", text="Hello World")

# 4. プレゼン全体のフォントを一括設定（英語・日本語フォントを個別に指定）
ppt_set_default_fonts(latin="Segoe UI", east_asian="Meiryo")

# 5. テーマカラーでMaterial Symbolsアイコンを挿入
ppt_add_svg_icon(slide_index=1, icon_name="rocket",
                 left=500, top=100, width=72, height=72,
                 color="accent1", style="rounded", filled=True)

# 6. PDFにエクスポート
ppt_export_pdf(file_path="C:\\output\\presentation.pdf")
```

## 🔍 機能の詳細

### 🎯 プレゼンテーション操作対象の指定

`ppt_activate_presentation` でセッションレベルの操作対象を設定すると、以降のすべてのツールがそのファイルに対して動作します。PowerPointのウィンドウが切り替わっても影響を受けません。再度呼び出すことで対象を切り替えられます。

```python
ppt_activate_presentation(presentation_name="report.pptx")
# 以降のツールはすべて report.pptx を操作
ppt_activate_presentation(presentation_name="demo.pptx")
# 切り替え — 以降は demo.pptx を操作
```

### 📁 テンプレート対応

個人用のPowerPointテンプレートフォルダを自動検出します（レジストリ、OneDrive、デフォルトパスを順に確認）。`ppt_list_templates` でテンプレートを一覧し、`ppt_create_presentation(template_path=...)` で任意のテンプレートから新規プレゼンテーションを作成できます。

### 🎨 Material Symbolsアイコン

`ppt_search_icons(query="...")` で2,500以上のアイコンをキーワード検索し、`ppt_add_svg_icon` でSVG画像として挿入：
- **3つのスタイル**: outlined、rounded、sharp
- **塗りつぶしバリアント**: `filled=True` で指定
- **テーマカラー**: `color="accent1"` でプレゼンのアクセントカラーを自動適用
- **自動フィット**: 指定エリア内でアスペクト比を保持

### ⚡ リアルタイムナビゲーション

書き込み操作のたびに、PowerPointの画面が自動的に対象スライドに移動します。変更がリアルタイムで目の前に表示されるため、手動でスライドを切り替える必要はありません。

### ✍️ テキスト制御

- `\n` — 改段落（Enter）。段落ごとに独自の箇条書き・インデントレベルを持つ
- `\v` — 改行（Shift+Enter）。同じ段落内に留まり、書式を維持
- `ppt_format_text_range` で文字単位の書式設定
- 自動調整: テキストをシェイプに収める、シェイプをリサイズ、またはオーバーフロー

## 📄 ライセンス

MIT

## 🙏 クレジット

- [FastMCP](https://github.com/jlowin/fastmcp) — Python MCPサーバーフレームワーク
- [pywin32](https://github.com/mhammond/pywin32) — Windows COM自動化
- [Model Context Protocol](https://modelcontextprotocol.io/) — by Anthropic
