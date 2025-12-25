# クイックスタートガイド

## インストール

```bash
# 1. uvのインストール（Windows PowerShell）
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# 2. リポジトリをクローン
git clone <your-repository-url>
cd document-relevance-matrix

# 3. 依存関係のインストール
uv sync
```

## 基本的な使い方

### Step 1: リンク抽出

```bash
# サンプルファイルで試す
uv run extract-links examples/test_files

# 自分のファイルで実行
uv run extract-links C:\path\to\your\excel\files
```

### Step 1.5: リンクを確認・修正（オプション）

```bash
# JSONファイルを開いて確認
notepad extraction_results\links_extracted_*.json
```

リンクが間違っていたら、JSONを直接編集できます！

### Step 2: 関連度計算

```bash
uv run calculate-relevance extraction_results/links_extracted_*.json
```

### Step 3: 結果を確認

```bash
# CSVファイルを開く
start relevance_results/relevance_matrix_combined_*.csv    # Windows (Excel等で開く)
open relevance_results/relevance_matrix_combined_*.csv     # Mac
xdg-open relevance_results/relevance_matrix_combined_*.csv # Linux

# または、ヒートマップ画像を表示
start relevance_results/heatmap_*.png             # Windows
```

## よくある使い方

### パターン1: 単一ディレクトリ

```bash
uv run extract-links ./specs
uv run build-matrix extraction_results/document_graph_*.json
```

### パターン2: 深い階層構造

```bash
# 再帰的に探索されます
uv run extract-links ./project
#   project/
#   ├── module_a/
#   │   ├── spec1.xlsx
#   │   └── spec2.xlsx
#   └── module_b/
#       └── spec3.xlsx
```

### パターン3: パスに日本語が含まれる場合

```bash
uv run extract-links "C:\Users\日本語ユーザー\Documents\仕様書"
```

## 出力ファイルの説明

### extraction_results/ (Step 1の出力)
- `links_extracted_*.json` - **📝 リンク抽出結果（編集可能！）**

### relevance_results/ (Step 2の出力)
- `relevance_matrix_combined_*.csv` - **📊 複合指標のマトリクス**
- `relevance_matrix_jaccard_*.csv` - Jaccard係数のマトリクス
- `ground_truth_*.json` - 検索評価用データ
- `heatmap_*.png` - ヒートマップ画像
- `relevance_edges_*.csv` - エッジリスト

## トラブルシューティング

### エラー: "No Excel files found"
→ ディレクトリパスを確認してください

### ヒートマップが生成されない
→ ドキュメント数が30を超える場合は自動的にスキップされます

### CSVが開けない
→ Excel、LibreOffice、Google Sheets等で開けます

## 次のステップ

- 詳細な使い方: [README.md](README.md)
- Jaccard係数について: README.md の「関連度の計算方法」を参照
