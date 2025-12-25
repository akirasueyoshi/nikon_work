# Document Relevance Matrix

エクセルファイルから資料間のリンク情報を抽出し、関連度マトリクスを生成するツール。

## 機能

- 📁 **再帰的なディレクトリ探索**: 指定ディレクトリ以下のすべてのエクセルファイルを自動検出
- 🔗 **リンク抽出**: エクセルファイル内の「機能仕様書名」欄から資料間のリンクを抽出
- 📊 **関連度計算**: 
  - Jaccard係数による共通リンク先ベースの類似度
  - 複合指標（直接リンク + 双方向リンク + 共通リンク）
- 📈 **ヒートマップ生成**: PNG形式のビジュアライゼーション
- 🎯 **検索評価用正解データ**: MCP検索機能の評価に使える正解データを自動生成

## インストール

### 前提条件

- Python 3.9以上
- [uv](https://github.com/astral-sh/uv) (推奨)

### uvを使ったインストール

```bash
# uvのインストール（まだの場合）
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# リポジトリをクローン
git clone <repository-url>
cd document-relevance-matrix

# 依存関係のインストール（自動的に仮想環境が作成されます）
uv sync
```

## 使い方

### 🔄 推奨ワークフロー（2段階）

#### Step 1: リンク抽出（確認・修正可能）

```bash
uv run extract-links <エクセルファイルがあるディレクトリ>
```

**例:**
```bash
uv run extract-links ./test_files
uv run extract-links C:\Users\username\Documents\specs
```

**出力:** `extraction_results/links_extracted_*.json`

このJSONファイルには抽出されたリンク情報が含まれます：
- `documents`: 各ドキュメントと抽出されたリンク
- `links`: マッチしたリンクのリスト
- `unmatched_links`: マッチしなかったリンク

#### 📝 Step 1.5: リンクの確認・手動修正（オプション）

```bash
# JSONファイルをテキストエディタで開く
notepad extraction_results\links_extracted_*.json  # Windows
code extraction_results/links_extracted_*.json     # VS Code
```

**修正例:**
```json
{
  "links": [
    {
      "source": "資料A",
      "target": "資料B",  // ← 修正可能
      "original_text": "...",
      "match_type": "exact"
    }
  ],
  "unmatched_links": [
    // ここから正しいリンクを "links" に移動できる
  ]
}
```

#### Step 2: 関連度計算

```bash
uv run calculate-relevance extraction_results/links_extracted_*.json

# または互換性のため（同じ動作）
uv run build-matrix extraction_results/links_extracted_*.json
```

**出力ファイル** (`relevance_results/`):
- `relevance_matrix_combined_*.csv` - 複合指標による関連度マトリクス
- `relevance_matrix_jaccard_*.csv` - Jaccard係数による関連度マトリクス
- `ground_truth_*.json` - 検索評価用の正解データ
- `heatmap_*.png` - 関連度のヒートマップ（30x30以下の場合）
- `relevance_edges_*.csv` - エッジリスト（閾値0.3以上）
- `summary_*.json` - 統計サマリー

## 出力ファイルの説明

### extraction_results/ (extract-links の出力)
- `document_graph_*.json` - 完全なドキュメントグラフ情報
- `relevance_matrix_jaccard_*.csv` - **Jaccard係数のマトリクス**
- `link_edges_*.csv` - リンクのエッジリスト
- `unmatched_links_*.csv` - 未マッチリンク一覧
- `summary_*.json` - 抽出結果のサマリー

### relevance_results/ (build-matrix の出力)
- `relevance_matrix_*.csv` - **複合指標のマトリクス（CSV）**
- `ground_truth_*.json` - **検索評価用データ**
- `heatmap_*.png` - ヒートマップ画像（30x30以下の場合）
- `relevance_edges_*.csv` - エッジリスト（閾値0.3以上）
- `summary_*.json` - 統計サマリー

### CSVファイルの見方

`relevance_matrix_*.csv`を開くと、以下のような形式になっています：

```csv
,資料A,資料B,資料C
資料A,1.0,0.5,0.3
資料B,0.5,1.0,0.0
資料C,0.3,0.0,1.0
```

- **対角線（1.0）**: 自分自身との関連度
- **0.7以上**: 強い関連
- **0.3～0.7**: 中程度の関連
- **0.3未満**: 弱い関連

## エクセルファイルの形式

ツールは以下の形式のエクセルファイルに対応しています：

### フォーマット1（B列に機能仕様書名）
```
| A列 | B列                          | C列 | ...
|-----|------------------------------|-----|
|     | 機能仕様書名                  |     |
|     | センサー制御仕様_20230401.xlsx |     |
|     | モーター制御仕様_20230315.xlsx |     |
|     | ...                          |     |
|     | 対応内容                     |     | ← ここまで
```

### フォーマット2（C列に機能仕様書名）
```
| A列 | B列 | C列                          | ...
|-----|-----|------------------------------|
|     |     | 機能仕様書名                  |
|     |     | センサー制御仕様_20230401.xlsx |
|     |     | ...                          |
```

## 関連度の計算方法

### Jaccard係数（extract-links）
共通して参照している資料の割合を計算：

```
Jaccard係数 = |A ∩ B| / |A ∪ B|
           = 共通リンク数 / 全リンク数
```

**例:**
- 資料A → {B, C, D}
- 資料E → {C, D, F}
- Jaccard係数 = 2個（C, D）/ 4個（B, C, D, F）= 0.5

### 複合指標（build-matrix）
複数の要素を組み合わせた関連度：

```
関連度 = 0.5 × 直接リンク + 0.3 × 双方向リンク + 0.2 × 共通リンク
```

- **直接リンク**: A→Bの直接的な参照関係
- **双方向リンク**: A→BかつB→Aの相互参照
- **共通リンク**: AとBが共通して参照する資料

## ディレクトリ構造

```
document-relevance-matrix/
├── pyproject.toml              # プロジェクト設定
├── README.md                   # このファイル
├── LICENSE                     # ライセンス
├── .gitignore                  # Git除外設定
├── document_relevance_matrix/  # メインパッケージ
│   ├── __init__.py
│   ├── extract_links.py        # リンク抽出スクリプト
│   ├── build_matrix.py         # マトリクス構築スクリプト
│   └── create_report.py        # レポート作成スクリプト
├── examples/                   # サンプルファイル
│   └── test_files/             # テスト用エクセルファイル
└── docs/                       # ドキュメント
    └── USAGE.md                # 詳細な使用方法
```

## トラブルシューティング

### Excelファイルが見つからない

```bash
# ディレクトリパスが正しいか確認
ls <ディレクトリパス>

# Windowsの場合、パスに日本語が含まれる場合は引用符で囲む
uv run extract-links "C:\Users\日本語\Documents\specs"
```

### ヒートマップが生成されない

ドキュメント数が30を超える場合、ヒートマップは自動的にスキップされます。これは表示が複雑になりすぎるためです。

### パスの区切り文字の違い

- **Windows**: バックスラッシュ `\` または `/`
- **Linux/Mac**: スラッシュ `/`

どちらの形式でも動作します。

## ライセンス

MIT License

## 作者

Akira Sueyoshi

## 更新履歴

### v1.0.0 (2024-12-18)
- 初回リリース
- リンク抽出機能
- 関連度マトリクス生成
- Excelレポート生成
