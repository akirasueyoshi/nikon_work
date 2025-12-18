# 資料間関連度定義システム

エクセル資料から機能仕様書名のリンク情報を抽出し、資料間の関連度マトリクスと検索評価用の正解データを生成します。

## システム構成

```
extract_document_links.py  → エクセルから情報抽出してJSON化
         ↓
document_graph.json        → リンク情報のJSON
         ↓
build_relevance_matrix.py  → 関連度マトリクスと正解データ生成
         ↓
ground_truth.json          → MCPサーバの検索評価用正解データ
```

## 使い方

### Step 1: エクセルファイルからリンク情報を抽出

```bash
python3 extract_document_links.py /path/to/excel/directory
```

**出力ファイル:**
- `document_graph_*.json` - 完全なリンク情報（メイン）
- `link_graph_*.json` - 簡易版グラフ構造
- `link_edges_*.csv` - エッジリスト
- `unmatched_links_*.csv` - マッチングできなかったリンク
- `summary_*.json` - 統計情報

### Step 2: 関連度マトリクスを作成

```bash
python3 build_relevance_matrix.py extraction_results/document_graph_*.json
```

**出力ファイル:**
- `relevance_matrix_*.csv` - 関連度マトリクス（全ペア）
- `relevance_edges_*.csv` - 閾値以上のエッジリスト
- `ground_truth_*.json` - 検索評価用正解データ（**重要**）
- `heatmap_*.png` - 可視化（30x30以下の場合）
- `summary_*.json` - 統計情報

## エクセルファイルのフォーマット要件

### 対応フォーマット

スクリプトは以下の構造を持つエクセルファイルに対応しています：

1. **B列またはC列**に「機能仕様書名」という見出しがある
2. その下に仕様書名が列挙される
3. 「対応内容」という文言が出たら、それ以降は無視

**例:**
```
    A列               B列
    ---               ---
                      機能仕様書名
                      RST-12iAS_20200401.xlsx
                      RST-01iAS_20200204.xlsx
                      開始条件確認データテーブル.xls
                      ...
                      対応内容          <- ここから下は無視
                      各種定義シートの追加
```

## 関連度計算方法

`build_relevance_matrix.py`では以下の方法で関連度を計算します（`method="combined"`）：

### 1. 直接リンク (重み: 0.5)
- A → B のリンクがある場合: 関連度 1.0
- ない場合: 関連度 0.0

### 2. 双方向リンク (重み: 0.3)
- A ⇔ B (相互リンク): 関連度 1.0
- A → B または B → A (片方向): 関連度 0.5

### 3. 共通リンク先 (重み: 0.2)
- Jaccard係数を使用
- 関連度 = |A∩B| / |A∪B|
- 例: A→{C,D,E}, B→{D,E,F} なら 2/4 = 0.5

### 最終的な関連度
```
relevance = direct * 0.5 + bidirectional * 0.3 + common * 0.2
```

## 手動でJSONを編集する場合

自動抽出がうまくいかなかった場合、`document_graph.json`を手動で編集できます。

### document_graph.json の構造

```json
{
  "metadata": {
    "extraction_date": "2024-12-18T09:35:50",
    "source_directory": "/path/to/excel",
    "total_documents": 100,
    "total_matched_links": 450,
    "total_unmatched_links": 20
  },
  "documents": [
    {
      "id": "資料A",
      "filename": "資料A.xlsx",
      "path": "/path/to/資料A.xlsx",
      "normalized_name": "資料A",
      "extracted_links_count": 3,
      "extracted_links": ["資料B", "資料C", "資料D"]
    }
  ],
  "links": [
    {
      "source": "資料A",
      "target": "資料B",
      "original_text": "資料B_20200401.xlsx",
      "match_type": "exact"
    }
  ],
  "unmatched_links": [
    {
      "source": "資料A",
      "original_text": "存在しない資料.xlsx",
      "normalized": "存在しない資料"
    }
  ]
}
```

### 手動編集の手順

1. `unmatched_links`を確認
2. 実際のファイル名と照合
3. `links`配列に手動でエントリを追加

```json
{
  "source": "資料A",
  "target": "資料B",
  "original_text": "手動で追加したリンク",
  "match_type": "manual"
}
```

4. 保存後、再度`build_relevance_matrix.py`を実行

## ground_truth.json の構造と使い方

検索結果の評価に使用する正解データです。

```json
[
  {
    "query_doc": "資料A",
    "relevant_docs": [
      {"doc_id": "資料B", "relevance": 0.85},
      {"doc_id": "資料C", "relevance": 0.72},
      {"doc_id": "資料D", "relevance": 0.45}
    ],
    "total_relevant": 5,
    "threshold": 0.3
  }
]
```

### MCPサーバの検索結果評価例

```python
import json

# 正解データを読み込み
with open("ground_truth.json", "r") as f:
    ground_truth = json.load(f)

# MCPサーバから検索結果を取得（例）
mcp_results = {
    "query": "資料A",
    "results": [
        {"file_name": "資料B", "score": 0.95},
        {"file_name": "資料D", "score": 0.82},
        {"file_name": "資料E", "score": 0.71}
    ]
}

# 評価関数
def evaluate_search(mcp_results, ground_truth, k=5):
    query = mcp_results["query"]
    retrieved = [r["file_name"] for r in mcp_results["results"][:k]]
    
    # 正解データから該当クエリを取得
    gt = next((g for g in ground_truth if g["query_doc"] == query), None)
    if not gt:
        return None
    
    relevant = [r["doc_id"] for r in gt["relevant_docs"]]
    
    # Precision@K
    hits = len(set(retrieved) & set(relevant))
    precision = hits / k if k > 0 else 0
    
    # Recall@K
    recall = hits / len(relevant) if len(relevant) > 0 else 0
    
    # F1-Score
    f1 = 2 * precision * recall / (precision + recall) if (precision + recall) > 0 else 0
    
    return {
        "precision@k": precision,
        "recall@k": recall,
        "f1@k": f1,
        "hits": hits,
        "total_relevant": len(relevant)
    }

# 評価実行
result = evaluate_search(mcp_results, ground_truth, k=5)
print(f"Precision@5: {result['precision@5']:.3f}")
print(f"Recall@5: {result['recall@5']:.3f}")
print(f"F1@5: {result['f1@5']:.3f}")
```

## パラメータ調整

### 関連度計算の重み調整

`build_relevance_matrix.py`の以下の部分を編集：

```python
# 重み付け統合（調整可能）
relevance = (m_direct * 0.5 +           # 直接リンクを最重視
             m_bidirectional * 0.3 +     # 双方向性
             m_common * 0.2)             # 共通性
```

### 閾値の調整

正解データ生成時の閾値を変更：

```python
ground_truth = create_ground_truth(
    relevance_matrix, 
    docs, 
    threshold=0.3,  # この値を調整（0.0～1.0）
    top_k=10
)
```

## トラブルシューティング

### Q1: リンクがマッチングされない

**原因:** ファイル名の表記ゆれ（日付、拡張子など）

**対策:**
1. `unmatched_links_*.csv`を確認
2. `document_graph.json`を手動編集
3. または`extract_document_links.py`の`normalize_doc_name()`関数を調整

### Q2: 関連度が低すぎる/高すぎる

**原因:** 重み付けが適切でない

**対策:**
1. `build_relevance_matrix.py`の重みを調整
2. または別のメソッド（`direct`, `bidirectional`, `common_links`）を試す

```python
relevance_matrix = calculate_relevance_matrix(adjacency, method="bidirectional")
```

### Q3: ヒートマップが文字化けする

**原因:** 日本語フォントがない

**対策:** 気にしなくてOK（CSV/JSONは正常）

## 次のステップ

1. 実際の100ファイルで`extract_document_links.py`を実行
2. `unmatched_links`を確認し、必要なら手動編集
3. `build_relevance_matrix.py`で関連度マトリクス生成
4. `ground_truth.json`をMCPサーバの検索評価に使用
5. 評価結果に基づいて閾値やリランキングを調整

## ファイル一覧

### 実行スクリプト
- `extract_document_links.py` - リンク抽出
- `build_relevance_matrix.py` - 関連度マトリクス生成

### 出力ディレクトリ
- `extraction_results/` - 抽出結果
- `relevance_results/` - 関連度マトリクスと正解データ

### ドキュメント
- `README.md` - このファイル
