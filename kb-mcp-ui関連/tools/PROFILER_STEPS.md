# Upload Profiler - Step分類の詳細

## 概要

`tools/profile_upload.py`は、ファイルアップロード処理を5つのステップに分類してパフォーマンスを測定します。

## Step 1: File I/O

**目的**: ディスクからのファイル読み込み

**内部処理**:
- `open(file_path, 'rb')`でバイナリモードでファイルを開く
- `f.read()`でファイル全体をメモリに読み込む
- ファイルシステムI/Oの速度を測定

**測定内容**:
- ファイル読み込みにかかる時間
- 読み込んだバイト数

## Step 2: Document Conversion

**目的**: ファイル形式からテキストへの変換

**内部処理**:
- `DocumentLoader.load_document()`を使用
- ファイルタイプに応じた専用ローダーを呼び出し:
  - Excel: `ExcelToMarkdownLoader` - セル内容をMarkdown形式に変換
  - Word: `WordToMarkdownLoader` - docxからMarkdown抽出
  - PowerPoint: `PowerPointLoader` - スライド内容を抽出
  - PDF: `PyPDFLoader` - テキスト抽出
  - Markdown/Text: `CustomTextLoader` - そのまま読み込み
- `Document`オブジェクトのリストを生成

**測定内容**:
- 変換処理にかかる時間
- 生成されたドキュメント数
- 総文字数
- 文字数/秒

## Step 3: Text Splitting

**目的**: ドキュメントをチャンクに分割

**内部処理**:
- `SemanticRAGTextSplitter.split_documents()`を使用
- ファイルタイプに応じた専用スプリッターに処理を委譲:
  - Markdown/Word: `MarkdownSplitter` - ヘッダー構造を保持しながら分割
  - Excel: `ExcelSplitter` - シート/セル構造を考慮
  - PDF: `PDFSplitter` - PDFレイアウトを考慮
  - その他: `DefaultSplitter` - `RecursiveCharacterTextSplitter`を使用
- トークン数計算: `tiktoken_len()`でGPT-3.5トークナイザーを使用
- チャンクサイズ/オーバーラップは`ChunkingConfig`で制御

**測定内容**:
- チャンク分割にかかる時間
- 生成されたチャンク数
- チャンクあたりの時間

**以前のボトルネック（解決済み）**:
- `_log_chunking_details()`メソッドによる詳細デバッグログ出力
- デフォルトで無効化され、パフォーマンスが大幅に改善

## Step 4-5: Server Upload (Embed + Store)

**目的**: エンベディング生成とベクトルDB格納

**内部処理**:

### Step 4相当: Embedding Generation (Dense)
- `QdrantVectorStoreService.add_documents()`内で実行
- `_generate_dense_vectors()`を呼び出し:
  - `HuggingFaceEmbeddings.embed_documents()`を使用
  - モデル: `cl-nagoya/ruri-v3-310m`（768次元）
  - バッチサイズ: 32（デフォルト）
  - デバイス: CPU（設定により変更可能）
  - プレフィックス付与: ドキュメントには`"passage: "`を付与
  
### Sparse Vector Generation (BM25)
- `_generate_sparse_vectors()`を呼び出し:
  - `BM25SparseVectorService`を使用
  - FastEmbed + Sudachi形態素解析
  - 日本語テキストに最適化されたBM25実装
  - 非常に高速（0.1秒程度）

### Step 5相当: Vector DB Upload
- `QdrantClient.upsert()`でベクトルDBに保存:
  - UUID生成による一意なID採番
  - Dense vector（768次元）とSparse vectorを同時格納
  - メタデータ（file_path, chunk_index等）も保存
  - バッチ単位でupsert（デフォルト32チャンク/バッチ）

**測定内容**:
- サーバーアップロード全体にかかる時間
- チャンク数
- チャンクあたりの時間

**現在のボトルネック**:
- **Dense Embedding生成**が全体の99%以上を占める
- CPUでの推論が主な要因
- バッチサイズ32での処理により、98チャンクを4バッチで処理
- 各バッチで平均50-87秒（Dense部分のみ）

## パフォーマンス最適化の履歴

### ✅ 完了: Debug Logging最適化
- **変更前**: チャンクあたり4.11秒（主にログ出力）
- **変更後**: チャンクあたり0.004秒
- **改善率**: **約1000倍高速化**
- **実装**: `ChunkingConfig.enable_chunking_debug_log`フラグを追加してデフォルト無効化

### 🔍 次の最適化候補

1. **Embedding生成の高速化**（最優先）:
   - GPU利用（CUDA/DirectML）
   - より小さなモデルへの変更
   - バッチサイズの最適化
   - ONNX版への切り替え

2. **トークン計算のキャッシング**:
   - 同じテキストのトークン数を再計算しない

3. **並列処理の導入**:
   - チャンク分割とエンベディング生成を並列化

## 使用方法

```powershell
# 単一ファイルをプロファイル
uv run tools/profile_upload.py <ファイルパス>

# ベンチマークファイルを作成
uv run tools/profile_upload.py --create

# 全ベンチマークファイルを測定
uv run tools/profile_upload.py --all
```

## 出力例

```
================================================================================
📊 Upload Performance Profile: excel_100KB.xlsx
================================================================================
File Size: 100.09 KB
Total Time: 225.379s

Step                               Time (s)        %              Details
--------------------------------------------------------------------------------
1. File I/O                           0.000     0.0%                    -
2. Document Conversion                0.271     0.1%               1 docs
3. Text Splitting                     0.395     0.2% 98 chunks, 0.004s/chunk
4-5. Server Upload (Embed + Store)  224.712    99.7% 98 chunks, 2.293s/chunk
================================================================================

🔍 Bottleneck Analysis (Top 3):
  1. 4-5. Server Upload (Embed + Store): 224.712s (99.7%)
  2. 3. Text Splitting: 0.395s (0.2%)
  3. 2. Document Conversion: 0.271s (0.1%)
```
