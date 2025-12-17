#!/usr/bin/env python3
"""
複数の検索クエリを順次実行して結果を記録するスクリプト（kb-mcp feat-ui対応版）

Usage:
    python batch_search_benchmark.py
"""

import sys
import time
import json
import csv
from pathlib import Path
from typing import List, Dict, Any
from datetime import datetime


class SearchBenchmark:
    """検索ベンチマークを実行するクラス"""
    
    def __init__(self, kb_mcp_root: Path = None):
        if kb_mcp_root is None:
            kb_mcp_root = Path.cwd()
        
        self.kb_mcp_root = kb_mcp_root
        sys.path.insert(0, str(kb_mcp_root))
        
        # MCPモジュールをインポート
        from ui.services.mcp_service import MCPService
        from ui.services.search_service import SearchService
        
        self.mcp_service = MCPService()
        self.search_service = SearchService(self.mcp_service)
        
        self.results = []
        
        print("✓ Initialized MCP client and search service")
    
    def connect(self, mode="stdio", command="uv", args="run src/main.py", http_url=None):
        """MCPサーバに接続"""
        print(f"\n🔌 Connecting to MCP server...")
        
        if mode == "http" and http_url:
            result = self.mcp_service.connect(
                transport_type="Streamable HTTP",
                connection_mode="Manual Connection",
                command="",
                args="",
                url=http_url,
                host="",
                port=""
            )
        else:
            result = self.mcp_service.connect(
                transport_type="Stdio",
                connection_mode="Automatic Launch",
                command=command,
                args=args,
                url="",
                host="",
                port=""
            )
        
        if "✅" in result:
            print(f"✓ {result}")
            time.sleep(1)
            return True
        else:
            print(f"❌ {result}")
            return False
    
    def search_single_query(
        self, 
        query: str, 
        limit: int = 5, 
        mode: str = "hybrid",
        use_rerank: bool = False
    ) -> Dict[str, Any]:
        """単一クエリで検索を実行"""
        try:
            # 時間計測開始
            start_time = time.time()
            
            # SearchServiceを使って検索実行
            response = self.search_service.search(
                query=query,
                limit=limit,
                mode=mode,
                domains_text="",  # 全ドメインを検索
                metadata_filter_text="",  # フィルタなし
                use_rerank=use_rerank,
                dense_threshold_text="",  # デフォルト閾値を使用
                sparse_threshold_text="",  # デフォルト閾値を使用
                rerank_threshold_text="",  # デフォルト閾値を使用
            )
            
            # 時間計測終了
            end_time = time.time()
            elapsed_time = (end_time - start_time) * 1000  # ミリ秒に変換
            
            # 結果の解析
            if response.get("success"):
                data = response.get("data", {})
                results = data.get("results", [])
                query_info = data.get("query_info", {})
                
                result_data = {
                    "query": query,
                    "mode": mode,
                    "limit": limit,
                    "use_rerank": use_rerank,
                    "success": True,
                    "response_time_ms": elapsed_time,
                    "result_count": len(results),
                    "query_info": query_info,
                    "results": []
                }
                
                # 各結果の詳細を記録
                for i, res in enumerate(results, 1):
                    # メタデータから情報を取得
                    metadata = res.get("metadata", {})
                    
                    result_entry = {
                        "rank": i,
                        "file_path": res.get("path", metadata.get("path", "")),
                        "score": res.get("score", 0),
                        "content": res.get("content", res.get("text", "")),  # フルコンテンツを保存
                        "chunk_id": metadata.get("chunk_id", ""),
                        "document_id": metadata.get("document_id", ""),
                        "metadata": metadata,
                    }
                    
                    # スコア詳細（metadataから取得）
                    scores = {}
                    if "dense_score" in metadata:
                        scores["dense"] = metadata["dense_score"]
                    if "sparse_score" in metadata:
                        scores["sparse"] = metadata["sparse_score"]
                    if "rerank_score" in metadata:
                        scores["rerank"] = metadata["rerank_score"]
                    if "rrf_score" in metadata:
                        scores["rrf"] = metadata["rrf_score"]
                    
                    if scores:
                        result_entry["scores"] = scores
                    
                    # チャンク情報
                    if "chunk_index" in metadata:
                        result_entry["chunk_index"] = metadata["chunk_index"]
                    
                    result_data["results"].append(result_entry)
                
                return result_data
            else:
                return {
                    "query": query,
                    "mode": mode,
                    "limit": limit,
                    "use_rerank": use_rerank,
                    "success": False,
                    "response_time_ms": elapsed_time,
                    "result_count": 0,
                    "error": response.get("error", "Unknown error"),
                    "query_info": {},
                    "results": []
                }
                
        except Exception as e:
            return {
                "query": query,
                "mode": mode,
                "limit": limit,
                "use_rerank": use_rerank,
                "success": False,
                "response_time_ms": 0,
                "result_count": 0,
                "error": str(e),
                "query_info": {},
                "results": []
            }
    
    def run_benchmark(
        self, 
        queries: List[str], 
        limit: int = 5, 
        mode: str = "hybrid",
        use_rerank: bool = False
    ):
        """複数クエリのベンチマークを実行（ターミナルモニタリング対応）"""
        total = len(queries)
        
        print(f"\n{'='*60}")
        print(f"🔍 Starting search benchmark: {total} quer{'y' if total == 1 else 'ies'}")
        print(f"   Mode: {mode}, Limit: {limit}, Rerank: {use_rerank}")
        print(f"{'='*60}\n")
        
        for i, query in enumerate(queries, 1):
            print(f"[{i}/{total}] Searching: {query}")
            
            result = self.search_single_query(query, limit, mode, use_rerank)
            self.results.append(result)
            
            if result["success"]:
                print(f"  ✅ Found {result['result_count']} results in {result['response_time_ms']:.2f}ms")
                
                # 上位3件の結果を表示（モニタリング）
                if result["results"]:
                    print("     Top results:")
                    for j, res in enumerate(result["results"][:3], 1):
                        file_name = Path(res["file_path"]).name
                        print(f"       {j}. {file_name} (score: {res['score']:.4f})")
            else:
                print(f"  ❌ Failed: {result.get('error', 'Unknown error')}")
            
            print()
        
        print(f"{'='*60}")
        print(f"✨ Benchmark completed!")
        print(f"{'='*60}\n")
    
    def save_results_json(self, output_dir: str = "summary/search"):
        """結果をJSON形式で保存"""
        # ディレクトリを作成
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # タイムスタンプ付きファイル名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        timestamped_filename = f"search_results_{timestamp}.json"
        filepath = output_path / timestamped_filename
        
        timestamp_iso = datetime.now().isoformat()
        
        output_data = {
            "timestamp": timestamp_iso,
            "total_queries": len(self.results),
            "results": self.results
        }
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        
        print(f"✓ Results saved to: {filepath}")
        return filepath
    
    def save_results_csv(self, output_dir: str = "summary/search"):
        """結果をCSV形式で保存（サマリー）"""
        # ディレクトリを作成
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # タイムスタンプ付きファイル名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        timestamped_filename = f"search_summary_{timestamp}.csv"
        filepath = output_path / timestamped_filename
        
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                'No.', 'Query', 'Mode', 'Rerank', 'Success', 'Response Time (ms)', 
                'Result Count', 'Top 1 File', 'Top 1 Score', 'Error'
            ])
            
            # データ行
            for i, result in enumerate(self.results, 1):
                top_file = result["results"][0]["file_path"] if result["results"] else ""
                top_score = result["results"][0]["score"] if result["results"] else 0
                
                writer.writerow([
                    i,
                    result["query"],
                    result["mode"],
                    "✅" if result.get("use_rerank", False) else "❌",
                    "✅" if result["success"] else "❌",
                    f"{result['response_time_ms']:.2f}",
                    result["result_count"],
                    top_file,
                    f"{top_score:.4f}",
                    result.get("error", "")
                ])
        
        print(f"✓ Results saved to: {filepath}")
        return filepath
    
    def save_detailed_results_csv(self, output_dir: str = "summary/search"):
        """結果をCSV形式で保存（詳細版：全検索結果）"""
        # ディレクトリを作成
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # タイムスタンプ付きファイル名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        timestamped_filename = f"search_detailed_{timestamp}.csv"
        filepath = output_path / timestamped_filename
        
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                'Query No.', 'Query', 'Mode', 'Rerank', 'Response Time (ms)', 
                'Result Rank', 'File Path', 'Score', 'Content Preview'
            ])
            
            # データ行
            for i, result in enumerate(self.results, 1):
                if result["results"]:
                    for res in result["results"]:
                        writer.writerow([
                            i,
                            result["query"],
                            result["mode"],
                            "✅" if result.get("use_rerank", False) else "❌",
                            f"{result['response_time_ms']:.2f}",
                            res["rank"],
                            res["file_path"],
                            f"{res['score']:.4f}",
                            res["content_preview"]
                        ])
                else:
                    # 結果がない場合も行を出力
                    writer.writerow([
                        i,
                        result["query"],
                        result["mode"],
                        "✅" if result.get("use_rerank", False) else "❌",
                        f"{result['response_time_ms']:.2f}",
                        0,
                        "",
                        "0.0000",
                        result.get("error", "No results")
                    ])
        
        print(f"✓ Results saved to: {filepath}")
        return filepath
    
    def save_summary_markdown(self, output_dir: str = "summary/search"):
        """サマリーをMarkdown形式で保存（詳細版）"""
        # ディレクトリを作成
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # タイムスタンプ付きファイル名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"search_summary_{timestamp}.md"
        filepath = output_path / filename
        
        # Markdown内容を生成
        md_content = []
        md_content.append("# Search Benchmark Report (Detailed)\n")
        md_content.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        md_content.append("---\n")
        
        # 統計情報
        successful = [r for r in self.results if r["success"]]
        failed = [r for r in self.results if not r["success"]]
        
        md_content.append("## 📊 Summary Statistics\n")
        md_content.append(f"- **Total Queries:** {len(self.results)}\n")
        md_content.append(f"- **Successful:** {len(successful)}\n")
        md_content.append(f"- **Failed:** {len(failed)}\n")
        
        if successful:
            response_times = [r["response_time_ms"] for r in successful]
            result_counts = [r["result_count"] for r in successful]
            
            avg_response_time = sum(response_times) / len(response_times)
            min_response_time = min(response_times)
            max_response_time = max(response_times)
            avg_result_count = sum(result_counts) / len(result_counts)
            
            md_content.append("\n### ⏱️ Response Time Statistics\n")
            md_content.append(f"- **Average:** {avg_response_time:.2f} ms\n")
            md_content.append(f"- **Min:** {min_response_time:.2f} ms\n")
            md_content.append(f"- **Max:** {max_response_time:.2f} ms\n")
            
            md_content.append("\n### 📋 Results Statistics\n")
            md_content.append(f"- **Average Results per Query:** {avg_result_count:.2f}\n")
        
        # 失敗したクエリ
        if failed:
            md_content.append("\n## ⚠️ Failed Queries\n")
            for result in failed:
                md_content.append(f"- **Query:** `{result['query']}`\n")
                md_content.append(f"  - **Error:** {result.get('error', 'Unknown error')}\n")
        
        # 個別のクエリ結果（詳細版）
        md_content.append("\n## 🔍 Detailed Query Results\n")
        for i, result in enumerate(self.results, 1):
            if result["success"] and result["results"]:
                md_content.append(f"\n### Query {i}: `{result['query']}`\n\n")
                md_content.append(f"**Mode:** {result['mode']} | **Rerank:** {'✅' if result.get('use_rerank', False) else '❌'}\n")
                md_content.append(f"**Response Time:** {result['response_time_ms']:.2f} ms | **Total Results:** {result['result_count']}\n\n")
                
                # クエリ情報（あれば）
                if result.get("query_info"):
                    query_info = result["query_info"]
                    md_content.append("**Query Analysis:**\n")
                    if "expanded_terms" in query_info:
                        md_content.append(f"- Expanded Terms: `{', '.join(query_info['expanded_terms'])}`\n")
                    if "token_count" in query_info:
                        md_content.append(f"- Token Count: {query_info['token_count']}\n")
                    md_content.append("\n")
                
                # 検索結果の詳細
                md_content.append("#### 📄 Search Results\n\n")
                
                for j, res in enumerate(result["results"], 1):
                    md_content.append(f"##### Result #{j}\n\n")
                    
                    # ファイル名（パスから抽出）
                    file_path = res.get('file_path', 'Unknown')
                    md_content.append(f"**Source File:** `{file_path}`\n")
                    md_content.append(f"**Score:** {res['score']:.4f}\n\n")
                    
                    # スコア詳細（hybrid検索の場合）
                    if "scores" in res and res["scores"]:
                        scores = res["scores"]
                        md_content.append("**Score Breakdown:**\n")
                        if "dense" in scores:
                            md_content.append(f"- 🔵 Dense (Semantic): {scores['dense']:.4f}\n")
                        if "sparse" in scores:
                            md_content.append(f"- 🔤 Sparse (Keyword): {scores['sparse']:.4f}\n")
                        if "rerank" in scores:
                            md_content.append(f"- ⭐ Rerank: {scores['rerank']:.4f}\n")
                        if "rrf" in scores:
                            md_content.append(f"- 🔀 RRF: {scores['rrf']:.4f}\n")
                        md_content.append("\n")
                    
                    # メタデータ（詳細版）
                    if "metadata" in res and res["metadata"]:
                        metadata = res["metadata"]
                        md_content.append("**Metadata:**\n")
                        
                        # 重要なメタデータを順番に表示
                        metadata_keys = [
                            "source", "file_path", "file_name", "file_type", "domain",
                            "document_id", "file_hash", 
                            "created_timestamp", "indexed_timestamp", "file_modified_timestamp",
                            "h1", "h2", "h3", "h4", "h5", "h6",
                            "chunk_index", "chunk_size", "total_chunks",
                            "score", "dense_score", "sparse_score", "rrf_score", "rerank_score"
                        ]
                        
                        for key in metadata_keys:
                            if key in metadata:
                                value = metadata[key]
                                md_content.append(f"- `{key}`: {value}\n")
                        
                        # その他のメタデータも表示（上記以外）
                        other_keys = [k for k in metadata.keys() if k not in metadata_keys 
                                     and k not in ["embedding", "page_content"]]
                        for key in sorted(other_keys):
                            value = metadata[key]
                            # 長すぎる値は省略
                            if isinstance(value, str) and len(value) > 100:
                                value = value[:100] + "..."
                            md_content.append(f"- `{key}`: {value}\n")
                        
                        md_content.append("\n")
                    
                    # チャンクの内容（全文）
                    if "content" in res and res["content"]:
                        content = res["content"]
                        md_content.append("**Chunk Content:**\n")
                        md_content.append("```\n")
                        md_content.append(content)
                        md_content.append("\n```\n")
                    
                    md_content.append("\n---\n\n")
        
        # ファイルに書き込み
        with open(filepath, 'w', encoding='utf-8') as f:
            f.writelines(md_content)
        
        print(f"✓ Summary saved to: {filepath}")
        return filepath
    
    def print_statistics(self):
        """統計情報を表示"""
        if not self.results:
            print("No results to analyze")
            return
        
        successful = [r for r in self.results if r["success"]]
        failed = [r for r in self.results if not r["success"]]
        
        if successful:
            response_times = [r["response_time_ms"] for r in successful]
            result_counts = [r["result_count"] for r in successful]
            
            avg_response_time = sum(response_times) / len(response_times)
            min_response_time = min(response_times)
            max_response_time = max(response_times)
            avg_result_count = sum(result_counts) / len(result_counts)
        else:
            avg_response_time = 0
            min_response_time = 0
            max_response_time = 0
            avg_result_count = 0
        
        print(f"\n{'='*60}")
        print(f"📊 Benchmark Statistics")
        print(f"{'='*60}")
        print(f"Total Queries:       {len(self.results)}")
        print(f"Successful:          {len(successful)}")
        print(f"Failed:              {len(failed)}")
        print(f"\n--- Response Time ---")
        print(f"Average:             {avg_response_time:.2f} ms")
        print(f"Min:                 {min_response_time:.2f} ms")
        print(f"Max:                 {max_response_time:.2f} ms")
        print(f"\n--- Results ---")
        print(f"Avg Results/Query:   {avg_result_count:.2f}")
        print(f"{'='*60}\n")
    
    def disconnect(self):
        """MCPサーバから切断"""
        if self.mcp_service.connection_status["connected"]:
            result = self.mcp_service.disconnect()
            print(f"\n{result}")
        
        self.mcp_service.stop_worker()


def generate_sample_queries() -> List[str]:
    """サンプルクエリを生成（100個）"""
    queries = [
        # 基本的な機能関連
        "入力チェック",
        "バリデーション",
        "データベース接続",
        "テーブル定義",
        "エラー処理",
        "例外ハンドリング",
        "ログ出力",
        "API エンドポイント",
        "認証機能",
        "権限管理",
        
        # UI/画面関連
        "画面遷移",
        "ボタン配置",
        "フォーム入力",
        "画面レイアウト",
        "ユーザーインターフェース",
        "メニュー構成",
        "ダイアログ表示",
        "モーダル画面",
        "確認メッセージ",
        "エラーメッセージ",
        
        # データ処理関連
        "データ変換",
        "CSV インポート",
        "Excel エクスポート",
        "ファイルアップロード",
        "ファイルダウンロード",
        "一括処理",
        "バッチ処理",
        "非同期処理",
        "トランザクション",
        "ロールバック",
        
        # セキュリティ関連
        "暗号化",
        "パスワード管理",
        "セッション管理",
        "トークン認証",
        "アクセス制御",
        "SQLインジェクション対策",
        "XSS対策",
        "CSRF対策",
        "セキュリティ要件",
        "個人情報保護",
        
        # パフォーマンス関連
        "性能要件",
        "レスポンス時間",
        "同時接続数",
        "キャッシュ",
        "インデックス",
        "クエリ最適化",
        "負荷分散",
        "スケーラビリティ",
        "メモリ管理",
        "リソース管理",
        
        # テスト関連
        "単体テスト",
        "結合テスト",
        "システムテスト",
        "受入テスト",
        "回帰テスト",
        "負荷テスト",
        "セキュリティテスト",
        "テストケース",
        "テストデータ",
        "テスト環境",
        
        # 運用関連
        "バックアップ",
        "リストア",
        "監視項目",
        "アラート通知",
        "障害対応",
        "運用手順",
        "メンテナンス",
        "バージョン管理",
        "リリース手順",
        "デプロイ",
        
        # 業務フロー関連
        "申請処理",
        "承認フロー",
        "ワークフロー",
        "通知機能",
        "メール送信",
        "レポート出力",
        "集計処理",
        "検索機能",
        "ソート機能",
        "フィルタ機能",
        
        # システム構成関連
        "システム構成図",
        "ネットワーク構成",
        "サーバー構成",
        "データベース構成",
        "アーキテクチャ",
        "インフラ要件",
        "冗長化構成",
        "DR対策",
        "バックアップ構成",
        "セキュリティ構成",
        
        # 追加の技術的クエリ
        "REST API",
        "JSON フォーマット",
        "XML パース",
        "WebSocket 通信",
        "HTTP ステータスコード",
    ]
    
    return queries[:100]  # 最初の100個を返す


def main():
    """
    ここで検索クエリのリストと設定を定義します
    """
    
    # ========================================
    # 検索クエリリストを定義（何個でもOK）
    # ========================================
    queries = [
        # ここに検索したいクエリを追加
        "MBD",
    ]
    
    # クエリが空の場合はサンプルクエリを使用
    if not queries:
        print("⚠️  No queries defined. Using sample queries...")
        queries = generate_sample_queries()
    
    # ========================================
    # 設定
    # ========================================
    LIMIT = 5  # 取得する検索結果の件数
    MODE = "hybrid"  # 検索モード: "semantic", "keyword", "hybrid"
    USE_HTTP = False  # HTTPモードを使うか
    HTTP_URL = "http://localhost:8000/mcp"
    USE_RERANK = False  # リランクを使用するか（config.pyのreranker.enabledに依存）
    
    # ========================================
    # ベンチマーク実行
    # ========================================
    print(f"📝 Total queries to execute: {len(queries)}")
    print(f"⚙️  Settings: mode={MODE}, limit={LIMIT}, rerank={USE_RERANK}")
    
    benchmark = SearchBenchmark()
    
    try:
        # MCPサーバに接続
        if USE_HTTP:
            if not benchmark.connect(mode="http", http_url=HTTP_URL):
                return
        else:
            if not benchmark.connect(mode="stdio"):
                return
        
        # ベンチマーク実行（ターミナルモニタリング付き）
        benchmark.run_benchmark(queries, LIMIT, MODE, USE_RERANK)
        
        # 統計情報を表示
        benchmark.print_statistics()
        
        # 結果を保存（Markdownのみ）
        md_path = benchmark.save_summary_markdown()
        
        # 以下はコメントアウト（必要に応じて有効化してください）
        # json_path = benchmark.save_results_json()
        # csv_path = benchmark.save_results_csv()
        # detail_path = benchmark.save_detailed_results_csv()
        
    finally:
        # 切断
        benchmark.disconnect()
    
    print("\n✨ All done!")
    print(f"\nOutput file: {md_path}")
    print("\n💡 Note: JSON/CSV output is disabled by default.")
    print("   To enable, uncomment the relevant lines in main()")


if __name__ == "__main__":
    main()