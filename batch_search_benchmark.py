#!/usr/bin/env python3
"""
複数の検索クエリを順次実行して結果を記録するスクリプト

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
        from ui.services.mcp_client import MCPClientManager
        from ui.services.search_service import SearchService
        
        self.client_manager = MCPClientManager()
        self.search_service = SearchService(self.client_manager)
        
        self.results = []
        
        print("✓ Initialized MCP client and search service")
    
    def connect(self, mode="stdio", command="uv", args="run src/main.py", http_url=None):
        """MCPサーバに接続"""
        print(f"\n🔌 Connecting to MCP server...")
        
        if mode == "http" and http_url:
            result = self.client_manager.connect(
                transport_type="Streamable HTTP",
                connection_mode="Manual Connection",
                command="",
                args="",
                url=http_url,
                host="",
                port=""
            )
        else:
            result = self.client_manager.connect(
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
    
    def search_single_query(self, query: str, limit: int = 5, mode: str = "hybrid") -> Dict[str, Any]:
        """単一クエリで検索を実行"""
        try:
            # 時間計測開始
            start_time = time.time()
            
            # 検索実行
            response = self.client_manager.send_request(
                "search",
                {
                    "query": query,
                    "limit": limit,
                    "mode": mode
                }
            )
            
            # 時間計測終了
            end_time = time.time()
            elapsed_time = (end_time - start_time) * 1000  # ミリ秒に変換
            
            # 結果の解析
            if response.get("success"):
                results = response.get("results", [])
                
                result_data = {
                    "query": query,
                    "mode": mode,
                    "limit": limit,
                    "success": True,
                    "response_time_ms": elapsed_time,
                    "result_count": len(results),
                    "results": []
                }
                
                # 各結果の詳細を記録
                for i, res in enumerate(results, 1):
                    result_data["results"].append({
                        "rank": i,
                        "file_path": res.get("path", ""),
                        "score": res.get("score", 0),
                        "content_preview": res.get("content", "")[:100] if res.get("content") else ""
                    })
                
                return result_data
            else:
                return {
                    "query": query,
                    "mode": mode,
                    "limit": limit,
                    "success": False,
                    "response_time_ms": elapsed_time,
                    "result_count": 0,
                    "error": response.get("error", "Unknown error"),
                    "results": []
                }
                
        except Exception as e:
            return {
                "query": query,
                "mode": mode,
                "limit": limit,
                "success": False,
                "response_time_ms": 0,
                "result_count": 0,
                "error": str(e),
                "results": []
            }
    
    def run_benchmark(self, queries: List[str], limit: int = 5, mode: str = "hybrid"):
        """複数クエリのベンチマークを実行"""
        total = len(queries)
        
        print(f"\n{'='*60}")
        print(f"🔍 Starting search benchmark: {total} quer{'y' if total == 1 else 'ies'}")
        print(f"   Mode: {mode}, Limit: {limit}")
        print(f"{'='*60}\n")
        
        for i, query in enumerate(queries, 1):
            print(f"[{i}/{total}] Searching: {query}")
            
            result = self.search_single_query(query, limit, mode)
            self.results.append(result)
            
            if result["success"]:
                print(f"  ✅ Found {result['result_count']} results in {result['response_time_ms']:.2f}ms")
            else:
                print(f"  ❌ Failed: {result.get('error', 'Unknown error')}")
            
            print()
        
        print(f"{'='*60}")
        print(f"✨ Benchmark completed!")
        print(f"{'='*60}\n")
    
    def save_results_json(self, output_file: str = "search_results.json"):
        """結果をJSON形式で保存"""
        timestamp = datetime.now().isoformat()
        
        output_data = {
            "timestamp": timestamp,
            "total_queries": len(self.results),
            "results": self.results
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        
        print(f"✓ Results saved to: {output_file}")
    
    def save_results_csv(self, output_file: str = "search_results.csv"):
        """結果をCSV形式で保存（サマリー）"""
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                'No.', 'Query', 'Mode', 'Success', 'Response Time (ms)', 
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
                    "✅" if result["success"] else "❌",
                    f"{result['response_time_ms']:.2f}",
                    result["result_count"],
                    top_file,
                    f"{top_score:.4f}",
                    result.get("error", "")
                ])
        
        print(f"✓ Results saved to: {output_file}")
    
    def save_detailed_results_csv(self, output_file: str = "search_results_detailed.csv"):
        """結果をCSV形式で保存（詳細版：全検索結果）"""
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                'Query No.', 'Query', 'Mode', 'Response Time (ms)', 
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
                            f"{result['response_time_ms']:.2f}",
                            res["rank"],
                            res["file_path"],
                            f"{res['score']:.4f}",
                            res["content_preview"]
                        ])
                else:
                    # 結果がない場合
                    writer.writerow([
                        i,
                        result["query"],
                        result["mode"],
                        f"{result['response_time_ms']:.2f}",
                        0,
                        "",
                        0,
                        result.get("error", "No results")
                    ])
        
        print(f"✓ Detailed results saved to: {output_file}")
    
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
        if self.client_manager.connection_status["connected"]:
            result = self.client_manager.disconnect()
            print(f"\n{result}")
        
        self.client_manager.stop_worker()


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
        "機能仕様書",
        "入力チェック",
        "データベース",
        "エラー処理",
        "API 仕様",
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
    
    # 出力ファイル名
    OUTPUT_JSON = "search_results.json"
    OUTPUT_CSV = "search_results_summary.csv"
    OUTPUT_CSV_DETAIL = "search_results_detailed.csv"
    
    # ========================================
    # ベンチマーク実行
    # ========================================
    print(f"📝 Total queries to execute: {len(queries)}")
    
    benchmark = SearchBenchmark()
    
    try:
        # MCPサーバに接続
        if USE_HTTP:
            if not benchmark.connect(mode="http", http_url=HTTP_URL):
                return
        else:
            if not benchmark.connect(mode="stdio"):
                return
        
        # ベンチマーク実行
        benchmark.run_benchmark(queries, LIMIT, MODE)
        
        # 統計情報を表示
        benchmark.print_statistics()
        
        # 結果を保存
        benchmark.save_results_json(OUTPUT_JSON)
        benchmark.save_results_csv(OUTPUT_CSV)
        benchmark.save_detailed_results_csv(OUTPUT_CSV_DETAIL)
        
    finally:
        # 切断
        benchmark.disconnect()
    
    print("\n✨ All done!")
    print(f"\nOutput files:")
    print(f"  - {OUTPUT_JSON} (JSON format, complete data)")
    print(f"  - {OUTPUT_CSV} (CSV format, summary)")
    print(f"  - {OUTPUT_CSV_DETAIL} (CSV format, all results)")


if __name__ == "__main__":
    main()
