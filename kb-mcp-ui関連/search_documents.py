"""
kb-mcp ドキュメント検索スクリプト

ナレッジベースを検索し、結果をファイルに保存します。
各クエリごとに個別のパラメータを設定できます。

実行方法:
    uv run search_documents.py

使い方:
    1. SEARCH_QUERIES リストにクエリと設定を追加
    2. スクリプトを実行
    3. search_result/ フォルダに結果が保存されます
"""

import asyncio
import json
from pathlib import Path
from datetime import datetime
from typing import Any

from src.utils.platform_utils import configure_console_encoding
from src.utils.mcp_client_wrapper import McpClientWrapper

configure_console_encoding()


# =============================================================================
# 設定
# =============================================================================

SERVER_COMMAND = "uv"
SERVER_ARGS = ["run", "kb-mcp-server", "--transport", "stdio"]

# 結果保存先ディレクトリ
OUTPUT_DIR = Path("search_result")


# =============================================================================
# 検索クエリの設定
# =============================================================================

# 各クエリの設定を辞書形式で定義
# "query": 必須、検索クエリ文字列
# その他のキーはオプション。指定しない場合はDEFAULT_PARAMSの値を使用
#
# パラメータの上書き・加算ルール:
# - mode, limit, scope_paths等: 指定すると上書きされる
# - exclude_paths: デフォルト + 個別指定が結合される（加算）
#
# 例:
# DEFAULT_PARAMS で exclude_paths = ["tmp/", "test/"]
# クエリで exclude_paths = ["archive/"]
# → 実際の除外パス: ["tmp/", "test/", "archive/"]
#
SEARCH_QUERIES = [
    {
        "query": "test.mdの内容は？",
        # パラメータを指定しない場合はデフォルト値を使用
    },
    {
        "query": "出身地はどこですか？",
        "mode": "semantic",  # デフォルトのmodeを上書き
        "limit": 5,          # デフォルトのlimitを上書き
    },
    {
        "query": "趣味は何ですか？",
        "scope_paths": ["files/"],     # デフォルトのscope_pathsを上書き
        "mode": "keyword",              # デフォルトのmodeを上書き
        "limit": 3,                     # デフォルトのlimitを上書き
        "exclude_paths": ["files/old/"], # デフォルトのexclude_pathsに追加
    },
]

# デフォルト検索パラメータ（個別に指定しない場合に使用）
# 
# 各クエリで同じキーを指定した場合:
# - exclude_paths: デフォルト + 個別指定が結合される（加算）
# - その他: 個別指定で上書きされる
#
DEFAULT_PARAMS = {
    # 検索対象のパス
    "scope_paths": None,  # Noneで全体検索、["files/"]でfilesフォルダのみ
    
    # 除外するパス（個別指定があれば加算される）
    "exclude_paths": None,  # 例: ["tmp/", "test/"]
    
    # メタデータフィルタ
    "metadata_filter": None,
    
    # 検索モード: "hybrid", "semantic", "keyword"
    "mode": "hybrid",
    
    # 返すドキュメント数
    "limit": 10,
    
    # リランキング使用
    "use_rerank": False,
    
    # スコア閾値
    "score_thresholds": None,
    
    # Greedyモード
    "greedy": False,
}


# =============================================================================
# ヘルパー関数
# =============================================================================


def parse_tool_result(result: Any) -> dict[str, Any]:
    """MCPツール実行結果からJSONデータを抽出"""
    # CallToolResult オブジェクトの場合
    if hasattr(result, 'content'):
        content_list = result.content
    # 辞書の場合
    elif isinstance(result, dict) and "content" in result:
        content_list = result["content"]
    else:
        return {}
    
    for content in content_list:
        # オブジェクトの場合
        if hasattr(content, 'type') and hasattr(content, 'text'):
            if content.type == "text":
                try:
                    return json.loads(content.text)
                except json.JSONDecodeError:
                    return {}
        # 辞書の場合
        elif isinstance(content, dict) and content.get("type") == "text":
            try:
                return json.loads(content.get("text", "{}"))
            except json.JSONDecodeError:
                return {}
    
    return {}


def create_output_directory() -> Path:
    """結果保存用ディレクトリを作成"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    return OUTPUT_DIR


def save_search_result(query: str, data: dict[str, Any], query_index: int, params: dict[str, Any]) -> Path:
    """検索結果を複数形式で保存
    
    Args:
        query: 検索クエリ
        data: 検索結果データ
        query_index: クエリのインデックス（1始まり）
        params: 使用した検索パラメータ
    
    Returns:
        保存先ディレクトリのパス
    """
    # タイムスタンプ付きディレクトリ名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_dir = OUTPUT_DIR / f"{timestamp}_query{query_index}"
    result_dir.mkdir(parents=True, exist_ok=True)
    
    # 1. JSON形式で保存（機械可読）
    json_path = result_dir / "result.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    # 2. テキスト形式で保存（人間可読）
    txt_path = result_dir / "result.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("=" * 80 + "\n")
        f.write(f"検索クエリ: {query}\n")
        f.write("=" * 80 + "\n\n")
        
        documents = data.get("documents", [])
        query_info = data.get("query_info", {})
        
        f.write(f"検索モード: {query_info.get('mode', 'N/A')}\n")
        f.write(f"実行時間: {query_info.get('took_ms', 0):.1f}ms\n")
        f.write(f"ヒット件数: {len(documents)} ドキュメント\n\n")
        
        if len(documents) == 0:
            f.write("⚠ ドキュメントが見つかりませんでした\n")
        else:
            for i, doc in enumerate(documents, 1):
                f.write("─" * 80 + "\n")
                f.write(f"[{i}] {doc.get('file_path', 'N/A')}\n")
                f.write("─" * 80 + "\n")
                f.write(f"タイトル: {doc.get('title', 'N/A')}\n")
                f.write(f"ドキュメントID: {doc.get('document_id', 'N/A')}\n")
                f.write(f"スコア: {doc.get('max_score', 0):.4f}\n")
                f.write(f"総チャンク数: {doc.get('chunk_count', 0)}\n\n")
                
                # 代表チャンク
                chunks = doc.get("representative_chunks", [])
                if chunks:
                    f.write(f"代表チャンク ({len(chunks)}件):\n\n")
                    for j, chunk in enumerate(chunks, 1):
                        f.write(f"  [{j}] スコア: {chunk.get('score', 0):.4f}\n")
                        text = chunk.get("text", "")
                        f.write(f"  テキスト:\n")
                        for line in text.split("\n"):
                            f.write(f"    {line}\n")
                        f.write("\n")
                
                f.write("\n")
    
    # 3. Markdown形式で保存（ドキュメント用）
    md_path = result_dir / "result.md"
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(f"# 検索結果: {query}\n\n")
        
        documents = data.get("documents", [])
        query_info = data.get("query_info", {})
        
        f.write("## 検索情報\n\n")
        f.write(f"- **検索モード**: {query_info.get('mode', 'N/A')}\n")
        f.write(f"- **実行時間**: {query_info.get('took_ms', 0):.1f}ms\n")
        f.write(f"- **ヒット件数**: {len(documents)} ドキュメント\n\n")
        
        if len(documents) == 0:
            f.write("⚠ ドキュメントが見つかりませんでした\n")
        else:
            f.write("## 検索結果\n\n")
            for i, doc in enumerate(documents, 1):
                f.write(f"### {i}. {doc.get('file_path', 'N/A')}\n\n")
                f.write(f"- **タイトル**: {doc.get('title', 'N/A')}\n")
                f.write(f"- **スコア**: {doc.get('max_score', 0):.4f}\n")
                f.write(f"- **チャンク数**: {doc.get('chunk_count', 0)}\n")
                f.write(f"- **ドキュメントID**: `{doc.get('document_id', 'N/A')}`\n\n")
                
                # 代表チャンク
                chunks = doc.get("representative_chunks", [])
                if chunks:
                    f.write(f"#### 代表チャンク\n\n")
                    for j, chunk in enumerate(chunks, 1):
                        f.write(f"**チャンク {j}** (スコア: {chunk.get('score', 0):.4f}):\n\n")
                        text = chunk.get("text", "")
                        f.write("```\n")
                        f.write(text)
                        f.write("\n```\n\n")
    
    # 4. クエリ情報を保存（使用したパラメータも含む）
    query_info_path = result_dir / "query.txt"
    with open(query_info_path, "w", encoding="utf-8") as f:
        f.write(f"クエリ: {query}\n")
        f.write(f"実行日時: {timestamp}\n\n")
        f.write(f"使用したパラメータ:\n")
        f.write(json.dumps(params, ensure_ascii=False, indent=2))
    
    return result_dir


def print_summary(query: str, data: dict[str, Any], result_dir: Path, params: dict[str, Any]) -> None:
    """検索結果のサマリーをコンソールに表示"""
    documents = data.get("documents", [])
    query_info = data.get("query_info", {})
    
    print(f"\n検索: {query}")
    print(f"  モード: {query_info.get('mode', 'N/A')}")
    print(f"  limit: {params.get('limit', 'N/A')}")
    
    # カスタムパラメータがある場合は表示
    custom_params = []
    if params.get('scope_paths'):
        custom_params.append(f"scope_paths={params['scope_paths']}")
    if params.get('exclude_paths'):
        custom_params.append(f"exclude_paths={params['exclude_paths']}")
    if params.get('use_rerank'):
        custom_params.append("use_rerank=True")
    if params.get('greedy'):
        custom_params.append("greedy=True")
    
    if custom_params:
        print(f"  カスタム設定: {', '.join(custom_params)}")
    
    print(f"  実行時間: {query_info.get('took_ms', 0):.1f}ms")
    print(f"  ヒット件数: {len(documents)} ドキュメント")
    
    if len(documents) > 0:
        print(f"  結果を保存: {result_dir}")
        print(f"    - result.json (JSON形式)")
        print(f"    - result.txt (テキスト形式)")
        print(f"    - result.md (Markdown形式)")
        print(f"    - query.txt (クエリ情報)")
    else:
        print("  ⚠ ドキュメントが見つかりませんでした")


# =============================================================================
# メイン処理
# =============================================================================


async def search_documents(client: McpClientWrapper, query_config: dict[str, Any], query_index: int) -> None:
    """ドキュメントを検索して結果を保存
    
    Args:
        client: MCPクライアント
        query_config: クエリ設定（query + オプションパラメータ）
        query_index: クエリのインデックス（1始まり）
    """
    # queryを取得（必須）
    query = query_config.get("query")
    if not query:
        print(f"\n❌ エラー: クエリが指定されていません (query_index: {query_index})")
        return
    
    # デフォルトパラメータをコピー
    params = DEFAULT_PARAMS.copy()
    
    # exclude_pathsは特別扱い（加算する）
    default_exclude_paths = DEFAULT_PARAMS.get("exclude_paths") or []
    query_exclude_paths = query_config.get("exclude_paths") or []
    
    # クエリ個別のパラメータで上書き
    for key, value in query_config.items():
        if key != "query" and key != "exclude_paths":  # queryとexclude_pathsは除外
            params[key] = value
    
    # exclude_pathsを結合（重複を除去）
    if default_exclude_paths or query_exclude_paths:
        combined_exclude_paths = list(set(default_exclude_paths + query_exclude_paths))
        params["exclude_paths"] = combined_exclude_paths if combined_exclude_paths else None
    
    # queryを追加
    params["query"] = query
    
    try:
        # 検索実行
        result = await client.call_tool(
            name="search_documents",
            arguments=params,
            timeout=60,
        )
        
        # 結果を解析
        data = parse_tool_result(result)
        
        if "error" in data:
            print(f"\n❌ エラー: {data['error']}")
            if "traceback" in data:
                print(f"トレースバック:\n{data['traceback']}")
            return
        
        # 結果を保存
        result_dir = save_search_result(query, data, query_index, params)
        
        # サマリーを表示
        print_summary(query, data, result_dir, params)
        
    except Exception as e:
        print(f"\n❌ 検索エラー ({query}): {e}")
        import traceback
        traceback.print_exc()


async def main() -> None:
    """メイン処理"""
    print("=" * 80)
    print("kb-mcp ドキュメント検索")
    print("=" * 80)
    
    # 出力ディレクトリを作成
    create_output_directory()
    print(f"\n結果保存先: {OUTPUT_DIR.absolute()}")
    
    client = McpClientWrapper()
    
    try:
        # サーバに接続
        print("\n[接続] サーバに接続中...")
        await client.connect_stdio(command=SERVER_COMMAND, args=SERVER_ARGS)
        print("✓ 接続完了")
        
        # 各クエリで検索を実行
        print(f"\n[検索] {len(SEARCH_QUERIES)} 件のクエリを実行")
        print("─" * 80)
        
        for i, query_config in enumerate(SEARCH_QUERIES, 1):
            await search_documents(client, query_config, i)
            
            # 次の検索まで少し待機
            if i < len(SEARCH_QUERIES):
                await asyncio.sleep(0.5)
        
        print("\n" + "─" * 80)
        print(f"✓ 全ての検索が完了しました")
        print(f"結果: {OUTPUT_DIR.absolute()}")
        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # 接続を切断
        print("\n[切断] 接続を切断中...")
        await client.disconnect()
        print("✓ 切断完了")
    
    print("\n" + "=" * 80)
    print("完了")
    print("=" * 80)


if __name__ == "__main__":
    asyncio.run(main())