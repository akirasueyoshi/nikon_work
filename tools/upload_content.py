"""
汎用的なコンテンツアップロードプログラム

機能:
- 指定ディレクトリ配下を再帰的に探索
- top直下のディレクトリをドメインとして認識
- ファイルパスを相対パスで保持
- サポートされているファイル形式のみアップロード
- MCPプロトコル経由ではなく、MCPサーバインスタンスを直接起動して使用
"""

import sys
import time
import asyncio
from pathlib import Path

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config import get_config
from src.protocol.server import KnowledgeBaseMCPServer
from src.utils.file_uploader_direct import DirectFileUploader


async def main(top_directory: Path):
    """
    メイン処理

    Args:
        top_directory: トップディレクトリ（この直下がドメイン）
    """
    print("\n" + "=" * 40)
    print("コンテンツアップロードプログラム")
    print("=" * 40)

    # バリデーション
    if not top_directory.exists():
        print(f"\n❌ ディレクトリが存在しません: {top_directory}")
        return

    if not top_directory.is_dir():
        print(f"\n❌ ディレクトリではありません: {top_directory}")
        return

    # サーバー初期化
    config = get_config()
    server = KnowledgeBaseMCPServer(config)
    server.initialize()

    try:
        # DirectFileUploader インスタンスを作成
        uploader = DirectFileUploader(server)

        print(f"\n📂 Scanning directory: {top_directory.absolute()}")
        print("-" * 80)

        # 進捗コールバック（オプション）
        def on_progress(domain, file_path, results):
            has_error = any(r.status == "failed" for r in results)
            if has_error:
                print(f"  ❌ {domain}/{file_path.name}")
                for r in results:
                    if r.status == "failed":
                        print(f"     - {r.message}")

        # ディレクトリ全体をアップロード
        start_time = time.time()
        results = await uploader.upload_directory(
            base_dir=top_directory,
            upload_images=True,
            auto_create_domains=True,
            overwrite=True,
            progress_callback=on_progress,
        )
        total_time = time.time() - start_time

        # サマリー表示
        uploaded = sum(1 for r in results if r.status == "success")
        failed = sum(1 for r in results if r.status == "failed")
        total_chunks = sum(r.chunks for r in results)

        print("\n" + "=" * 80)
        print("📊 アップロード統計詳細")
        print("=" * 80)
        print(f"{'File':<50} | {'Status':<10} | {'Time (s)':>10}")
        print("-" * 80)

        for r in results:
            # パスを短く表示するために相対パス取得を試みる
            try:
                display_path = Path(r.file_path).name
                parent = Path(r.file_path).parent.name
                if parent:
                    display_path = f"{parent}/{display_path}"
            except Exception:
                display_path = Path(r.file_path).name

            # 長すぎる場合は切り詰める
            if len(display_path) > 48:
                display_path = "..." + display_path[-45:]

            status_icon = "✅" if r.status == "success" else "❌"
            print(
                f"{status_icon} {display_path:<47} | {r.status:<10} | {r.elapsed:>10.2f}"
            )

        print("-" * 80)
        print(f"  Total Execution Time: {total_time:.2f}s")
        print(f"  Total items processed: {len(results):>6}")
        print(f"  Successfully uploaded: {uploaded:>6}")
        print(f"  Failed:               {failed:>6}")
        print(f"  Total chunks created: {total_chunks:>6}")
        print("=" * 80)

    finally:
        # クリーンアップ処理
        server.close()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="再帰的にファイルをアップロード（top直下をドメインとして認識）"
    )
    parser.add_argument(
        "directory",
        type=Path,
        nargs="?",
        default=Path("upload_files"),
        help="トップディレクトリのパス（デフォルト: upload_files）",
    )

    args = parser.parse_args()

    asyncio.run(main(args.directory))
