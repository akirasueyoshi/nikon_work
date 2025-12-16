"""
EurekaBoxコンテンツ用のアップロードプログラム

- data_source/ からファイルを読み込み、upload_document API でアップロード
- MCPプロトコル経由ではなく、MCPサーバインスタンスを直接起動して使用
- Markdownファイル内に参照されている画像も自動的にアップロード
- data_source/ の直下のフォルダをドメインとして扱う
"""

import sys
import time
import asyncio
from pathlib import Path

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config import get_config, KnowledgeBaseConfig
from src.protocol.server import KnowledgeBaseMCPServer
from src.protocol.schemas import DeleteDomainInput
from src.utils.file_uploader_direct import DirectFileUploader

BASE_DIR = Path(__file__).resolve().parent
DATA_SOURCE_DIR = BASE_DIR.parent / "data_source"


async def init_index(
    config: KnowledgeBaseConfig | None = None,
    force_recreate: bool = False,
    data_source_path: Path | None = None,
    upload_images: bool = True,
):
    """
    初期インデックスを作成
    data_source/ からファイルを読み込み、upload_document API でアップロード

    Args:
        config: 設定（Noneの場合はデフォルト）
        force_recreate: 既存インデックスを削除して再作成
        data_source_path: データソースのパス
        upload_images: Markdownから参照される画像も自動アップロード
    """
    config = config or get_config()
    data_source_path = data_source_path or DATA_SOURCE_DIR

    print("=" * 80)
    print("Knowledge Base - Index Initialization (with Images)")
    print("=" * 80)
    print(f"Data Source: {data_source_path}")
    print(f"Storage Path: {config.upload_files_path}")
    print(f"Force Recreate: {force_recreate}")
    print(f"Upload Images: {upload_images}")
    print("=" * 80)
    print()

    # データソースの存在チェック
    if not data_source_path.exists():
        print(f"❌ Data source not found: {data_source_path}")
        print("   Create a data source directory with domain subdirectories")
        print(f"   Example: mkdir -p {data_source_path}/USDM")
        print("=" * 80)
        return

    # ドメイン検出
    domains = sorted(
        [
            d.name
            for d in data_source_path.iterdir()
            if d.is_dir() and not d.name.startswith(".")
        ]
    )

    if not domains:
        print(f"⚠️  No domains found in {data_source_path}")
        print("   Create domain directories (e.g., USDM, MBD)")
        print("=" * 80)
        return

    print(f"Detected Domains: {', '.join(domains)}")
    print()

    # サーバーインスタンス作成
    server = KnowledgeBaseMCPServer(config)
    server.initialize()

    try:
        # 既存ドメインの削除（force_recreateの場合）
        if force_recreate:
            print("⚠️  Force recreate mode: deleting existing domains...")
            for domain in domains:
                try:
                    result = server.document_service.delete_domain(
                        DeleteDomainInput(
                            domain_name=domain,
                            confirm=True,
                            delete_vectors=True,
                        )
                    )
                    if result.status == "deleted":
                        print(f"   ✅ Deleted domain: {domain}")
                    elif result.status == "not_found":
                        print(f"   ℹ️  Domain not found (will be created): {domain}")
                except Exception as e:
                    print(f"   ⚠️  Error deleting domain '{domain}': {e}")
            print()

        # DirectFileUploader インスタンスを作成（Markdownファイルのみ処理）
        uploader = DirectFileUploader(server, supported_extensions={".md"})

        # ディレクトリ全体をアップロード
        print("📤 Uploading files...")
        print("-" * 80)

        start_time = time.time()
        results = await uploader.upload_directory(
            base_dir=data_source_path,
            upload_images=upload_images,
            auto_create_domains=True,
            overwrite=True,
        )
        total_time = time.time() - start_time

        # 全体サマリー
        uploaded = sum(1 for r in results if r.status == "success")
        failed = sum(1 for r in results if r.status == "failed")
        total_chunks = sum(r.chunks for r in results)

        print()
        print("=" * 80)
        print("📊 Index Initialization Summary")
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
        print(f"Total Execution Time: {total_time:.2f}s")
        print(f"Total Files: {uploaded}")
        print(f"Total Chunks: {total_chunks}")
        if failed > 0:
            print(f"Total Errors: {failed}")
        print("=" * 80)

    finally:
        # クリーンアップ
        server.close()


async def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(
        description="Initialize Knowledge Base index from data_source/"
    )
    parser.add_argument(
        "--force-recreate",
        action="store_true",
        help="Force recreate index (delete existing)",
    )
    parser.add_argument(
        "--data-source",
        type=str,
        default=None,
        help="Data source directory path",
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="Skip automatic image upload",
    )

    args = parser.parse_args()

    # 設定の取得
    config = get_config()
    data_source_path = Path(args.data_source) if args.data_source else DATA_SOURCE_DIR

    # インデックス作成
    try:
        await init_index(
            config=config,
            force_recreate=args.force_recreate,
            data_source_path=data_source_path,
            upload_images=not args.no_images,
        )
    except KeyboardInterrupt:
        print("\n⚠️  Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())
