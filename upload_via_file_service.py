#!/usr/bin/env python3
"""
kb-mcpのFileServiceとMCPClientManagerを直接使ってファイルをアップロードするスクリプト

Usage:
    # MCPサーバがStdioモードで起動している場合
    python upload_via_file_service.py --files file1.docx file2.xlsx --dest domain/path/
    
    # MCPサーバがHTTPモードで起動している場合
    python upload_via_file_service.py --files file1.docx --dest domain/file.docx --http-url http://localhost:8000/mcp

Requirements:
    - kb-mcp リポジトリが存在すること
    - MCPサーバが起動していない場合は自動起動します
"""

import argparse
import sys
from pathlib import Path
import os
import time

# kb-mcpのパスを追加
SCRIPT_DIR = Path(__file__).parent.resolve()


class SimpleUploader:
    """FileServiceを使った簡単なアップローダー"""
    
    def __init__(self, kb_mcp_root: Path):
        self.kb_mcp_root = kb_mcp_root
        sys.path.insert(0, str(kb_mcp_root))
        
        # MCPモジュールをインポート
        from ui.services.mcp_client import MCPClientManager
        from ui.services.file_service import FileService
        
        self.client_manager = MCPClientManager()
        self.file_service = FileService(self.client_manager)
        print("✓ Initialized MCP client and file service")
    
    def connect_stdio(self, command="uv", args="run src/main.py"):
        """Stdioモードで接続"""
        print(f"🔌 Connecting to MCP server (Stdio mode)...")
        print(f"   Command: {command} {args}")
        
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
            return True
        else:
            print(f"❌ {result}")
            return False
    
    def connect_http(self, url="http://localhost:8000/mcp"):
        """HTTPモードで接続"""
        print(f"🔌 Connecting to MCP server (HTTP mode)...")
        print(f"   URL: {url}")
        
        result = self.client_manager.connect(
            transport_type="Streamable HTTP",
            connection_mode="Manual Connection",
            command="",
            args="",
            url=url,
            host="",
            port=""
        )
        
        if "✅" in result:
            print(f"✓ {result}")
            return True
        else:
            print(f"❌ {result}")
            return False
    
    def upload_files(self, file_paths: list[Path], destination_path: str, overwrite: bool = True):
        """ファイルをアップロード"""
        print(f"\n🚀 Starting upload of {len(file_paths)} file(s)...")
        print(f"   Destination: {destination_path}")
        print(f"   Overwrite: {overwrite}")
        print()
        
        # FileService.upload_filesはジェネレータなので、ログを順次表示
        try:
            gen = self.file_service.upload_files(
                source="explorer",
                local_paths=[str(f) for f in file_paths],
                upload_path_str=destination_path,
                overwrite_flag=overwrite,
                progress_callback=None  # プログレスコールバックは不要
            )
            
            # ジェネレータからログを取得
            for log_chunk in gen:
                print(log_chunk)
            
        except Exception as e:
            print(f"\n❌ Upload failed: {e}")
            import traceback
            traceback.print_exc()
    
    def disconnect(self):
        """MCPサーバから切断"""
        if self.client_manager.connection_status["connected"]:
            result = self.client_manager.disconnect()
            print(f"\n{result}")
        
        # ワーカープロセスを停止
        self.client_manager.stop_worker()


def main():
    parser = argparse.ArgumentParser(
        description="Upload files using kb-mcp FileService",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Stdioモードで単一ファイルをアップロード
  python upload_via_file_service.py --files spec.docx --dest specifications/spec.docx
  
  # Stdioモードで複数ファイルをアップロード
  python upload_via_file_service.py --files spec1.docx spec2.xlsx --dest specifications/
  
  # HTTPモードでアップロード
  python upload_via_file_service.py --files spec.docx --dest specifications/ --http-url http://localhost:8000/mcp
  
  # カスタムコマンドでMCPサーバを起動
  python upload_via_file_service.py --files spec.docx --dest specifications/ --command python --args "src/main.py"
        """
    )
    
    parser.add_argument(
        "--files", 
        nargs="+", 
        required=True, 
        help="Files to upload"
    )
    parser.add_argument(
        "--dest", 
        required=True, 
        help="Destination path (e.g., 'domain/path/' or 'domain/path/file.docx')"
    )
    parser.add_argument(
        "--overwrite", 
        action="store_true",
        default=True,
        help="Overwrite existing files (default: True)"
    )
    parser.add_argument(
        "--no-overwrite",
        action="store_false",
        dest="overwrite",
        help="Do not overwrite existing files"
    )
    parser.add_argument(
        "--http-url",
        help="HTTP URL for MCP server (if not provided, uses Stdio mode)"
    )
    parser.add_argument(
        "--command",
        default="uv",
        help="Command to launch MCP server (Stdio mode only, default: uv)"
    )
    parser.add_argument(
        "--args",
        default="run src/main.py",
        help="Arguments for MCP server command (Stdio mode only, default: 'run src/main.py')"
    )
    parser.add_argument(
        "--kb-mcp-root",
        type=Path,
        default=Path.cwd(),
        help="Path to kb-mcp repository root (default: current directory)"
    )
    
    args = parser.parse_args()
    
    # ファイルパスの解決
    file_paths = []
    for file_str in args.files:
        file_path = Path(file_str).resolve()
        if not file_path.exists():
            print(f"❌ File not found: {file_path}")
            return
        file_paths.append(file_path)
    
    # アップロード実行
    uploader = SimpleUploader(args.kb_mcp_root)
    
    try:
        # MCPサーバに接続
        if args.http_url:
            if not uploader.connect_http(args.http_url):
                return
        else:
            if not uploader.connect_stdio(args.command, args.args):
                return
        
        time.sleep(2)  # 接続安定化のため少し待つ
        
        # ファイルをアップロード
        uploader.upload_files(file_paths, args.dest, args.overwrite)
        
    finally:
        # 切断
        uploader.disconnect()


if __name__ == "__main__":
    main()
