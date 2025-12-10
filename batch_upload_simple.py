#!/usr/bin/env python3
"""
一度のMCP接続で複数ファイルをアップロードするスクリプト（シンプル版）

Usage:
    スクリプト内のuploadsリストを編集してから実行:
    python batch_upload_simple.py
"""

import sys
import time
from pathlib import Path
from typing import List, Tuple


class BatchUploader:
    """一度の接続で複数ファイルをアップロードするクラス"""
    
    def __init__(self, kb_mcp_root: Path = None):
        if kb_mcp_root is None:
            kb_mcp_root = Path.cwd()
        
        self.kb_mcp_root = kb_mcp_root
        sys.path.insert(0, str(kb_mcp_root))
        
        # MCPモジュールをインポート
        from ui.services.mcp_client import MCPClientManager
        from ui.services.file_service import FileService
        
        self.client_manager = MCPClientManager()
        self.file_service = FileService(self.client_manager)
        
        self.success_count = 0
        self.fail_count = 0
        self.failed_files = []
        
        print("✓ Initialized MCP client and file service")
    
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
            time.sleep(1)  # 接続安定化
            return True
        else:
            print(f"❌ {result}")
            return False
    
    def upload_single_file(self, file_path: str, dest_path: str, overwrite: bool = True):
        """単一ファイルをアップロード（MCPサーバは接続済み）"""
        try:
            file_path_obj = Path(file_path)
            
            if not file_path_obj.exists():
                print(f"  ⚠️  File not found: {file_path}")
                self.fail_count += 1
                self.failed_files.append((file_path, dest_path, "File not found"))
                return False
            
            # FileService.upload_filesを使用（ジェネレータ）
            gen = self.file_service.upload_files(
                source="explorer",
                local_paths=[str(file_path_obj)],
                upload_path_str=dest_path,
                overwrite_flag=overwrite,
                progress_callback=None
            )
            
            # ジェネレータから最後のログを取得
            last_log = ""
            for log_chunk in gen:
                last_log = log_chunk
            
            # 成功判定（簡易版）
            if "✅" in last_log or "Upload Complete" in last_log:
                print(f"  ✅ Success: {file_path_obj.name}")
                self.success_count += 1
                return True
            else:
                print(f"  ❌ Failed: {file_path_obj.name}")
                self.fail_count += 1
                self.failed_files.append((file_path, dest_path, "Upload failed"))
                return False
            
        except Exception as e:
            print(f"  ❌ Error: {file_path}")
            print(f"     {str(e)}")
            self.fail_count += 1
            self.failed_files.append((file_path, dest_path, str(e)))
            return False
    
    def upload_batch(self, uploads: List[Tuple[str, str, str]], overwrite: bool = True):
        """複数ファイルをバッチアップロード
        
        Args:
            uploads: (file_path, dest_path, description) のタプルのリスト
            overwrite: 既存ファイルを上書きするか
        """
        total = len(uploads)
        
        print(f"\n{'='*60}")
        print(f"🚀 Starting batch upload: {total} file(s)")
        print(f"{'='*60}\n")
        
        for i, (file_path, dest_path, description) in enumerate(uploads, 1):
            print(f"[{i}/{total}] {description or file_path}")
            print(f"  File: {file_path}")
            print(f"  Dest: {dest_path}")
            
            self.upload_single_file(file_path, dest_path, overwrite)
            print()
        
        self.print_summary()
    
    def disconnect(self):
        """MCPサーバから切断"""
        if self.client_manager.connection_status["connected"]:
            result = self.client_manager.disconnect()
            print(f"\n{result}")
        
        # ワーカープロセスを停止
        self.client_manager.stop_worker()
    
    def print_summary(self):
        """アップロード結果のサマリーを表示"""
        print(f"{'='*60}")
        print(f"📊 Upload Summary")
        print(f"{'='*60}")
        print(f"✅ Success: {self.success_count}")
        print(f"❌ Failed:  {self.fail_count}")
        
        if self.failed_files:
            print(f"\n⚠️  Failed files:")
            for file_path, dest_path, reason in self.failed_files:
                print(f"   - {file_path} → {dest_path}")
                if reason:
                    print(f"     Reason: {reason}")
        
        print(f"{'='*60}\n")


def main():
    """
    ここでアップロードするファイルのリストを定義します
    """
    
    # ========================================
    # ここを編集してファイルリストを設定
    # ========================================
    uploads = [
        # (ファイルパス, アップロード先, 説明)
        ("files/test/test.md", "domain/path/", "Test markdown file"),
        ("files/specs/spec1.docx", "specifications/", "Specification document 1"),
        ("files/specs/spec2.docx", "specifications/", "Specification document 2"),
        ("files/design/design.xlsx", "designs/", "Design spreadsheet"),
        ("files/manuals/manual.pdf", "manuals/", "User manual"),
    ]
    
    # ========================================
    # 設定（必要に応じて変更）
    # ========================================
    OVERWRITE = True  # 既存ファイルを上書きするか
    USE_HTTP = False  # HTTPモードを使うか（Falseの場合はStdioモード）
    HTTP_URL = "http://localhost:8000/mcp"  # HTTPモードのURL
    
    # ========================================
    # アップロード実行
    # ========================================
    uploader = BatchUploader()
    
    try:
        # MCPサーバに接続
        if USE_HTTP:
            if not uploader.connect(mode="http", http_url=HTTP_URL):
                return
        else:
            if not uploader.connect(mode="stdio"):
                return
        
        # アップロード実行
        uploader.upload_batch(uploads, OVERWRITE)
        
    finally:
        # 切断
        uploader.disconnect()
    
    # 失敗があれば終了コード1で終了
    if uploader.fail_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
