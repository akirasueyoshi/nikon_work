#!/usr/bin/env python3
"""
複数ファイルを順次アップロードするスクリプト

Usage:
    python sequential_upload.py
"""

import subprocess
import sys
from pathlib import Path
from typing import List, Tuple


class SequentialUploader:
    """複数ファイルを順次アップロードするクラス"""
    
    def __init__(self):
        self.success_count = 0
        self.fail_count = 0
        self.failed_files = []
    
    def upload_file(self, file_path: str, dest_path: str, description: str = "") -> bool:
        """単一ファイルをアップロード"""
        if description:
            print(f"\n{'='*60}")
            print(f"📤 {description}")
            print(f"{'='*60}")
        
        print(f"   File: {file_path}")
        print(f"   Dest: {dest_path}")
        print()
        
        # uv run コマンドを構築
        cmd = [
            "uv", "run",
            "upload_via_file_service.py",
            "--files", file_path,
            "--dest", dest_path
        ]
        
        try:
            # サブプロセスとして実行
            result = subprocess.run(
                cmd,
                capture_output=False,  # 出力をそのまま表示
                text=True,
                check=True  # エラーがあれば例外を発生
            )
            
            print(f"✅ Success: {file_path}")
            self.success_count += 1
            return True
            
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed: {file_path}")
            print(f"   Error code: {e.returncode}")
            self.fail_count += 1
            self.failed_files.append((file_path, dest_path, description))
            return False
        
        except FileNotFoundError:
            print(f"❌ Error: 'uv' command not found")
            print(f"   Please make sure uv is installed and in your PATH")
            sys.exit(1)
    
    def upload_batch(self, uploads: List[Tuple[str, str, str]]):
        """複数ファイルをバッチアップロード
        
        Args:
            uploads: (file_path, dest_path, description) のタプルのリスト
        """
        total = len(uploads)
        print(f"\n{'='*60}")
        print(f"🚀 Starting batch upload: {total} files")
        print(f"{'='*60}\n")
        
        for i, (file_path, dest_path, description) in enumerate(uploads, 1):
            desc_with_counter = f"[{i}/{total}] {description}" if description else f"[{i}/{total}]"
            self.upload_file(file_path, dest_path, desc_with_counter)
        
        # 結果サマリー
        self.print_summary()
    
    def print_summary(self):
        """アップロード結果のサマリーを表示"""
        print(f"\n{'='*60}")
        print(f"📊 Upload Summary")
        print(f"{'='*60}")
        print(f"✅ Success: {self.success_count}")
        print(f"❌ Failed:  {self.fail_count}")
        
        if self.failed_files:
            print(f"\n⚠️  Failed files:")
            for file_path, dest_path, description in self.failed_files:
                print(f"   - {file_path} → {dest_path}")
                if description:
                    print(f"     ({description})")
        
        print(f"{'='*60}\n")


def main():
    """メイン処理"""
    
    # アップロードするファイルのリスト
    # (ファイルパス, アップロード先, 説明)
    uploads = [
        ("files/test/test.md", "domain/path/", "Test markdown file"),
        ("files/test/test2.md", "domain/path/", "Test markdown file"),
    ]
    
    # アップロード実行
    uploader = SequentialUploader()
    uploader.upload_batch(uploads)
    
    # 失敗があれば終了コード1で終了
    if uploader.fail_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
