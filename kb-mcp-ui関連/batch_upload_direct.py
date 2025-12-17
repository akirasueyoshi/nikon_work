#!/usr/bin/env python3
"""
一度の接続で複数ファイルをアップロードするスクリプト（詳細モニタリング版）

特徴:
- MCPサーバを直接起動するため、ベクトル化の詳細なプロセスがターミナルに表示される
- バッチ処理の進捗（Dense/Sparse/Upsert時間）がリアルタイムで確認可能
- DirectFileUploaderを使用
- batch_upload_simple.pyと同じパス指定形式

Usage:
    スクリプト内のuploadsリストを編集してから実行:
    
    uploads = [
        # (ファイルパス, アップロード先, 説明)
        ("files/doc.md", "domain/path/", "Document"),
    ]
    
    python batch_upload_direct.py
"""

import sys
import time
import asyncio
from pathlib import Path
from typing import List, Tuple
from datetime import datetime


class BatchUploaderDirect:
    """DirectFileUploaderを使用した詳細モニタリング対応アップローダー"""
    
    def __init__(self, kb_mcp_root: Path = None):
        if kb_mcp_root is None:
            kb_mcp_root = Path.cwd()
        
        self.kb_mcp_root = kb_mcp_root
        sys.path.insert(0, str(kb_mcp_root))
        
        # MCPモジュールをインポート
        from src.config import get_config
        from src.protocol.server import KnowledgeBaseMCPServer
        from src.utils.file_uploader_direct import DirectFileUploader
        
        # サーバー初期化
        self.config = get_config()
        self.server = KnowledgeBaseMCPServer(self.config)
        self.server.initialize()
        
        # DirectFileUploaderを作成
        self.uploader = DirectFileUploader(self.server)
        
        self.success_count = 0
        self.fail_count = 0
        self.failed_files = []
        self.upload_stats = []
        
        print("✓ Initialized MCP server and direct file uploader")
        print("✓ Detailed vectorization monitoring enabled")
    
    async def upload_single_file(
        self, 
        file_path: str, 
        dest_path: str, 
        overwrite: bool = True
    ):
        """単一ファイルをアップロード
        
        Args:
            file_path: ローカルファイルパス
            dest_path: アップロード先パス（"domain/relative/path.ext"形式）
            overwrite: 上書きフラグ
        """
        try:
            file_path_obj = Path(file_path)
            
            if not file_path_obj.exists():
                print(f"  ⚠️  File not found: {file_path}")
                self.fail_count += 1
                self.failed_files.append((file_path, dest_path, "File not found"))
                return False
            
            # dest_pathからドメインと相対パスを抽出
            # dest_pathの形式:
            #   - "domain/path/" → ファイル名を自動追加
            #   - "domain/path/file.ext" → そのまま使用
            #   - "domain/" → ドメイン直下にファイル名で配置
            
            dest_path = dest_path.rstrip("/")  # 末尾のスラッシュを削除
            parts = dest_path.split("/", 1)
            
            if len(parts) < 2:
                # パスにスラッシュがない場合はドメイン名のみ
                domain = parts[0]
                relative_path = file_path_obj.name
            else:
                domain, path_part = parts
                
                # path_partの末尾にファイル名が含まれているか確認
                # 拡張子がない場合はディレクトリとみなす
                if "." not in Path(path_part).name or path_part.endswith("/"):
                    # ディレクトリパス → ファイル名を追加
                    relative_path = f"{path_part}/{file_path_obj.name}"
                else:
                    # ファイルパス → そのまま使用
                    relative_path = path_part
            
            # ファイルサイズを取得
            file_size = file_path_obj.stat().st_size
            file_size_mb = file_size / (1024 * 1024)
            
            # 時間計測開始
            start_time = time.time()
            
            print(f"\n{'='*80}")
            print(f"📤 Uploading: {file_path_obj.name}")
            print(f"   Destination: {dest_path}")
            print(f"   Size: {file_size_mb:.2f} MB")
            print(f"{'='*80}")
            
            # アップロード実行（DirectFileUploaderを使用）
            # これにより、サーバ側の詳細なログがターミナルに出力される
            result = await self.uploader.upload_file(
                file_path=file_path_obj,
                domain=domain,
                relative_path=relative_path,
                overwrite=overwrite
            )
            
            # 時間計測終了
            end_time = time.time()
            elapsed_time = end_time - start_time
            
            if result.status == "success":
                print(f"\n  ✅ Success: {file_path_obj.name}")
                print(f"     Chunks: {result.chunks}")
                print(f"     Total Time: {elapsed_time:.2f} sec")
                
                self.upload_stats.append({
                    'file_name': file_path_obj.name,
                    'file_path': file_path,
                    'file_size_bytes': file_size,
                    'file_size_mb': file_size_mb,
                    'upload_time_sec': elapsed_time,
                    'chunk_count': result.chunks,
                    'success': True
                })
                
                self.success_count += 1
                return True
            else:
                print(f"\n  ❌ Failed: {file_path_obj.name}")
                print(f"     Error: {result.message}")
                
                self.upload_stats.append({
                    'file_name': file_path_obj.name,
                    'file_path': file_path,
                    'file_size_bytes': file_size,
                    'file_size_mb': file_size_mb,
                    'upload_time_sec': elapsed_time,
                    'chunk_count': 0,
                    'success': False
                })
                
                self.fail_count += 1
                self.failed_files.append((file_path, dest_path, result.message))
                return False
            
        except Exception as e:
            print(f"\n  ❌ Error: {file_path}")
            print(f"     {str(e)}")
            self.fail_count += 1
            self.failed_files.append((file_path, dest_path, str(e)))
            return False
    
    async def upload_batch(
        self, 
        uploads: List[Tuple[str, str, str]], 
        overwrite: bool = True
    ):
        """複数ファイルをバッチアップロード
        
        Args:
            uploads: (file_path, dest_path, description) のタプルのリスト
            overwrite: 既存ファイルを上書きするか
        """
        total = len(uploads)
        
        print(f"\n{'='*80}")
        print(f"🚀 Starting batch upload: {total} file(s)")
        print(f"{'='*80}")
        
        for i, (file_path, dest_path, description) in enumerate(uploads, 1):
            print(f"\n[{i}/{total}] {description or file_path}")
            
            await self.upload_single_file(file_path, dest_path, overwrite)
        
        self.print_summary()
        self.save_summary_markdown()
    
    def print_summary(self):
        """アップロード結果のサマリーを表示"""
        print(f"\n{'='*80}")
        print(f"📊 Upload Summary")
        print(f"{'='*80}")
        print(f"✅ Success: {self.success_count}")
        print(f"❌ Failed:  {self.fail_count}")
        
        if self.failed_files:
            print(f"\n⚠️  Failed files:")
            for file_path, dest_path, reason in self.failed_files:
                print(f"   - {file_path} → {dest_path}")
                if reason:
                    print(f"     Reason: {reason}")
        
        # 統計情報を表示
        if self.upload_stats:
            print(f"\n{'='*80}")
            print(f"⏱️  Upload Statistics")
            print(f"{'='*80}")
            
            successful_stats = [s for s in self.upload_stats if s['success']]
            
            if successful_stats:
                total_size = sum(s['file_size_mb'] for s in successful_stats)
                total_time = sum(s['upload_time_sec'] for s in successful_stats)
                total_chunks = sum(s['chunk_count'] for s in successful_stats)
                avg_time = total_time / len(successful_stats) if successful_stats else 0
                avg_speed = total_size / total_time if total_time > 0 else 0
                avg_chunks = total_chunks / len(successful_stats) if successful_stats else 0
                
                print(f"Total Size:      {total_size:.2f} MB")
                print(f"Total Time:      {total_time:.2f} sec")
                print(f"Total Chunks:    {total_chunks}")
                print(f"Average Time:    {avg_time:.2f} sec/file")
                print(f"Average Speed:   {avg_speed:.2f} MB/sec")
                print(f"Average Chunks:  {avg_chunks:.2f} chunks/file")
                
                print(f"\n📋 Individual File Statistics:")
                print(f"{'No.':<4} {'File Name':<30} {'Size (MB)':<12} {'Time (sec)':<12} {'Chunks':<10} {'Speed (MB/s)':<12}")
                print(f"{'-'*90}")
                
                for i, stat in enumerate(self.upload_stats, 1):
                    speed = stat['file_size_mb'] / stat['upload_time_sec'] if stat['upload_time_sec'] > 0 else 0
                    status = "✅" if stat['success'] else "❌"
                    
                    file_name = stat['file_name']
                    if len(file_name) > 28:
                        file_name = file_name[:25] + "..."
                    
                    print(f"{i:<4} {file_name:<30} {stat['file_size_mb']:<12.2f} {stat['upload_time_sec']:<12.2f} {stat['chunk_count']:<10} {speed:<12.2f} {status}")
        
        print(f"{'='*80}\n")
    
    def save_summary_markdown(self, output_dir: str = "summary/upload"):
        """サマリーをMarkdown形式で保存"""
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"upload_summary_direct_{timestamp}.md"
        filepath = output_path / filename
        
        md_content = []
        md_content.append("# Upload Summary Report (Direct Mode)\n")
        md_content.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        md_content.append("**Mode:** Direct Server Access (Detailed Monitoring)\n")
        md_content.append("---\n")
        
        # サマリー
        md_content.append("## 📊 Upload Summary\n")
        md_content.append(f"- ✅ **Success:** {self.success_count}\n")
        md_content.append(f"- ❌ **Failed:** {self.fail_count}\n")
        
        # 失敗ファイル
        if self.failed_files:
            md_content.append("\n### ⚠️ Failed Files\n")
            for file_path, dest_path, reason in self.failed_files:
                md_content.append(f"- **File:** `{file_path}`\n")
                md_content.append(f"  - **Destination:** `{dest_path}`\n")
                if reason:
                    md_content.append(f"  - **Reason:** {reason}\n")
        
        # 統計情報
        if self.upload_stats:
            successful_stats = [s for s in self.upload_stats if s['success']]
            
            if successful_stats:
                total_size = sum(s['file_size_mb'] for s in successful_stats)
                total_time = sum(s['upload_time_sec'] for s in successful_stats)
                total_chunks = sum(s['chunk_count'] for s in successful_stats)
                avg_time = total_time / len(successful_stats) if successful_stats else 0
                avg_speed = total_size / total_time if total_time > 0 else 0
                avg_chunks = total_chunks / len(successful_stats) if successful_stats else 0
                
                md_content.append("\n## ⏱️ Upload Statistics\n")
                md_content.append(f"- **Total Size:** {total_size:.2f} MB\n")
                md_content.append(f"- **Total Time:** {total_time:.2f} sec\n")
                md_content.append(f"- **Total Chunks:** {total_chunks}\n")
                md_content.append(f"- **Average Time:** {avg_time:.2f} sec/file\n")
                md_content.append(f"- **Average Speed:** {avg_speed:.2f} MB/sec\n")
                md_content.append(f"- **Average Chunks:** {avg_chunks:.2f} chunks/file\n")
                
                # 個別ファイル統計
                md_content.append("\n## 📋 Individual File Statistics\n")
                md_content.append("| No. | File Name | Size (MB) | Time (sec) | Chunks | Speed (MB/s) | Status |\n")
                md_content.append("|-----|-----------|-----------|------------|--------|--------------|--------|\n")
                
                for i, stat in enumerate(self.upload_stats, 1):
                    speed = stat['file_size_mb'] / stat['upload_time_sec'] if stat['upload_time_sec'] > 0 else 0
                    status = "✅" if stat['success'] else "❌"
                    
                    md_content.append(
                        f"| {i} | `{stat['file_name']}` | {stat['file_size_mb']:.2f} | "
                        f"{stat['upload_time_sec']:.2f} | {stat['chunk_count']} | "
                        f"{speed:.2f} | {status} |\n"
                    )
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.writelines(md_content)
        
        print(f"✓ Summary saved to: {filepath}")
    
    def cleanup(self):
        """リソースのクリーンアップ"""
        if self.server:
            self.server.close()
            print("✓ Server closed")


async def main():
    """
    ここでアップロードするファイルのリストを定義します
    """
    
    # ========================================
    # ここを編集してファイルリストを設定
    # ========================================
    uploads = [
        # (ファイルパス, アップロード先, 説明)
        ("files/output_excels/excel_1500KB.xlsx", "domain/path/", ""),
    ]
    
    # ========================================
    # 設定
    # ========================================
    OVERWRITE = True  # 既存ファイルを上書きするか
    
    # ========================================
    # アップロード実行
    # ========================================
    uploader = BatchUploaderDirect()
    
    try:
        await uploader.upload_batch(uploads, OVERWRITE)
    finally:
        uploader.cleanup()


if __name__ == "__main__":
    asyncio.run(main())