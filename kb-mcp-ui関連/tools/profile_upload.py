"""
アップロード処理の詳細プロファイリングツール

各処理ステップの実行時間を詳細に計測し、ボトルネックを特定します。

使用方法:
    uv run tools/profile_upload.py [ファイルパス]
    uv run tools/profile_upload.py data/sample.xlsx
    uv run tools/profile_upload.py --all  # すべてのベンチマークファイルを測定
"""

import sys
import time
import base64
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List
import json

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config import get_config
from src.protocol.server import KnowledgeBaseMCPServer
from src.services.document import DocumentService
from src.protocol.schemas import UploadDocumentInput


@dataclass
class StepTiming:
    """各ステップの計測結果"""
    name: str
    duration: float  # 秒
    metadata: Dict = field(default_factory=dict)


@dataclass
class ProfileResult:
    """プロファイリング結果全体"""
    file_path: str
    file_size_kb: float
    total_time: float
    steps: List[StepTiming]
    
    def print_report(self):
        """レポートを出力"""
        print("\n" + "=" * 80)
        print(f"📊 Upload Performance Profile: {Path(self.file_path).name}")
        print("=" * 80)
        print(f"File Size: {self.file_size_kb:.2f} KB")
        print(f"Total Time: {self.total_time:.3f}s")
        print()
        
        # ステップごとの詳細
        print(f"{'Step':<30} {'Time (s)':>12} {'%':>8} {'Details':>20}")
        print("-" * 80)
        
        for step in self.steps:
            percentage = (step.duration / self.total_time * 100) if self.total_time > 0 else 0
            
            # 詳細情報の抽出
            details = []
            if "chunks" in step.metadata:
                details.append(f"{step.metadata['chunks']} chunks")
            if "documents" in step.metadata:
                details.append(f"{step.metadata['documents']} docs")
            if "vectors" in step.metadata:
                details.append(f"{step.metadata['vectors']} vectors")
            if "time_per_chunk" in step.metadata:
                details.append(f"{step.metadata['time_per_chunk']:.3f}s/chunk")
            
            details_str = ", ".join(details) if details else "-"
            
            print(f"{step.name:<30} {step.duration:>12.3f} {percentage:>7.1f}% {details_str:>20}")
        
        print("=" * 80)
        
        # ボトルネック分析
        sorted_steps = sorted(self.steps, key=lambda x: x.duration, reverse=True)
        print("\n🔍 Bottleneck Analysis (Top 3):")
        for i, step in enumerate(sorted_steps[:3], 1):
            percentage = (step.duration / self.total_time * 100) if self.total_time > 0 else 0
            print(f"  {i}. {step.name}: {step.duration:.3f}s ({percentage:.1f}%)")
        print()


class UploadProfiler:
    """アップロード処理のプロファイラー"""
    
    def __init__(self):
        self.config = get_config()
        self.server = KnowledgeBaseMCPServer(self.config)
        self.server.initialize()
        
        # DocumentServiceを使用（完全なアップロードフロー）
        self.document_service = DocumentService(
            vectorstore_service=self.server.vectorstore_service,
            embedding_service=self.server.embedding_service,
            config=self.config
        )
    
    def profile_file(self, file_path: Path, domain: str = "benchmark") -> ProfileResult:
        """ファイルのアップロード処理をプロファイリング"""
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        file_size_kb = file_path.stat().st_size / 1024
        steps = []
        overall_start = time.time()
        
        # Step 1: ファイル読み込み & Base64エンコード
        print(f"🔄 Step 1/3: File Loading & Encoding...")
        start = time.time()
        with open(file_path, 'rb') as f:
            file_bytes = f.read()
        file_content_b64 = base64.b64encode(file_bytes).decode('utf-8')
        duration = time.time() - start
        steps.append(StepTiming(
            name="1. File I/O & Encoding",
            duration=duration,
            metadata={"bytes": len(file_bytes)}
        ))
        
        # Step 2-3: 完全なアップロードフロー
        # （ファイル保存、ドキュメント変換、チャンク分割、エンベディング生成、ベクトルDB格納）
        print(f"🔄 Step 2-3: Complete Upload Flow (Save + Vectorize)...")
        start = time.time()
        
        # DocumentServiceを使用して完全なアップロード
        upload_path = f"{domain}/{file_path.name}"
        upload_input = UploadDocumentInput(
            path=upload_path,
            content=file_content_b64,
            encoding="base64",
            overwrite=True
        )
        
        result = self.document_service.upload_document(upload_input)
        
        duration = time.time() - start
        
        steps.append(StepTiming(
            name="2-3. Complete Upload (Save + Convert + Split + Embed + Store)",
            duration=duration,
            metadata={
                "chunks": result.chunks_created,
                "time_per_chunk": duration / result.chunks_created if result.chunks_created > 0 else 0,
                "status": result.status
            }
        ))
        
        total_time = time.time() - overall_start
        
        return ProfileResult(
            file_path=str(file_path),
            file_size_kb=file_size_kb,
            total_time=total_time,
            steps=steps
        )
    
    def cleanup(self):
        """クリーンアップ"""
        try:
            self.server.close()
        except Exception as e:
            print(f"Warning: Cleanup error: {e}")


def create_benchmark_files(output_dir: Path):
    """ベンチマーク用のテストファイルを作成"""
    output_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        import pandas as pd
        
        # 様々なサイズのExcelファイルを作成
        sizes = {
            "small_1kb": 10,      # 10行
            "medium_10kb": 100,   # 100行
            "large_100kb": 1000,  # 1000行
        }
        
        created_files = []
        for name, rows in sizes.items():
            file_path = output_dir / f"benchmark_{name}.xlsx"
            
            # ダミーデータ作成
            data = {
                "ID": list(range(1, rows + 1)),
                "Name": [f"Item {i}" for i in range(1, rows + 1)],
                "Description": [f"This is a description for item {i} with some additional text to make it longer" for i in range(1, rows + 1)],
                "Value": [i * 1.5 for i in range(1, rows + 1)],
                "Category": [f"Category {i % 10}" for i in range(1, rows + 1)],
            }
            
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)
            
            created_files.append(file_path)
            print(f"✅ Created: {file_path.name} ({file_path.stat().st_size / 1024:.1f} KB)")
        
        return created_files
    
    except ImportError:
        print("❌ pandas not available, cannot create benchmark files")
        return []


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="Profile upload performance")
    parser.add_argument("file", nargs="?", help="File to profile")
    parser.add_argument("--all", action="store_true", help="Profile all benchmark files")
    parser.add_argument("--create", action="store_true", help="Create benchmark files")
    parser.add_argument("--output", default="summary/benchmark", help="Output directory for results")
    
    args = parser.parse_args()
    
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # ベンチマークファイル作成
    if args.create:
        benchmark_dir = Path("data/benchmark")
        create_benchmark_files(benchmark_dir)
        return
    
    # プロファイリング実行
    profiler = UploadProfiler()
    
    try:
        if args.all:
            # すべてのベンチマークファイルを測定
            benchmark_dir = Path("data/benchmark")
            if not benchmark_dir.exists():
                print("❌ Benchmark directory not found. Run with --create first.")
                return
            
            files = sorted(benchmark_dir.glob("benchmark_*.xlsx"))
            if not files:
                print("❌ No benchmark files found. Run with --create first.")
                return
            
            results = []
            for file_path in files:
                print(f"\n{'='*80}")
                print(f"Profiling: {file_path.name}")
                print(f"{'='*80}")
                
                result = profiler.profile_file(file_path)
                result.print_report()
                results.append(result)
            
            # サマリー保存
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            summary_file = output_dir / f"profile_summary_{timestamp}.json"
            
            summary_data = {
                "timestamp": timestamp,
                "results": [
                    {
                        "file": result.file_path,
                        "file_size_kb": result.file_size_kb,
                        "total_time": result.total_time,
                        "steps": [
                            {
                                "name": step.name,
                                "duration": step.duration,
                                "metadata": step.metadata
                            }
                            for step in result.steps
                        ]
                    }
                    for result in results
                ]
            }
            
            with open(summary_file, 'w', encoding='utf-8') as f:
                json.dump(summary_data, f, indent=2, ensure_ascii=False)
            
            print(f"\n📄 Summary saved to: {summary_file}")
        
        elif args.file:
            # 単一ファイルを測定
            file_path = Path(args.file)
            result = profiler.profile_file(file_path)
            result.print_report()
            
            # 結果を保存
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            result_file = output_dir / f"profile_{file_path.stem}_{timestamp}.txt"
            
            # テキスト形式で保存
            with open(result_file, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write(f"📊 Upload Performance Profile: {file_path.name}\n")
                f.write("=" * 80 + "\n")
                f.write(f"File Size: {result.file_size_kb:.2f} KB\n")
                f.write(f"Total Time: {result.total_time:.3f}s\n")
                f.write("\n")
                
                f.write(f"{'Step':<30} {'Time (s)':>12} {'%':>8} {'Details':>20}\n")
                f.write("-" * 80 + "\n")
                
                for step in result.steps:
                    percentage = (step.duration / result.total_time * 100) if result.total_time > 0 else 0
                    
                    details = []
                    if "chunks" in step.metadata:
                        details.append(f"{step.metadata['chunks']} chunks")
                    if "documents" in step.metadata:
                        details.append(f"{step.metadata['documents']} docs")
                    if "time_per_chunk" in step.metadata:
                        details.append(f"{step.metadata['time_per_chunk']:.3f}s/chunk")
                    
                    details_str = ", ".join(details) if details else "-"
                    f.write(f"{step.name:<30} {step.duration:>12.3f} {percentage:>7.1f}% {details_str:>20}\n")
                
                f.write("=" * 80 + "\n")
                
                # ボトルネック分析
                sorted_steps = sorted(result.steps, key=lambda x: x.duration, reverse=True)
                f.write("\n🔍 Bottleneck Analysis (Top 3):\n")
                for i, step in enumerate(sorted_steps[:3], 1):
                    percentage = (step.duration / result.total_time * 100) if result.total_time > 0 else 0
                    f.write(f"  {i}. {step.name}: {step.duration:.3f}s ({percentage:.1f}%)\n")
            
            print(f"\n📄 Result saved to: {result_file}")
        
        else:
            parser.print_help()
    
    finally:
        profiler.cleanup()


if __name__ == "__main__":
    main()
