"""
MCP サーバのメモリ使用量監視スクリプト

使用方法:
    uv run tools/monitor_memory.py

オプション:
    --interval, -i: 監視間隔（秒）。デフォルト: 5
    --output, -o: 出力CSVファイル名。デフォルト: memory_usage_YYYYMMDD_HHMMSS.csv
    --process-name, -p: 監視対象プロセス名パターン。デフォルト: kb-mcp-server
"""

import argparse
import csv
import time
from datetime import datetime
from pathlib import Path

import psutil


def find_mcp_server_processes(
    pattern: str = "kb-mcp-server",
    main_only: bool = True,
) -> list[psutil.Process]:
    """
    MCP サーバプロセスを検出する

    Args:
        pattern: 検索パターン
        main_only: True の場合、最もメモリを使用しているプロセスのみ返す
    """
    processes = []
    for proc in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            cmdline = proc.info.get("cmdline") or []
            cmdline_str = " ".join(cmdline).lower()
            if pattern.lower() in cmdline_str:
                processes.append(proc)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    if main_only and processes:
        # メモリ使用量が最大のプロセスをメインとみなす
        processes = sorted(
            processes,
            key=lambda p: p.memory_info().rss if p.is_running() else 0,
            reverse=True,
        )
        return [processes[0]]

    return processes


def get_memory_info(proc: psutil.Process) -> dict | None:
    """プロセスのメモリ情報を取得する"""
    try:
        mem_info = proc.memory_info()
        return {
            "pid": proc.pid,
            "rss_mb": mem_info.rss / 1024 / 1024,  # Resident Set Size
            "vms_mb": mem_info.vms / 1024 / 1024,  # Virtual Memory Size
            "percent": proc.memory_percent(),
        }
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return None


def format_memory_line(timestamp: str, mem_info: dict) -> str:
    """コンソール表示用のフォーマット"""
    return (
        f"[{timestamp}] PID={mem_info['pid']:>6} | "
        f"RSS: {mem_info['rss_mb']:>8.1f} MB | "
        f"VMS: {mem_info['vms_mb']:>10.1f} MB | "
        f"Percent: {mem_info['percent']:>5.1f}%"
    )


def monitor_memory(
    interval: float = 5.0,
    output_file: str | None = None,
    process_pattern: str = "kb-mcp-server",
    main_only: bool = True,
):
    """メモリ使用量を監視してCSVに出力する"""

    # 出力ファイル名の決定
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"memory_usage_{timestamp}.csv"

    output_path = Path(output_file)

    print(f"MCP Server Memory Monitor")
    print(f"=" * 60)
    print(f"Process pattern: {process_pattern}")
    print(f"Interval: {interval} seconds")
    print(f"Output file: {output_path.absolute()}")
    print(f"=" * 60)
    print()

    # CSV ヘッダー書き込み
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "timestamp",
                "pid",
                "rss_mb",
                "vms_mb",
                "memory_percent",
            ]
        )

    print("Waiting for MCP server process...")
    print("Press Ctrl+C to stop monitoring.\n")

    try:
        while True:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            processes = find_mcp_server_processes(process_pattern, main_only)

            if not processes:
                print(f"[{timestamp}] No MCP server process found. Waiting...")
            else:
                for proc in processes:
                    mem_info = get_memory_info(proc)
                    if mem_info:
                        # コンソール表示
                        print(format_memory_line(timestamp, mem_info))

                        # CSV 追記
                        with open(output_path, "a", newline="", encoding="utf-8") as f:
                            writer = csv.writer(f)
                            writer.writerow(
                                [
                                    timestamp,
                                    mem_info["pid"],
                                    f"{mem_info['rss_mb']:.2f}",
                                    f"{mem_info['vms_mb']:.2f}",
                                    f"{mem_info['percent']:.2f}",
                                ]
                            )

            time.sleep(interval)

    except KeyboardInterrupt:
        print(f"\n\nMonitoring stopped.")
        print(f"Results saved to: {output_path.absolute()}")


def main():
    parser = argparse.ArgumentParser(
        description="Monitor MCP server memory usage",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "-i",
        "--interval",
        type=float,
        default=5.0,
        help="Monitoring interval in seconds (default: 5)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=None,
        help="Output CSV file name (default: memory_usage_YYYYMMDD_HHMMSS.csv)",
    )
    parser.add_argument(
        "-p",
        "--process-name",
        type=str,
        default="kb-mcp-server",
        help="Process name pattern to monitor (default: kb-mcp-server)",
    )
    parser.add_argument(
        "-a",
        "--all",
        action="store_true",
        help="Show all matching processes (default: only main process)",
    )

    args = parser.parse_args()

    monitor_memory(
        interval=args.interval,
        output_file=args.output,
        process_pattern=args.process_name,
        main_only=not args.all,
    )


if __name__ == "__main__":
    main()
