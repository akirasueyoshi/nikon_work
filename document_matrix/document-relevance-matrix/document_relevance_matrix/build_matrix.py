#!/usr/bin/env python3
"""
リンク抽出結果JSONから関連度マトリクスを計算
"""

import json
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
import csv


def load_links_data(json_path):
    """リンク抽出結果JSONを読み込む"""
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data


def build_adjacency_matrix(data):
    """
    隣接行列を構築（仮想ドキュメントも含む）
    
    Returns:
        matrix: numpy array (directed graph)
        docs: list of document IDs
        all_docs: list of all document IDs (including virtual)
    """
    # 実在するドキュメントIDリスト
    real_docs = sorted([d['id'] for d in data['documents']])
    
    # すべてのリンク先を収集（仮想ドキュメント含む）
    all_doc_ids = set(real_docs)
    for link in data['links']:
        all_doc_ids.add(link['source'])
        all_doc_ids.add(link['target'])
    
    all_docs = sorted(all_doc_ids)
    n = len(all_docs)
    doc_to_idx = {doc: i for i, doc in enumerate(all_docs)}
    
    # 隣接行列を初期化
    matrix = np.zeros((n, n))
    
    # リンク情報から行列を構築
    for link in data['links']:
        source = link['source']
        target = link['target']
        
        if source in doc_to_idx and target in doc_to_idx:
            i = doc_to_idx[source]
            j = doc_to_idx[target]
            matrix[i][j] = 1.0
    
    return matrix, all_docs


def calculate_jaccard_matrix(adjacency):
    """Jaccard係数マトリクスを計算"""
    n = adjacency.shape[0]
    relevance = np.eye(n)
    
    for i in range(n):
        for j in range(i+1, n):
            links_i = set(np.where(adjacency[i] > 0)[0])
            links_j = set(np.where(adjacency[j] > 0)[0])
            
            union = links_i | links_j
            if len(union) > 0:
                jaccard = len(links_i & links_j) / len(union)
                relevance[i][j] = jaccard
                relevance[j][i] = jaccard
    
    return relevance


def calculate_combined_matrix(adjacency):
    """複合指標マトリクスを計算"""
    n = adjacency.shape[0]
    
    # 1. 直接リンク
    m_direct = adjacency.copy()
    
    # 2. 双方向性
    m_bidirectional = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i == j:
                m_bidirectional[i][j] = 1.0
            else:
                a_to_b = adjacency[i][j]
                b_to_a = adjacency[j][i]
                
                if a_to_b and b_to_a:
                    m_bidirectional[i][j] = 1.0
                elif a_to_b or b_to_a:
                    m_bidirectional[i][j] = 0.5
    
    # 3. 共通リンク先
    m_common = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i == j:
                m_common[i][j] = 1.0
            else:
                links_i = set(np.where(adjacency[i] > 0)[0])
                links_j = set(np.where(adjacency[j] > 0)[0])
                
                union = links_i | links_j
                if len(union) > 0:
                    m_common[i][j] = len(links_i & links_j) / len(union)
    
    # 重み付け統合
    relevance = (m_direct * 0.5 + m_bidirectional * 0.3 + m_common * 0.2)
    relevance = np.clip(relevance, 0, 1)
    np.fill_diagonal(relevance, 1.0)
    
    return relevance


def create_ground_truth(relevance_matrix, docs, threshold=0.3, top_k=10):
    """検索評価用の正解データを生成"""
    ground_truth = []
    
    for i, query_doc in enumerate(docs):
        relevances = []
        for j, target_doc in enumerate(docs):
            if i != j:
                score = relevance_matrix[i][j]
                if score >= threshold:
                    relevances.append((target_doc, float(score)))
        
        relevances = sorted(relevances, key=lambda x: x[1], reverse=True)
        relevances = relevances[:top_k]
        
        ground_truth.append({
            "query_doc": query_doc,
            "relevant_docs": relevances,
            "total_relevant": len(relevances),
            "threshold": threshold
        })
    
    return ground_truth


def save_results(data, relevance_combined, relevance_jaccard, all_docs, ground_truth, output_dir="relevance_results", threshold=0.3):
    """
    結果を保存（実在するファイルのみ）
    
    Parameters:
        data: リンク抽出結果
        relevance_combined: 全ドキュメントの複合指標マトリクス
        relevance_jaccard: 全ドキュメントのJaccard係数マトリクス
        all_docs: 全ドキュメントID（仮想含む）
        ground_truth: 正解データ
        output_dir: 出力ディレクトリ
        threshold: 閾値
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 実在するドキュメントIDのみ抽出
    real_doc_ids = [d['id'] for d in data['documents']]
    real_doc_indices = [i for i, doc in enumerate(all_docs) if doc in real_doc_ids]
    
    # 実在ドキュメント間の関連度マトリクスを抽出
    n_real = len(real_doc_ids)
    relevance_combined_real = np.zeros((n_real, n_real))
    relevance_jaccard_real = np.zeros((n_real, n_real))
    
    for i, idx_i in enumerate(real_doc_indices):
        for j, idx_j in enumerate(real_doc_indices):
            relevance_combined_real[i][j] = relevance_combined[idx_i][idx_j]
            relevance_jaccard_real[i][j] = relevance_jaccard[idx_i][idx_j]
    
    # 1. 複合指標マトリクスCSV（実在ファイルのみ）
    df_combined = pd.DataFrame(relevance_combined_real, index=real_doc_ids, columns=real_doc_ids)
    matrix_path = output_dir / f"relevance_matrix_combined_{timestamp}.csv"
    df_combined.to_csv(matrix_path, encoding='utf-8-sig')
    print(f"✓ Saved combined relevance matrix: {matrix_path}")
    
    # 2. Jaccard係数マトリクスCSV（実在ファイルのみ）
    df_jaccard = pd.DataFrame(relevance_jaccard_real, index=real_doc_ids, columns=real_doc_ids)
    jaccard_path = output_dir / f"relevance_matrix_jaccard_{timestamp}.csv"
    df_jaccard.to_csv(jaccard_path, encoding='utf-8-sig')
    print(f"✓ Saved Jaccard relevance matrix: {jaccard_path}")
    
    # 3. エッジリストCSV（閾値以上、実在ファイル間のみ）
    edges = []
    for i, source in enumerate(real_doc_ids):
        for j, target in enumerate(real_doc_ids):
            if i != j:
                score = relevance_combined_real[i][j]
                if score >= threshold:
                    edges.append({
                        'source': source,
                        'target': target,
                        'relevance': round(score, 3)
                    })
    
    edges_path = output_dir / f"relevance_edges_{timestamp}.csv"
    with open(edges_path, "w", encoding="utf-8-sig", newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['source', 'target', 'relevance'])
        writer.writeheader()
        writer.writerows(edges)
    print(f"✓ Saved edge list: {edges_path}")
    
    # 4. 正解データJSON（実在ファイルのみ）
    ground_truth_real = []
    for gt in ground_truth:
        if gt['query_doc'] in real_doc_ids:
            # relevant_docsも実在ファイルのみにフィルタ
            relevant_docs_real = [(doc, score) for doc, score in gt['relevant_docs'] if doc in real_doc_ids]
            ground_truth_real.append({
                "query_doc": gt['query_doc'],
                "relevant_docs": relevant_docs_real,
                "total_relevant": len(relevant_docs_real),
                "threshold": gt['threshold']
            })
    
    gt_path = output_dir / f"ground_truth_{timestamp}.json"
    with open(gt_path, "w", encoding="utf-8") as f:
        json.dump(ground_truth_real, f, ensure_ascii=False, indent=2)
    print(f"✓ Saved ground truth: {gt_path}")
    
    # 5. サマリーJSON
    non_diag = []
    for i in range(n_real):
        for j in range(n_real):
            if i != j:
                non_diag.append(relevance_combined_real[i][j])
    
    summary = {
        "timestamp": timestamp,
        "source_file": str(data.get('metadata', {}).get('source_directory', 'N/A')),
        "total_documents": len(real_doc_ids),
        "total_virtual_documents": len(all_docs) - len(real_doc_ids),
        "total_edges_above_threshold": len(edges),
        "threshold": threshold,
        "statistics": {
            "mean_relevance": round(float(np.mean(non_diag)), 3) if non_diag else 0.0,
            "median_relevance": round(float(np.median(non_diag)), 3) if non_diag else 0.0,
            "std_relevance": round(float(np.std(non_diag)), 3) if non_diag else 0.0,
            "min_relevance": round(float(np.min(non_diag)), 3) if non_diag else 0.0,
            "max_relevance": round(float(np.max(non_diag)), 3) if non_diag else 0.0
        },
        "ground_truth_stats": {
            "total_queries": len(ground_truth_real),
            "avg_relevant_docs_per_query": round(np.mean([g['total_relevant'] for g in ground_truth_real]), 1) if ground_truth_real else 0.0,
            "queries_with_no_relevant": len([g for g in ground_truth_real if g['total_relevant'] == 0])
        }
    }
    
    summary_path = output_dir / f"summary_{timestamp}.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"✓ Saved summary: {summary_path}")
    
    # コンソールサマリー
    print("\n" + "="*60)
    print("RELEVANCE CALCULATION SUMMARY")
    print("="*60)
    print(f"Real documents:            {summary['total_documents']}")
    print(f"Virtual documents:         {summary['total_virtual_documents']}")
    print(f"Edges (≥{threshold}):           {summary['total_edges_above_threshold']}")
    print(f"Mean relevance:            {summary['statistics']['mean_relevance']:.3f}")
    print(f"Median relevance:          {summary['statistics']['median_relevance']:.3f}")
    print(f"\nGround Truth Stats:")
    print(f"Total queries:             {summary['ground_truth_stats']['total_queries']}")
    print(f"Avg relevant docs/query:   {summary['ground_truth_stats']['avg_relevant_docs_per_query']:.1f}")
    print(f"Queries with no relevant:  {summary['ground_truth_stats']['queries_with_no_relevant']}")
    print("="*60)
    
    return matrix_path


def visualize_matrix(relevance_matrix, docs, output_path):
    """ヒートマップ生成（実在ドキュメントのみ）"""
    if len(docs) > 30:
        print(f"Skipping heatmap (too many documents: {len(docs)})")
        return
    
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import seaborn as sns
        
        # 日本語フォント設定（警告を抑制）
        import matplotlib.font_manager as fm
        import warnings
        
        # 日本語フォントを探す
        japanese_fonts = [
            'Yu Gothic',      # Windows
            'Meiryo',         # Windows
            'MS Gothic',      # Windows
            'Hiragino Sans',  # Mac
            'AppleGothic',    # Mac
            'TakaoPGothic',   # Linux
            'IPAGothic',      # Linux
        ]
        
        font_found = False
        for font_name in japanese_fonts:
            if any(font_name.lower() in f.name.lower() for f in fm.fontManager.ttflist):
                plt.rcParams['font.family'] = font_name
                font_found = True
                break
        
        # フォントが見つからない場合は警告を抑制
        if not font_found:
            warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib')
        
        plt.figure(figsize=(16, 14))
        sns.heatmap(relevance_matrix, 
                    xticklabels=docs, 
                    yticklabels=docs,
                    annot=True, 
                    fmt='.2f', 
                    cmap='YlOrRd',
                    vmin=0, 
                    vmax=1,
                    square=True,
                    cbar_kws={'label': 'Relevance Score'})
        
        plt.title('Document Relevance Matrix (Combined)', fontsize=16, pad=20)
        plt.xlabel('Target Document', fontsize=12)
        plt.ylabel('Source Document', fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.yticks(rotation=0)
        plt.tight_layout()
        
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        print(f"✓ Saved heatmap: {output_path}")
        
    except Exception as e:
        print(f"Warning: Could not create heatmap: {e}")


def main():
    """メイン実行関数"""
    import sys
    
    # コマンドライン引数がある場合はそれを使用
    if len(sys.argv) > 1:
        json_path = sys.argv[1]
    else:
        # GUIファイル選択ダイアログを表示
        try:
            import tkinter as tk
            from tkinter import filedialog
            
            # ルートウィンドウを非表示
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            
            # ファイル選択ダイアログ
            print("\n📁 Opening file selection dialog...")
            json_path = filedialog.askopenfilename(
                title="Select links_extracted JSON file",
                initialdir="extraction_results",
                filetypes=[
                    ("JSON files", "*.json"),
                    ("All files", "*.*")
                ]
            )
            
            root.destroy()
            
            # キャンセルされた場合
            if not json_path:
                print("❌ File selection cancelled.")
                return
                
        except Exception as e:
            # GUIが使えない場合はフォールバック
            print(f"⚠️  GUI not available: {e}")
            json_path = input("Enter path to links_extracted JSON file: ").strip()
    
    json_path = Path(json_path)
    if not json_path.exists():
        print(f"Error: File not found: {json_path}")
        return
    
    print("\n" + "="*60)
    print("RELEVANCE CALCULATOR")
    print("="*60)
    
    # データ読み込み
    print(f"\nLoading: {json_path}")
    data = load_links_data(json_path)
    
    # 隣接行列構築
    print("Building adjacency matrix...")
    adjacency, docs = build_adjacency_matrix(data)
    print(f"  Documents: {len(docs)}")
    print(f"  Links: {int(np.sum(adjacency))}")
    
    # Jaccard係数計算
    print("\nCalculating Jaccard similarity...")
    relevance_jaccard = calculate_jaccard_matrix(adjacency)
    
    # 複合指標計算
    print("Calculating combined relevance...")
    relevance_combined = calculate_combined_matrix(adjacency)
    
    # 正解データ生成
    print("Creating ground truth data...")
    ground_truth = create_ground_truth(relevance_combined, docs, threshold=0.3, top_k=10)
    
    # 結果保存
    print("\nSaving results...")
    matrix_path = save_results(data, relevance_combined, relevance_jaccard, docs, ground_truth)
    
    # ヒートマップ生成（実在ドキュメントのみ）
    real_doc_ids = [d['id'] for d in data['documents']]
    real_doc_indices = [i for i, doc in enumerate(docs) if doc in real_doc_ids]
    
    n_real = len(real_doc_ids)
    relevance_combined_real = np.zeros((n_real, n_real))
    for i, idx_i in enumerate(real_doc_indices):
        for j, idx_j in enumerate(real_doc_indices):
            relevance_combined_real[i][j] = relevance_combined[idx_i][idx_j]
    
    output_dir = matrix_path.parent
    heatmap_path = output_dir / f"heatmap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    visualize_matrix(relevance_combined_real, real_doc_ids, heatmap_path)
    
    print(f"\n✓ Relevance calculation completed successfully!")


if __name__ == "__main__":
    main()
