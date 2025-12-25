#!/usr/bin/env python3
"""
エクセルファイルから資料間のリンク情報を抽出してJSON化するスクリプト
"""

import pandas as pd
from pathlib import Path
import json
from datetime import datetime
import re


def extract_links_from_excel(excel_path):
    """
    エクセルファイルから機能仕様書名を抽出
    
    フォーマット:
    - B列またはC列に「機能仕様書名」という見出しがある
    - その下に仕様書名が列挙される
    - 「対応内容」という文言が出たらそれ以降は無視
    
    Returns:
        list: 抽出された仕様書名のリスト
    """
    try:
        # header=Noneで生データとして読み込み
        df = pd.read_excel(excel_path, sheet_name=0, header=None)
        
        links = []
        
        # B列(index=1)とC列(index=2)をチェック
        for col_idx in [1, 2]:
            if col_idx >= len(df.columns):
                continue
            
            col = df.iloc[:, col_idx]
            
            # 「機能仕様書名」の開始位置を探す
            start_idx = None
            end_idx = None
            
            for idx, cell in enumerate(col):
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    # 開始マーカー
                    if '機能仕様書名' in cell_str:
                        start_idx = idx + 1  # 次の行から
                        continue
                    
                    # 終了マーカー
                    if start_idx is not None and '対応内容' in cell_str:
                        end_idx = idx
                        break
            
            # 範囲内のセルを抽出
            if start_idx is not None:
                if end_idx is None:
                    end_idx = len(col)
                
                for idx in range(start_idx, end_idx):
                    if idx >= len(col):
                        break
                    
                    cell = col.iloc[idx]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()
                        
                        # 空行や数値のみは除外
                        if cell_str and not cell_str.replace('.', '').replace('-', '').isdigit():
                            # 明らかに仕様書名ではない文字列を除外
                            if cell_str not in ['NaN', 'nan', '']:
                                links.append(cell_str)
        
        # 重複除去
        links = list(dict.fromkeys(links))  # 順序を保持しつつ重複除去
        
        return links
    
    except Exception as e:
        print(f"Error reading {excel_path.name}: {e}")
        return []


def normalize_doc_name(name):
    """
    資料名を正規化（比較用）
    - 拡張子を除去
    - 空白を統一
    - 日付パターンを認識
    """
    # 拡張子を除去
    name = re.sub(r'\.(xlsx?|xls|docx?|pdf)$', '', name, flags=re.IGNORECASE)
    
    # 全角・半角スペース、アンダースコアなどを統一
    name = name.replace('　', ' ').replace('_', ' ').strip()
    
    return name


def calculate_jaccard_similarity(links_a, links_b):
    """
    2つのリンクリスト間のJaccard係数を計算
    
    Parameters:
        links_a: 資料Aが持つリンク先のセット
        links_b: 資料Bが持つリンク先のセット
    
    Returns:
        float: Jaccard係数 (0.0 ~ 1.0)
    """
    set_a = set(links_a)
    set_b = set(links_b)
    
    if len(set_a) == 0 and len(set_b) == 0:
        return 0.0
    
    intersection = len(set_a & set_b)
    union = len(set_a | set_b)
    
    if union == 0:
        return 0.0
    
    return intersection / union


def build_document_graph(excel_dir):
    """
    全エクセルファイルからドキュメントグラフを構築（再帰的に探索）
    
    Returns:
        dict: {
            'documents': [資料情報のリスト],
            'links': [リンク情報のリスト],
            'metadata': {メタデータ},
            'relevance_matrix': {共通情報ベースの関連度マトリクス}
        }
    """
    excel_dir = Path(excel_dir)
    
    # 全エクセルファイルを再帰的に取得（一時ファイルを除外）
    print(f"Searching for Excel files in {excel_dir} (recursively)...")
    excel_files = []
    
    # rglob を使って再帰的に検索
    for pattern in ['**/*.xlsx', '**/*.xls']:
        for f in excel_dir.glob(pattern):
            # 一時ファイルと隠しファイルを除外
            if not f.name.startswith('~$') and not f.name.startswith('.'):
                excel_files.append(f)
    
    excel_files = sorted(set(excel_files))  # 重複除去・ソート
    
    print(f"Found {len(excel_files)} Excel files")
    
    if len(excel_files) == 0:
        print("Warning: No Excel files found!")
        return {
            'metadata': {
                'extraction_date': datetime.now().isoformat(),
                'source_directory': str(excel_dir),
                'total_documents': 0,
                'total_matched_links': 0,
                'total_unmatched_links': 0,
                'relevance_calculation_method': 'jaccard_similarity'
            },
            'documents': [],
            'links': [],
            'unmatched_links': [],
            'relevance_matrix': {}
        }
    
    # ディレクトリ構造を表示
    print("\nDirectory structure:")
    dirs = set([f.parent for f in excel_files])
    for d in sorted(dirs):
        relative_path = d.relative_to(excel_dir) if d != excel_dir else Path(".")
        count = len([f for f in excel_files if f.parent == d])
        print(f"  {relative_path}: {count} file(s)")
    print()
    
    # ドキュメント情報を収集
    documents = []
    raw_links = {}  # {source_doc: [linked_docs]}
    
    for excel_file in excel_files:
        doc_name = excel_file.stem
        relative_path = excel_file.relative_to(excel_dir)
        print(f"Processing: {relative_path}")
        
        # リンク先を抽出
        links = extract_links_from_excel(excel_file)
        
        # ドキュメント情報を記録
        doc_info = {
            'id': doc_name,
            'filename': excel_file.name,
            'path': str(excel_file),
            'relative_path': str(relative_path),
            'directory': str(excel_file.parent.relative_to(excel_dir)),
            'normalized_name': normalize_doc_name(doc_name),
            'extracted_links_count': len(links),
            'extracted_links': links
        }
        documents.append(doc_info)
        raw_links[doc_name] = links
        
        print(f"  -> Found {len(links)} links")
    
    # 全ドキュメント名リスト（正規化版も作成）
    all_doc_names = {normalize_doc_name(d['id']): d['id'] for d in documents}
    
    # リンクをマッチング
    matched_links = []
    unmatched_links = []
    doc_to_matched_links = {}  # {doc_id: [matched_link_targets]}
    
    for source_doc, link_list in raw_links.items():
        doc_to_matched_links[source_doc] = []
        
        for link in link_list:
            normalized_link = normalize_doc_name(link)
            
            # 完全一致を試みる
            matched = False
            if normalized_link in all_doc_names:
                target_doc = all_doc_names[normalized_link]
                matched_links.append({
                    'source': source_doc,
                    'target': target_doc,
                    'original_text': link,
                    'match_type': 'exact'
                })
                doc_to_matched_links[source_doc].append(target_doc)
                matched = True
            else:
                # 部分一致を試みる
                for norm_doc, actual_doc in all_doc_names.items():
                    # より長い方を基準に部分一致判定
                    if len(normalized_link) > len(norm_doc):
                        if norm_doc in normalized_link:
                            matched_links.append({
                                'source': source_doc,
                                'target': actual_doc,
                                'original_text': link,
                                'match_type': 'partial'
                            })
                            doc_to_matched_links[source_doc].append(actual_doc)
                            matched = True
                            break
                    else:
                        if normalized_link in norm_doc:
                            matched_links.append({
                                'source': source_doc,
                                'target': actual_doc,
                                'original_text': link,
                                'match_type': 'partial'
                            })
                            doc_to_matched_links[source_doc].append(actual_doc)
                            matched = True
                            break
            
            if not matched:
                unmatched_links.append({
                    'source': source_doc,
                    'original_text': link,
                    'normalized': normalized_link
                })
    
    # 共通情報ベースの関連度マトリクスを計算（Jaccard係数）
    print("\nCalculating relevance matrix based on common links (Jaccard similarity)...")
    doc_ids = [d['id'] for d in documents]
    relevance_matrix = {}
    
    for i, doc_a in enumerate(doc_ids):
        relevance_matrix[doc_a] = {}
        links_a = doc_to_matched_links.get(doc_a, [])
        
        for j, doc_b in enumerate(doc_ids):
            if doc_a == doc_b:
                relevance_matrix[doc_a][doc_b] = 1.0
            else:
                links_b = doc_to_matched_links.get(doc_b, [])
                similarity = calculate_jaccard_similarity(links_a, links_b)
                relevance_matrix[doc_a][doc_b] = round(similarity, 3)
    
    # 結果をまとめる
    result = {
        'metadata': {
            'extraction_date': datetime.now().isoformat(),
            'source_directory': str(excel_dir),
            'total_documents': len(documents),
            'total_matched_links': len(matched_links),
            'total_unmatched_links': len(unmatched_links),
            'relevance_calculation_method': 'jaccard_similarity',
            'subdirectories_searched': len(set([d['directory'] for d in documents]))
        },
        'documents': documents,
        'links': matched_links,
        'unmatched_links': unmatched_links,
        'relevance_matrix': relevance_matrix
    }
    
    return result


def save_results(result, output_dir="extraction_results"):
    """
    抽出結果を複数形式で保存
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 1. 完全なJSON（全情報）
    json_path = output_dir / f"document_graph_{timestamp}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n✓ Saved complete data: {json_path}")
    
    # 2. 簡易版JSON（グラフ構造のみ）
    graph = {}
    for link in result['links']:
        source = link['source']
        target = link['target']
        if source not in graph:
            graph[source] = []
        graph[source].append(target)
    
    simple_json = {
        'metadata': result['metadata'],
        'graph': graph
    }
    
    simple_path = output_dir / f"link_graph_{timestamp}.json"
    with open(simple_path, "w", encoding="utf-8") as f:
        json.dump(simple_json, f, ensure_ascii=False, indent=2)
    print(f"✓ Saved simplified graph: {simple_path}")
    
    # 3. CSVエクスポート（エッジリスト）
    import csv
    
    csv_path = output_dir / f"link_edges_{timestamp}.csv"
    with open(csv_path, "w", encoding="utf-8-sig", newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['source', 'target', 'original_text', 'match_type'])
        for link in result['links']:
            writer.writerow([
                link['source'],
                link['target'],
                link['original_text'],
                link['match_type']
            ])
    print(f"✓ Saved edge list: {csv_path}")
    
    # 4. 未マッチリストCSV
    if result['unmatched_links']:
        unmatched_path = output_dir / f"unmatched_links_{timestamp}.csv"
        with open(unmatched_path, "w", encoding="utf-8-sig", newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['source', 'original_text', 'normalized'])
            for link in result['unmatched_links']:
                writer.writerow([
                    link['source'],
                    link['original_text'],
                    link['normalized']
                ])
        print(f"✓ Saved unmatched links: {unmatched_path}")
    
    # 5. 関連度マトリクスCSV（共通情報ベース）
    if 'relevance_matrix' in result:
        relevance_path = output_dir / f"relevance_matrix_jaccard_{timestamp}.csv"
        doc_ids = sorted(result['relevance_matrix'].keys())
        
        with open(relevance_path, "w", encoding="utf-8-sig", newline='') as f:
            writer = csv.writer(f)
            # ヘッダー行
            writer.writerow([''] + doc_ids)
            # データ行
            for doc_a in doc_ids:
                row = [doc_a]
                for doc_b in doc_ids:
                    row.append(result['relevance_matrix'][doc_a][doc_b])
                writer.writerow(row)
        print(f"✓ Saved relevance matrix (Jaccard): {relevance_path}")
    
    # 6. サマリーレポート
    summary = {
        'total_documents': result['metadata']['total_documents'],
        'total_matched_links': result['metadata']['total_matched_links'],
        'total_unmatched_links': result['metadata']['total_unmatched_links'],
        'extraction_date': result['metadata']['extraction_date'],
        'documents_with_links': len([d for d in result['documents'] if d['extracted_links_count'] > 0]),
        'documents_without_links': len([d for d in result['documents'] if d['extracted_links_count'] == 0]),
    }
    
    # 関連度マトリクスの統計
    if 'relevance_matrix' in result:
        relevances = []
        doc_ids = list(result['relevance_matrix'].keys())
        for i, doc_a in enumerate(doc_ids):
            for j, doc_b in enumerate(doc_ids):
                if i < j:  # 上三角のみ（重複を避ける）
                    relevances.append(result['relevance_matrix'][doc_a][doc_b])
        
        if relevances:
            import statistics
            summary['relevance_stats'] = {
                'method': 'jaccard_similarity',
                'mean': round(statistics.mean(relevances), 3),
                'median': round(statistics.median(relevances), 3),
                'min': round(min(relevances), 3),
                'max': round(max(relevances), 3),
                'pairs_above_0.3': len([r for r in relevances if r >= 0.3]),
                'pairs_above_0.5': len([r for r in relevances if r >= 0.5]),
                'pairs_above_0.7': len([r for r in relevances if r >= 0.7])
            }
    
    summary_path = output_dir / f"summary_{timestamp}.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"✓ Saved summary: {summary_path}")
    
    # コンソールにサマリー表示
    print("\n" + "="*60)
    print("EXTRACTION SUMMARY")
    print("="*60)
    print(f"Source directory:          {result['metadata']['source_directory']}")
    print(f"Subdirectories searched:   {result['metadata'].get('subdirectories_searched', 1)}")
    print(f"Total documents:           {summary['total_documents']}")
    print(f"Documents with links:      {summary['documents_with_links']}")
    print(f"Documents without links:   {summary['documents_without_links']}")
    print(f"Matched links:             {summary['total_matched_links']}")
    print(f"Unmatched links:           {summary['total_unmatched_links']}")
    
    if 'relevance_stats' in summary:
        print("\nRelevance Statistics (Jaccard Similarity):")
        print(f"Mean relevance:            {summary['relevance_stats']['mean']}")
        print(f"Median relevance:          {summary['relevance_stats']['median']}")
        print(f"Pairs with relevance ≥0.3: {summary['relevance_stats']['pairs_above_0.3']}")
        print(f"Pairs with relevance ≥0.5: {summary['relevance_stats']['pairs_above_0.5']}")
        print(f"Pairs with relevance ≥0.7: {summary['relevance_stats']['pairs_above_0.7']}")
    
    print("="*60)
    
    return json_path, simple_path


def main():
    """メイン実行関数"""
    import sys
    
    if len(sys.argv) > 1:
        excel_dir = sys.argv[1]
    else:
        excel_dir = input("Enter directory path containing Excel files: ").strip()
    
    if not Path(excel_dir).exists():
        print(f"Error: Directory not found: {excel_dir}")
        return
    
    print("\n" + "="*60)
    print("DOCUMENT LINK EXTRACTOR")
    print("="*60)
    
    # 抽出実行
    result = build_document_graph(excel_dir)
    
    # 結果保存
    json_path, simple_path = save_results(result)
    
    print(f"\n✓ Extraction completed successfully!")
    print(f"\nMain output file: {json_path}")
    print(f"Use this JSON file to build the relevance matrix.")
    

if __name__ == "__main__":
    main()
