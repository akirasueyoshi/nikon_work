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


def build_document_graph(excel_dir):
    """
    全エクセルファイルからドキュメントグラフを構築
    
    Returns:
        dict: {
            'documents': [資料情報のリスト],
            'links': [リンク情報のリスト],
            'metadata': {メタデータ}
        }
    """
    excel_dir = Path(excel_dir)
    
    # 全エクセルファイルを取得（一時ファイルを除外）
    excel_files = []
    for pattern in ['*.xlsx', '*.xls']:
        excel_files.extend([f for f in excel_dir.glob(pattern) if not f.name.startswith('~$')])
    
    excel_files = sorted(set(excel_files))  # 重複除去・ソート
    
    print(f"Found {len(excel_files)} Excel files in {excel_dir}")
    
    # ドキュメント情報を収集
    documents = []
    raw_links = {}  # {source_doc: [linked_docs]}
    
    for excel_file in excel_files:
        doc_name = excel_file.stem
        print(f"Processing: {doc_name}")
        
        # リンク先を抽出
        links = extract_links_from_excel(excel_file)
        
        # ドキュメント情報を記録
        doc_info = {
            'id': doc_name,
            'filename': excel_file.name,
            'path': str(excel_file),
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
    
    for source_doc, link_list in raw_links.items():
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
                            matched = True
                            break
            
            if not matched:
                unmatched_links.append({
                    'source': source_doc,
                    'original_text': link,
                    'normalized': normalized_link
                })
    
    # 結果をまとめる
    result = {
        'metadata': {
            'extraction_date': datetime.now().isoformat(),
            'source_directory': str(excel_dir),
            'total_documents': len(documents),
            'total_matched_links': len(matched_links),
            'total_unmatched_links': len(unmatched_links)
        },
        'documents': documents,
        'links': matched_links,
        'unmatched_links': unmatched_links
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
    
    # 5. サマリーレポート
    summary = {
        'total_documents': result['metadata']['total_documents'],
        'total_matched_links': result['metadata']['total_matched_links'],
        'total_unmatched_links': result['metadata']['total_unmatched_links'],
        'extraction_date': result['metadata']['extraction_date'],
        'documents_with_links': len([d for d in result['documents'] if d['extracted_links_count'] > 0]),
        'documents_without_links': len([d for d in result['documents'] if d['extracted_links_count'] == 0]),
    }
    
    summary_path = output_dir / f"summary_{timestamp}.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"✓ Saved summary: {summary_path}")
    
    # コンソールにサマリー表示
    print("\n" + "="*60)
    print("EXTRACTION SUMMARY")
    print("="*60)
    print(f"Total documents:           {summary['total_documents']}")
    print(f"Documents with links:      {summary['documents_with_links']}")
    print(f"Documents without links:   {summary['documents_without_links']}")
    print(f"Matched links:             {summary['total_matched_links']}")
    print(f"Unmatched links:           {summary['total_unmatched_links']}")
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
