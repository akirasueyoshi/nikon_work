#!/usr/bin/env python3
"""
エクセルファイルから資料間のリンク情報を抽出（関連度計算なし）
"""

import pandas as pd
from pathlib import Path
import json
from datetime import datetime
import re


def extract_links_from_excel(excel_path):
    """
    エクセルファイルから機能仕様書名を抽出
    
    Returns:
        list: 抽出された仕様書名のリスト
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=0, header=None)
        links = []
        
        # B列(index=1)とC列(index=2)をチェック
        for col_idx in [1, 2]:
            if col_idx >= len(df.columns):
                continue
            
            col = df.iloc[:, col_idx]
            start_idx = None
            end_idx = None
            
            for idx, cell in enumerate(col):
                if pd.notna(cell):
                    cell_str = str(cell).strip()
                    
                    if '機能仕様書' in cell_str:
                        start_idx = idx + 1
                        continue
                    
                    if start_idx is not None and '対応内容' in cell_str:
                        end_idx = idx
                        break
            
            if start_idx is not None:
                if end_idx is None:
                    end_idx = len(col)
                
                for idx in range(start_idx, end_idx):
                    if idx >= len(col):
                        break
                    
                    cell = col.iloc[idx]
                    if pd.notna(cell):
                        cell_str = str(cell).strip()

                        # "01_機能仕様_"を除去
                        cell_str = cell_str.replace("01_機能仕様_", "")

                        if cell_str and not cell_str.replace('.', '').replace('-', '').isdigit():
                            if cell_str not in ['NaN', 'nan', '']:
                                links.append(cell_str)
        
        links = list(dict.fromkeys(links))
        return links
    
    except Exception as e:
        print(f"Error reading {excel_path.name}: {e}")
        return []


def normalize_doc_name(name):
    """資料名を正規化"""
    # 拡張子を削除（.xlsx, .xls, .xlsm, .doc, .docx, .pdf など）
    name = re.sub(r'\.(xlsx?m?|xls|docx?|pdf|eap)$', '', name, flags=re.IGNORECASE)
    
    # 末尾の _数字 パターン（バージョン番号）を削除
    # 例: "document_v1" → "document", "spec_2" → "spec"
    name = re.sub(r'_\d+$', '', name)
    
    # 全角スペースを半角に、アンダースコアをスペースに変換
    name = name.replace('　', ' ').replace('_', ' ').strip()
    return name


def build_document_graph(excel_dir):
    """
    指定ディレクトリ内のExcelファイルからリンク情報を抽出
    """
    excel_dir = Path(excel_dir)
    
    # 再帰的にExcelファイルを探索
    print(f"Searching for Excel files in {excel_dir} (recursively)...")
    excel_files = sorted([
        f for f in excel_dir.glob('**/*.xlsx')
        if not f.name.startswith('~$') and not f.name.startswith('.')
    ])
    
    print(f"Found {len(excel_files)} Excel files")
    
    # ディレクトリ構造を表示
    dirs = {}
    for f in excel_files:
        rel_dir = f.parent.relative_to(excel_dir)
        dir_str = str(rel_dir) if str(rel_dir) != '.' else '.'
        dirs[dir_str] = dirs.get(dir_str, 0) + 1
    
    print("Directory structure:")
    for dir_name, count in sorted(dirs.items()):
        print(f"  {dir_name}: {count} file(s)")
    
    # ドキュメント情報を構築
    documents = []
    doc_id_to_normalized = {}
    
    for excel_file in excel_files:
        doc_id = excel_file.stem
        normalized_name = normalize_doc_name(doc_id)
        
        rel_path = excel_file.relative_to(excel_dir)
        directory = str(rel_path.parent) if str(rel_path.parent) != '.' else '.'
        
        print(f"Processing: {rel_path}")
        extracted_links = extract_links_from_excel(excel_file)
        print(f"  -> Found {len(extracted_links)} links")
        
        doc_info = {
            "id": doc_id,
            "filename": excel_file.name,
            "path": str(excel_file),
            "relative_path": str(rel_path),
            "directory": directory,
            "normalized_name": normalized_name,
            "extracted_links_count": len(extracted_links),
            "extracted_links": extracted_links
        }
        
        documents.append(doc_info)
        doc_id_to_normalized[doc_id] = normalized_name
    
    # リンクをマッチング（実在するファイルと仮想ドキュメント）
    print("\nMatching links to documents...")
    print("\nAvailable document IDs (normalized):")
    for doc_id, normalized in sorted(doc_id_to_normalized.items()):
        print(f"  '{doc_id}' → '{normalized}'")
    
    # すべての抽出されたリンクを収集（仮想ドキュメントも含む）
    all_link_texts = set()
    for doc in documents:
        for link_text in doc['extracted_links']:
            all_link_texts.add(link_text)
    
    # 仮想ドキュメントIDマッピングを作成
    virtual_doc_mapping = {}
    for link_text in all_link_texts:
        normalized = normalize_doc_name(link_text)
        # 実在するドキュメントとマッチするか確認
        matched_real_doc = None
        for doc_id, doc_normalized in doc_id_to_normalized.items():
            if normalized == doc_normalized:
                matched_real_doc = doc_id
                break
        
        if matched_real_doc:
            virtual_doc_mapping[link_text] = matched_real_doc
        else:
            # 仮想ドキュメントとして扱う
            virtual_doc_id = normalized.replace(' ', '_')  # 正規化名をIDに
            virtual_doc_mapping[link_text] = virtual_doc_id
    
    links = []
    unmatched_links = []
    
    print("\n📝 Link mapping:")
    for doc in documents:
        source_id = doc['id']
        
        for link_text in doc['extracted_links']:
            target_id = virtual_doc_mapping.get(link_text)
            
            if target_id:
                links.append({
                    "source": source_id,
                    "target": target_id,
                    "original_text": link_text,
                    "match_type": "exact" if target_id in doc_id_to_normalized else "virtual"
                })
                print(f"  ✓ {source_id} → {target_id} ({link_text})")
            else:
                unmatched_links.append({
                    "source": source_id,
                    "original_text": link_text,
                    "normalized": normalize_doc_name(link_text)
                })
                print(f"  ⚠️  Unmatched: '{link_text}' → '{normalize_doc_name(link_text)}'")
    
    print(f"  Matched: {len(links)} links")
    print(f"  Unmatched: {len(unmatched_links)} links")
    
    # 結果を構築
    result = {
        "metadata": {
            "extraction_date": datetime.now().isoformat(),
            "source_directory": str(excel_dir),
            "total_documents": len(documents),
            "total_matched_links": len(links),
            "total_unmatched_links": len(unmatched_links),
            "subdirectories_searched": len(dirs)
        },
        "documents": documents,
        "links": links,
        "unmatched_links": unmatched_links
    }
    
    return result


def save_results(result, output_dir="extraction_results"):
    """リンク抽出結果のみを保存"""
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # リンク抽出結果JSON（編集可能な形式）
    json_path = output_dir / f"links_extracted_{timestamp}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n✓ Saved extraction results: {json_path}")
    
    # コンソールにサマリー表示
    print("\n" + "="*60)
    print("LINK EXTRACTION SUMMARY")
    print("="*60)
    print(f"Source directory:          {result['metadata']['source_directory']}")
    print(f"Subdirectories searched:   {result['metadata'].get('subdirectories_searched', 1)}")
    print(f"Total documents:           {result['metadata']['total_documents']}")
    print(f"Documents with links:      {len([d for d in result['documents'] if d['extracted_links_count'] > 0])}")
    print(f"Documents without links:   {len([d for d in result['documents'] if d['extracted_links_count'] == 0])}")
    print(f"Matched links:             {result['metadata']['total_matched_links']}")
    print(f"Unmatched links:           {result['metadata']['total_unmatched_links']}")
    print("="*60)
    
    print(f"\n📝 Review the extracted links in:")
    print(f"   {json_path}")
    print(f"\n💡 To edit links manually:")
    print(f"   1. Open {json_path} in a text editor")
    print(f"   2. Modify the 'links' and 'unmatched_links' sections")
    print(f"   3. Run: uv run calculate-relevance {json_path}")
    
    return json_path


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
    
    # リンク抽出のみ実行
    result = build_document_graph(excel_dir)
    
    # 結果保存
    json_path = save_results(result)
    
    print(f"\n✓ Link extraction completed successfully!")
    

if __name__ == "__main__":
    main()
