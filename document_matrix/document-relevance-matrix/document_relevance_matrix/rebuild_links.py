#!/usr/bin/env python3
"""
JSONファイルのdocuments部を元にlinks部を再計算して新規JSONファイルを生成
"""

import json
from pathlib import Path
from datetime import datetime
import re
import sys


def normalize_doc_name(name):
    """資料名を正規化（extract_links.pyと同じロジック）"""
    # 拡張子を削除
    name = re.sub(r'\.(xlsx?m?|xls|docx?|pdf)$', '', name, flags=re.IGNORECASE)
    
    # 末尾の _数字 パターン（バージョン番号）を削除
    name = re.sub(r'_\d+$', '', name)
    
    # 全角スペースを半角に、アンダースコアをスペースに変換
    name = name.replace('　', ' ').replace('_', ' ').strip()
    return name


def rebuild_links(input_json_path, output_json_path=None):
    """
    JSONファイルのdocuments部を元にlinks部を再構築
    
    Parameters:
        input_json_path: 入力JSONファイルパス（documents部が手動修正されたもの）
        output_json_path: 出力JSONファイルパス（省略時は自動生成）
    
    Returns:
        output_json_path: 生成されたJSONファイルのパス
    """
    input_path = Path(input_json_path)
    
    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        return None
    
    # JSONファイル読み込み
    print(f"Loading JSON file: {input_path}")
    with open(input_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    documents = data.get('documents', [])
    
    if not documents:
        print("Error: No documents found in JSON file")
        return None
    
    print(f"Found {len(documents)} documents")
    
    # ドキュメントIDと正規化名のマッピングを作成
    doc_id_to_normalized = {}
    for doc in documents:
        doc_id = doc.get('id')
        normalized = doc.get('normalized_name', normalize_doc_name(doc.get('filename', '')))
        doc_id_to_normalized[doc_id] = normalized
        print(f"  Document: '{doc_id}' → normalized: '{normalized}'")
    
    # すべての抽出されたリンクを収集
    all_link_texts = set()
    for doc in documents:
        extracted_links = doc.get('extracted_links', [])
        for link_text in extracted_links:
            all_link_texts.add(link_text)
    
    print(f"\nTotal unique link texts: {len(all_link_texts)}")
    
    # リンクテキストとターゲットドキュメントのマッピングを作成
    print("\nBuilding link mapping...")
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
            print(f"  ✓ '{link_text}' → Real doc: '{matched_real_doc}'")
        else:
            # 仮想ドキュメントとして扱う
            virtual_doc_id = normalized.replace(' ', '_')
            virtual_doc_mapping[link_text] = virtual_doc_id
            print(f"  ⚠️  '{link_text}' → Virtual doc: '{virtual_doc_id}'")
    
    # リンクを再構築
    print("\nRebuilding links...")
    new_links = []
    new_unmatched_links = []
    
    for doc in documents:
        source_id = doc.get('id')
        extracted_links = doc.get('extracted_links', [])
        
        for link_text in extracted_links:
            target_id = virtual_doc_mapping.get(link_text)
            
            if target_id:
                # マッチした場合
                match_type = "exact" if target_id in doc_id_to_normalized else "virtual"
                new_links.append({
                    "source": source_id,
                    "target": target_id,
                    "original_text": link_text,
                    "match_type": match_type
                })
                print(f"  Link: {source_id} → {target_id} (type: {match_type})")
            else:
                # マッチしない場合
                new_unmatched_links.append({
                    "source": source_id,
                    "original_text": link_text,
                    "normalized": normalize_doc_name(link_text)
                })
                print(f"  Unmatched: {source_id} → '{link_text}'")
    
    # 統計情報
    matched_count = len(new_links)
    unmatched_count = len(new_unmatched_links)
    total_count = matched_count + unmatched_count
    
    print("\n" + "="*60)
    print("LINK REBUILD SUMMARY")
    print("="*60)
    print(f"Total links processed:     {total_count}")
    print(f"  - Matched links:         {matched_count}")
    print(f"  - Unmatched links:       {unmatched_count}")
    print(f"Match rate:                {matched_count/max(1, total_count)*100:.1f}%")
    print("="*60)
    
    # 新しいJSONデータを構築
    new_data = {
        "metadata": {
            "rebuild_date": datetime.now().isoformat(),
            "source_file": str(input_path),
            "original_extraction_date": data.get('metadata', {}).get('extraction_date', 'N/A'),
            "source_directory": data.get('metadata', {}).get('source_directory', 'N/A'),
            "total_documents": len(documents),
            "total_matched_links": matched_count,
            "total_unmatched_links": unmatched_count,
            "subdirectories_searched": data.get('metadata', {}).get('subdirectories_searched', 1)
        },
        "documents": documents,  # documents部はそのまま使用
        "links": new_links,
        "unmatched_links": new_unmatched_links
    }
    
    # 出力ファイル名を決定
    if output_json_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_json_path = input_path.parent / f"links_rebuilt_{timestamp}.json"
    else:
        output_json_path = Path(output_json_path)
    
    # JSONファイルに保存
    print(f"\nSaving rebuilt JSON to: {output_json_path}")
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(new_data, f, ensure_ascii=False, indent=2)
    
    print(f"✓ Successfully saved: {output_json_path}")
    
    # 変更点のサマリー
    if 'links' in data:
        original_matched = data['metadata'].get('total_matched_links', 0)
        original_unmatched = data['metadata'].get('total_unmatched_links', 0)
        
        print("\n" + "="*60)
        print("CHANGES FROM ORIGINAL")
        print("="*60)
        print(f"Matched links:    {original_matched} → {matched_count} (Δ{matched_count - original_matched:+d})")
        print(f"Unmatched links:  {original_unmatched} → {unmatched_count} (Δ{unmatched_count - original_unmatched:+d})")
        print("="*60)
    
    return output_json_path


def select_file_gui():
    """GUIでファイル選択ダイアログを表示"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        root = tk.Tk()
        root.withdraw()  # メインウィンドウを非表示
        root.attributes('-topmost', True)  # ダイアログを最前面に
        
        print("ファイル選択ダイアログを開いています...")
        file_path = filedialog.askopenfilename(
            title="JSONファイルを選択してください",
            filetypes=[
                ("JSON files", "*.json"),
                ("All files", "*.*")
            ],
            initialdir="."
        )
        
        root.destroy()
        
        if file_path:
            return file_path
        else:
            print("ファイルが選択されませんでした。")
            return None
    
    except ImportError:
        print("Error: tkinter がインストールされていません。")
        print("コマンドライン引数でファイルパスを指定してください。")
        return None


def main():
    """メイン実行関数"""
    # コマンドライン引数をチェック
    if len(sys.argv) >= 2:
        # 引数が指定されている場合
        input_json = sys.argv[1]
        output_json = sys.argv[2] if len(sys.argv) > 2 else None
    else:
        # 引数がない場合はGUIで選択
        print("\n" + "="*60)
        print("DOCUMENT LINKS REBUILDER - FILE SELECTION")
        print("="*60)
        print("コマンドライン引数が指定されていません。")
        print("GUIでファイルを選択します...\n")
        
        input_json = select_file_gui()
        
        if input_json is None:
            print("\nUsage: python3 rebuild_links.py [input_json] [output_json]")
            print("\nDescription:")
            print("  JSONファイルのdocuments部を元にlinks部を再構築します。")
            print("  引数を省略するとGUIでファイルを選択できます。")
            print("\nExamples:")
            print("  # GUIでファイル選択（引数なし）")
            print("  python3 rebuild_links.py")
            print("\n  # 自動で出力ファイル名を生成")
            print("  python3 rebuild_links.py links_extracted_20260107_123456.json")
            print("\n  # 出力ファイル名を指定")
            print("  python3 rebuild_links.py input.json output.json")
            print("\nWorkflow:")
            print("  1. extract_links.py でリンク抽出 → links_extracted_*.json")
            print("  2. JSONファイルのdocuments部を手動編集（テキストエディタで）")
            print("  3. このスクリプトでlinks部を再生成")
            print("  4. 生成されたJSONファイルでbuild_matrix.pyを実行")
            return
        
        output_json = None
    
    print("\n" + "="*60)
    print("DOCUMENT LINKS REBUILDER")
    print("="*60)
    print(f"Input:  {input_json}")
    if output_json:
        print(f"Output: {output_json}")
    else:
        print(f"Output: (auto-generated)")
    print("="*60 + "\n")
    
    # リンク再構築実行
    result_path = rebuild_links(input_json, output_json)
    
    if result_path:
        print("\n" + "="*60)
        print("✓ Link rebuild completed successfully!")
        print(f"Output: {result_path}")
        print("="*60)
        print("\n💡 Next steps:")
        print(f"  1. Review the rebuilt links in: {result_path}")
        print(f"  2. Run: uv run calculate-relevance {result_path}")
    else:
        print("\n❌ Link rebuild failed!")


if __name__ == "__main__":
    main()
