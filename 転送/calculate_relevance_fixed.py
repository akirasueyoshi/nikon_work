#!/usr/bin/env python3
"""
calculate-relevanceの正しい実装
extracted_linksから直接Jaccard係数を計算（linksセクションは使わない）
"""

import json
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime
import re

def select_json_file():
    """GUIでJSONファイルを選択"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    file_path = filedialog.askopenfilename(
        title='リンク抽出JSONファイルを選択',
        filetypes=[('JSON files', '*.json'), ('All files', '*.*')],
        initialdir=os.getcwd()
    )
    
    root.destroy()
    
    if not file_path:
        print("ファイルが選択されませんでした。")
        return None
    
    return file_path

def normalize_link(link_text):
    """
    リンクテキストを正規化
    extract_links.pyのnormalize_doc_name()と同じロジック
    """
    # 拡張子を削除
    link = re.sub(r'\.(xlsx?m?|xls|docx?|pdf|eap)$', '', link_text, flags=re.IGNORECASE)
    
    # 末尾の _数字 パターンを削除
    link = re.sub(r'_\d+$', '', link)
    
    # 全角スペースを半角に、アンダースコアをスペースに変換
    link = link.replace('　', ' ').replace('_', ' ').strip()
    
    return link

def calculate_jaccard_matrix_from_extracted_links(data):
    """
    documentsセクションのextracted_linksから直接Jaccard係数を計算
    linksセクションは一切使用しない
    """
    documents = data['documents']
    n = len(documents)
    
    # ドキュメントIDリスト
    doc_ids = [doc['id'] for doc in documents]
    
    print(f"\nJaccard係数を計算中（extracted_linksを直接使用）...")
    
    # 各ドキュメントのextracted_linksを正規化してセット化
    doc_links = []
    for i, doc in enumerate(documents):
        normalized_links = set(normalize_link(link) for link in doc['extracted_links'])
        doc_links.append(normalized_links)
        
        if (i + 1) % 50 == 0:
            print(f"  処理中: {i+1}/{n}")
    
    # Jaccard係数マトリクスを計算
    matrix = np.zeros((n, n))
    
    total_pairs = n * (n - 1) // 2
    processed = 0
    
    for i in range(n):
        matrix[i][i] = 1.0  # 対角成分
        
        for j in range(i+1, n):
            set1 = doc_links[i]
            set2 = doc_links[j]
            
            intersection = len(set1 & set2)
            union = len(set1 | set2)
            
            if union > 0:
                jaccard = intersection / union
            else:
                jaccard = 0.0
            
            matrix[i][j] = jaccard
            matrix[j][i] = jaccard  # 対称行列
            
            processed += 1
            if processed % 10000 == 0:
                print(f"  ペア処理: {processed}/{total_pairs}")
    
    df = pd.DataFrame(matrix, index=doc_ids, columns=doc_ids)
    return df

def create_output_directory():
    """出力用ディレクトリを作成"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    top_level_dir = os.path.join(script_dir, "relevance_matrices")
    os.makedirs(top_level_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = os.path.join(top_level_dir, timestamp)
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"✓ 出力ディレクトリ作成: relevance_matrices/{timestamp}")
    return output_dir

def save_matrix(df, output_dir, original_filename):
    """マトリクスをCSVとして保存"""
    base_name = os.path.splitext(original_filename)[0]
    output_filename = f"{base_name}_relevance_jaccard.csv"
    output_path = os.path.join(output_dir, output_filename)
    
    df.to_csv(output_path, encoding='utf-8-sig')
    print(f"✓ Jaccard係数マトリクス保存: {output_filename}")
    return output_path

def save_statistics(df, data, output_dir, original_filename):
    """統計情報を保存"""
    base_name = os.path.splitext(original_filename)[0]
    stats_filename = f"{base_name}_statistics.txt"
    stats_path = os.path.join(output_dir, stats_filename)
    
    with open(stats_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write("実ファイル間関連度（Jaccard係数）マトリクス - 統計情報\n")
        f.write("="*60 + "\n\n")
        f.write(f"処理日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n")
        f.write(f"元ファイル: {original_filename}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("計算方法\n")
        f.write("-"*60 + "\n")
        f.write("extracted_linksを正規化してJaccard係数を計算\n")
        f.write("linksセクションは使用せず、extracted_linksを直接使用\n")
        f.write("この方法により、ファイル追加による影響を受けない\n\n")
        
        f.write("-"*60 + "\n")
        f.write("入力データ統計\n")
        f.write("-"*60 + "\n")
        f.write(f"ドキュメント数: {data['metadata']['total_documents']}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("Jaccard係数マトリクス統計\n")
        f.write("-"*60 + "\n")
        f.write(f"マトリクスサイズ: {df.shape[0]} x {df.shape[1]}\n")
        
        # 対角成分を除いた統計
        values = df.values[np.triu_indices_from(df.values, k=1)]
        f.write(f"最大類似度: {values.max():.6f}\n")
        f.write(f"平均類似度: {values.mean():.6f}\n")
        f.write(f"中央値: {np.median(values):.6f}\n")
        f.write(f"類似度0のペア数: {(values == 0).sum()} / {len(values)}\n")
        f.write(f"類似度0.5以上のペア数: {(values >= 0.5).sum()} / {len(values)}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("類似度が高いトップ10ペア\n")
        f.write("-"*60 + "\n")
        
        # 上三角行列から値を取得
        pairs = []
        for i in range(len(df)):
            for j in range(i+1, len(df)):
                if df.iloc[i, j] > 0:
                    pairs.append((df.index[i], df.columns[j], df.iloc[i, j]))
        
        pairs.sort(key=lambda x: x[2], reverse=True)
        for rank, (doc1, doc2, jaccard) in enumerate(pairs[:10], 1):
            f.write(f"{rank:2d}. Jaccard: {jaccard:.6f}\n")
            f.write(f"    {doc1}\n")
            f.write(f"    ⟷ {doc2}\n\n")
    
    print(f"✓ 統計情報保存: {stats_filename}")
    return stats_path

def save_summary(output_dir, original_filename, matrix_size, total_docs):
    """処理結果のサマリーを保存"""
    summary_path = os.path.join(output_dir, "README.txt")
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write("実ファイル間関連度（Jaccard係数）マトリクス\n")
        f.write("="*60 + "\n\n")
        f.write(f"処理日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n")
        f.write(f"元ファイル: {original_filename}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("Jaccard係数の定義\n")
        f.write("-"*60 + "\n")
        f.write("2つのドキュメントのextracted_linksの類似度を計算。\n")
        f.write("Jaccard係数 = |A ∩ B| / |A ∪ B|\n")
        f.write("  A, B: 各ドキュメントのextracted_links（正規化後）\n")
        f.write("  値の範囲: 0.0～1.0\n")
        f.write("  1.0: 完全に同じリンクを持つ\n")
        f.write("  0.0: 共通のリンクが全くない\n\n")
        
        f.write("-"*60 + "\n")
        f.write("重要な実装詳細\n")
        f.write("-"*60 + "\n")
        f.write("・linksセクションは使用せず、extracted_linksを直接使用\n")
        f.write("・リンクテキストをextract_links.pyと同じロジックで正規化\n")
        f.write("・ファイル追加による既存ドキュメント間の類似度への影響なし\n\n")
        
        f.write("-"*60 + "\n")
        f.write("出力ファイル\n")
        f.write("-"*60 + "\n")
        f.write("1. *_relevance_jaccard.csv\n")
        f.write("   Jaccard係数マトリクス（対称行列、対角=1.0）\n\n")
        f.write("2. *_statistics.txt\n")
        f.write("   統計情報とトップランキング\n\n")
        
        f.write("-"*60 + "\n")
        f.write("マトリクス情報\n")
        f.write("-"*60 + "\n")
        f.write(f"マトリクスサイズ: {matrix_size} x {matrix_size}\n")
        f.write(f"ドキュメント数: {total_docs}\n")
    
    print(f"✓ サマリー保存: README.txt")
    return summary_path

def main():
    """メイン処理"""
    print("="*60)
    print("実ファイル間関連度（Jaccard係数）計算ツール")
    print("="*60)
    print("\nこのツールはextracted_linksから直接Jaccard係数を計算します。")
    print("linksセクションは使用しないため、ファイル追加の影響を受けません。")
    
    # ファイル選択
    print("\nリンク抽出JSONファイルを選択してください...")
    json_path = select_json_file()
    
    if not json_path:
        return
    
    original_filename = os.path.basename(json_path)
    print(f"\n選択されたファイル: {original_filename}")
    
    # JSONファイル読み込み
    print("\nJSONファイルを読み込み中...")
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"✓ 読み込み成功")
        print(f"  - ドキュメント数: {data['metadata']['total_documents']}")
    except Exception as e:
        print(f"✗ エラー: {e}")
        return
    
    # Jaccard係数マトリクスを計算
    df = calculate_jaccard_matrix_from_extracted_links(data)
    print(f"✓ Jaccard係数マトリクス作成完了: {df.shape[0]} x {df.shape[1]}")
    
    # 出力ディレクトリ作成
    output_dir = create_output_directory()
    
    # ファイル保存
    print("\n" + "-"*60)
    save_matrix(df, output_dir, original_filename)
    save_statistics(df, data, output_dir, original_filename)
    save_summary(output_dir, original_filename, df.shape[0], 
                 data['metadata']['total_documents'])
    
    print("\n" + "="*60)
    print("完了！")
    print("="*60)
    print(f"\n出力ディレクトリ: {output_dir}")
    print("\n生成されたファイル:")
    print("  ├─ README.txt")
    print("  ├─ *_relevance_jaccard.csv (Jaccard係数マトリクス)")
    print("  └─ *_statistics.txt (統計情報)")
    print("\nこのマトリクスはファイル追加の影響を受けません。")
    print("="*60)

if __name__ == "__main__":
    main()