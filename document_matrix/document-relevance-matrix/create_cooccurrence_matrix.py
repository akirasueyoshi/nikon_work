import json
import pandas as pd
import numpy as np
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime

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

def load_links_json(json_path):
    """JSONファイルを読み込み"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"✓ JSONファイル読み込み成功")
        print(f"  - ドキュメント数: {data['metadata']['total_documents']}")
        print(f"  - リンク総数: {data['metadata']['total_matched_links']}")
        return data
    except Exception as e:
        print(f"✗ エラー: {e}")
        return None

def extract_all_link_targets(data):
    """全てのリンクターゲット（extracted_links内の参照先）を抽出"""
    link_targets = set()
    
    for doc in data['documents']:
        for link in doc['extracted_links']:
            # 拡張子を除去して正規化
            target_name = link.replace('.xlsx', '').replace('.xls', '')
            link_targets.add(target_name)
    
    # ソート済みリストとして返す
    return sorted(list(link_targets))

def build_cooccurrence_matrix(data):
    """
    同時出現行列を構築
    各ドキュメント内で一緒に記載されているリンクの出現回数をカウント
    """
    # 全てのリンクターゲットを取得
    all_targets = extract_all_link_targets(data)
    print(f"\n✓ 検出されたリンクターゲット数: {len(all_targets)}")
    
    # 同時出現カウント用の辞書
    cooccurrence = defaultdict(lambda: defaultdict(int))
    
    # 各ドキュメントを処理
    for doc in data['documents']:
        # このドキュメント内のリンクを正規化
        links_in_doc = [link.replace('.xlsx', '').replace('.xls', '') 
                        for link in doc['extracted_links']]
        
        # ドキュメント内の全てのリンクペアについて同時出現をカウント
        for i, link1 in enumerate(links_in_doc):
            for link2 in links_in_doc[i+1:]:  # 自分自身とのペアは除外
                # 対称的にカウント
                cooccurrence[link1][link2] += 1
                cooccurrence[link2][link1] += 1
    
    return all_targets, cooccurrence

def create_cooccurrence_dataframe(all_targets, cooccurrence):
    """同時出現行列をDataFrameに変換"""
    n = len(all_targets)
    matrix = np.zeros((n, n), dtype=int)
    
    for i, target1 in enumerate(all_targets):
        for j, target2 in enumerate(all_targets):
            if i == j:
                matrix[i][j] = 0  # 対角成分は0
            else:
                matrix[i][j] = cooccurrence[target1][target2]
    
    df = pd.DataFrame(matrix, index=all_targets, columns=all_targets)
    return df

def print_sample_analysis(df, sample_target='機能仕様書A'):
    """サンプル分析結果を表示"""
    print(f"\n{'='*60}")
    print(f"{sample_target} と他のドキュメントとの関連度")
    print('='*60)
    
    if sample_target in df.index:
        relationships = df.loc[sample_target].sort_values(ascending=False)
        relationships = relationships[relationships > 0]  # 0より大きいもののみ
        
        for target, count in relationships.items():
            print(f"  ・{target}: {count}")
    else:
        print(f"  {sample_target} はマトリクスに含まれていません")

def create_output_directory():
    """出力用ディレクトリを作成（スクリプトと同じディレクトリ）"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    top_level_dir = os.path.join(script_dir, "cooccurrence_matrices")
    os.makedirs(top_level_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = os.path.join(top_level_dir, timestamp)
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"✓ 出力ディレクトリ作成: cooccurrence_matrices/{timestamp}")
    return output_dir

def save_matrix(df, output_dir, original_filename):
    """マトリクスをCSVとして保存"""
    base_name = os.path.splitext(original_filename)[0]
    output_filename = f"{base_name}_cooccurrence_matrix.csv"
    output_path = os.path.join(output_dir, output_filename)
    
    df.to_csv(output_path, encoding='utf-8-sig')
    print(f"✓ マトリクス保存: {output_filename}")
    return output_path

def save_statistics(df, data, output_dir, original_filename):
    """統計情報を保存"""
    base_name = os.path.splitext(original_filename)[0]
    stats_filename = f"{base_name}_statistics.txt"
    stats_path = os.path.join(output_dir, stats_filename)
    
    with open(stats_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write("ドキュメント同時出現関連度マトリクス - 統計情報\n")
        f.write("="*60 + "\n\n")
        f.write(f"処理日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n")
        f.write(f"元ファイル: {original_filename}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("入力データ統計\n")
        f.write("-"*60 + "\n")
        f.write(f"ソースドキュメント数: {data['metadata']['total_documents']}\n")
        f.write(f"リンク総数: {data['metadata']['total_matched_links']}\n")
        f.write(f"ユニークなリンクターゲット数: {len(df)}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("関連度マトリクス統計\n")
        f.write("-"*60 + "\n")
        f.write(f"マトリクスサイズ: {df.shape[0]} x {df.shape[1]}\n")
        
        # 対角成分を除いた統計
        values = df.values[np.triu_indices_from(df.values, k=1)]
        f.write(f"最大関連度: {values.max()}\n")
        f.write(f"平均関連度: {values.mean():.2f}\n")
        f.write(f"関連度0のペア数: {(values == 0).sum()} / {len(values)}\n")
        f.write(f"関連度1以上のペア数: {(values > 0).sum()} / {len(values)}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("関連度が高いトップ10ペア\n")
        f.write("-"*60 + "\n")
        
        # 上三角行列から値を取得
        pairs = []
        for i in range(len(df)):
            for j in range(i+1, len(df)):
                if df.iloc[i, j] > 0:
                    pairs.append((df.index[i], df.columns[j], df.iloc[i, j]))
        
        pairs.sort(key=lambda x: x[2], reverse=True)
        for rank, (doc1, doc2, count) in enumerate(pairs[:10], 1):
            f.write(f"{rank:2d}. {doc1} ⟷ {doc2}: {count}\n")
        
        f.write("\n" + "-"*60 + "\n")
        f.write("各ドキュメントの関連度合計（降順トップ10）\n")
        f.write("-"*60 + "\n")
        
        total_relations = df.sum(axis=1).sort_values(ascending=False)
        for rank, (doc, total) in enumerate(total_relations.head(10).items(), 1):
            f.write(f"{rank:2d}. {doc}: {int(total)}\n")
    
    print(f"✓ 統計情報保存: {stats_filename}")
    return stats_path

def save_detailed_relationships(df, data, output_dir, original_filename):
    """各ドキュメントの詳細な関連情報を保存"""
    base_name = os.path.splitext(original_filename)[0]
    detail_filename = f"{base_name}_detailed_relationships.txt"
    detail_path = os.path.join(output_dir, detail_filename)
    
    with open(detail_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write("各ドキュメントの詳細関連度\n")
        f.write("="*60 + "\n\n")
        
        for target in df.index:
            f.write("-"*60 + "\n")
            f.write(f"{target}\n")
            f.write("-"*60 + "\n")
            
            relationships = df.loc[target].sort_values(ascending=False)
            relationships = relationships[relationships > 0]
            
            if len(relationships) > 0:
                for related_doc, count in relationships.items():
                    f.write(f"  ・{related_doc}: {count}\n")
            else:
                f.write("  （関連するドキュメントなし）\n")
            f.write("\n")
    
    print(f"✓ 詳細関連度保存: {detail_filename}")
    return detail_path

def save_summary(output_dir, original_filename, matrix_size, total_docs):
    """処理結果のサマリーを保存"""
    summary_path = os.path.join(output_dir, "README.txt")
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write("ドキュメント同時出現関連度マトリクス\n")
        f.write("="*60 + "\n\n")
        f.write(f"処理日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n")
        f.write(f"元ファイル: {original_filename}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("関連度の定義\n")
        f.write("-"*60 + "\n")
        f.write("同じソースドキュメント内に同時に記載された回数を関連度として定義。\n")
        f.write("例: 機能仕様書Aと機能仕様書Bが3つのドキュメントで同時に\n")
        f.write("    参照されている場合、関連度は3となる。\n\n")
        
        f.write("-"*60 + "\n")
        f.write("出力ファイル\n")
        f.write("-"*60 + "\n")
        f.write("1. *_cooccurrence_matrix.csv\n")
        f.write("   関連度マトリクス（対称行列）\n\n")
        f.write("2. *_statistics.txt\n")
        f.write("   統計情報とトップランキング\n\n")
        f.write("3. *_detailed_relationships.txt\n")
        f.write("   各ドキュメントの詳細関連度リスト\n\n")
        
        f.write("-"*60 + "\n")
        f.write("マトリクス情報\n")
        f.write("-"*60 + "\n")
        f.write(f"マトリクスサイズ: {matrix_size} x {matrix_size}\n")
        f.write(f"ソースドキュメント数: {total_docs}\n")
        f.write("\n対角成分は0（自分自身との関連度は定義しない）\n")
        f.write("対称行列（A→Bの関連度 = B→Aの関連度）\n")
    
    print(f"✓ サマリー保存: README.txt")
    return summary_path

def main():
    """メイン処理"""
    print("="*60)
    print("ドキュメント同時出現関連度マトリクス作成ツール")
    print("="*60)
    
    # ファイル選択
    print("\nリンク抽出JSONファイルを選択してください...")
    json_path = select_json_file()
    
    if not json_path:
        return
    
    original_filename = os.path.basename(json_path)
    print(f"\n選択されたファイル: {original_filename}")
    
    # JSONファイル読み込み
    data = load_links_json(json_path)
    if data is None:
        return
    
    # 同時出現行列を構築
    print("\n同時出現行列を構築中...")
    all_targets, cooccurrence = build_cooccurrence_matrix(data)
    
    # DataFrameに変換
    df = create_cooccurrence_dataframe(all_targets, cooccurrence)
    print(f"✓ マトリクス作成完了: {df.shape[0]} x {df.shape[1]}")
    
    # サンプル分析表示
    if '機能仕様書A' in df.index:
        print_sample_analysis(df, '機能仕様書A')
    
    # 出力ディレクトリ作成
    output_dir = create_output_directory()
    
    # ファイル保存
    print("\n" + "-"*60)
    save_matrix(df, output_dir, original_filename)
    save_statistics(df, data, output_dir, original_filename)
    save_detailed_relationships(df, data, output_dir, original_filename)
    save_summary(output_dir, original_filename, df.shape[0], 
                 data['metadata']['total_documents'])
    
    print("\n" + "="*60)
    print("完了！")
    print("="*60)
    print(f"\n出力ディレクトリ: {output_dir}")
    print("\n生成されたファイル:")
    print("  ├─ README.txt")
    print("  ├─ *_cooccurrence_matrix.csv (関連度マトリクス)")
    print("  ├─ *_statistics.txt (統計情報)")
    print("  └─ *_detailed_relationships.txt (詳細関連度)")
    print("="*60)

if __name__ == "__main__":
    main()