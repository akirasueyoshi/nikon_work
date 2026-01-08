import pandas as pd
import numpy as np
from scipy.cluster import hierarchy
from scipy.spatial.distance import squareform
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime

def select_csv_file():
    """GUIでCSVファイルを選択"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    file_path = filedialog.askopenfilename(
        title='類似度マトリクスCSVファイルを選択',
        filetypes=[('CSV files', '*.csv'), ('All files', '*.*')],
        initialdir=os.getcwd()
    )
    
    root.destroy()
    
    if not file_path:
        print("ファイルが選択されませんでした。")
        return None
    
    return file_path

def create_output_directory():
    """出力用ディレクトリを作成（スクリプトと同じディレクトリ）"""
    # スクリプトが配置されているディレクトリを取得
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 最上位フォルダ
    top_level_dir = os.path.join(script_dir, "reordered_matrices")
    os.makedirs(top_level_dir, exist_ok=True)
    
    # タイムスタンプ付きサブフォルダ
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = os.path.join(top_level_dir, timestamp)
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"✓ 出力ディレクトリ作成: reordered_matrices/{timestamp}")
    
    return output_dir

def load_similarity_matrix(csv_path):
    """類似度マトリクスの読み込み"""
    try:
        df = pd.read_csv(csv_path, index_col=0)
        print(f"✓ マトリクス読み込み成功: {df.shape[0]} x {df.shape[1]}")
        return df
    except Exception as e:
        print(f"✗ エラー: {e}")
        return None

def reorder_by_optimal_leaf_ordering(df):
    """
    最適葉順序付け (Optimal Leaf Ordering) による並べ替え
    隣接する要素間の距離を最小化し、見やすいマトリクスを作成
    """
    # 類似度→距離に変換（1 - 類似度）
    distance_matrix = 1 - df.values
    
    # 対角成分を0に補正（数値誤差対策）
    np.fill_diagonal(distance_matrix, 0)
    
    # 対称行列を圧縮形式に変換
    condensed_dist = squareform(distance_matrix, checks=False)
    
    # Ward法でリンケージ作成
    linkage = hierarchy.linkage(condensed_dist, method='ward')
    
    # 最適葉順序付け
    optimal_linkage = hierarchy.optimal_leaf_ordering(linkage, condensed_dist)
    
    # デンドログラムの順序を取得
    dendro = hierarchy.dendrogram(optimal_linkage, no_plot=True)
    order = dendro['leaves']
    
    # マトリクスを並べ替え
    reordered_df = df.iloc[order, order]
    
    return reordered_df, order

def reorder_by_hierarchical_clustering(df, method='average'):
    """
    階層的クラスタリングによる並べ替え
    
    Parameters:
    -----------
    df : DataFrame
        類似度マトリクス
    method : str
        'average' (平均連結法), 'ward' (Ward法), 'complete' (完全連結法)
    """
    distance_matrix = 1 - df.values
    np.fill_diagonal(distance_matrix, 0)
    
    condensed_dist = squareform(distance_matrix, checks=False)
    linkage = hierarchy.linkage(condensed_dist, method=method)
    
    dendro = hierarchy.dendrogram(linkage, no_plot=True)
    order = dendro['leaves']
    
    reordered_df = df.iloc[order, order]
    
    return reordered_df, order

def save_reordered_matrix(reordered_df, output_dir, base_name, method_name):
    """並べ替え後のマトリクスをCSV保存"""
    output_filename = f"{base_name}_reordered_{method_name}.csv"
    output_path = os.path.join(output_dir, output_filename)
    
    # CSV保存
    reordered_df.to_csv(output_path, encoding='utf-8-sig')
    print(f"✓ 保存完了: {output_filename}")
    
    return output_path

def save_order_info(order, original_labels, output_dir, base_name, method_name):
    """並べ替え順序情報をテキストファイルで保存"""
    output_filename = f"{base_name}_order_{method_name}.txt"
    output_path = os.path.join(output_dir, output_filename)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write(f"並べ替え方法: {method_name}\n")
        f.write("="*60 + "\n\n")
        f.write("元の順序 → 並べ替え後の順序:\n")
        f.write("-" * 60 + "\n")
        
        for new_idx, orig_idx in enumerate(order, 1):
            f.write(f"{new_idx:3d}. {original_labels[orig_idx]}\n")
    
    print(f"✓ 順序情報保存: {output_filename}")
    return output_path

def save_summary(output_dir, base_name, df_shape, original_file):
    """処理結果のサマリーを保存"""
    summary_path = os.path.join(output_dir, "README.txt")
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write("類似度マトリクス並べ替え結果\n")
        f.write("="*60 + "\n\n")
        f.write(f"処理日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n")
        f.write(f"元ファイル: {original_file}\n")
        f.write(f"マトリクスサイズ: {df_shape[0]} x {df_shape[1]}\n\n")
        
        f.write("-"*60 + "\n")
        f.write("出力ファイル:\n")
        f.write("-"*60 + "\n\n")
        
        f.write("1. 最適葉順序付け（推奨）\n")
        f.write(f"   - {base_name}_reordered_optimal.csv\n")
        f.write(f"   - {base_name}_order_optimal.txt\n")
        f.write("   説明: 隣接する要素間の距離を最小化。最も見やすい配置。\n\n")
        
        f.write("2. 平均連結法\n")
        f.write(f"   - {base_name}_reordered_average.csv\n")
        f.write(f"   - {base_name}_order_average.txt\n")
        f.write("   説明: クラスタ間の平均距離に基づく並べ替え。\n\n")
        
        f.write("3. Ward法\n")
        f.write(f"   - {base_name}_reordered_ward.csv\n")
        f.write(f"   - {base_name}_order_ward.txt\n")
        f.write("   説明: クラスタ内の分散を最小化。明確なグループ分け。\n\n")
        
        f.write("-"*60 + "\n")
        f.write("推奨:\n")
        f.write("-"*60 + "\n")
        f.write("最適葉順序付け（optimal）の結果が最も視覚的に見やすい配置です。\n")
        f.write("ヒートマップやExcelで開いて確認してください。\n")
    
    print(f"✓ サマリー保存: README.txt")
    return summary_path

def print_reorder_info(original_df, reordered_df, order):
    """並べ替え情報をコンソールに表示"""
    print("\n並べ替え結果（最初の10件）:")
    print("-" * 60)
    
    original_labels = original_df.index.tolist()
    
    for new_idx, orig_idx in enumerate(order[:10], 1):
        print(f"{new_idx:3d}. {original_labels[orig_idx]}")
    
    if len(order) > 10:
        print(f"  ... 他{len(order) - 10}件")

def main():
    """メイン処理"""
    print("="*60)
    print("類似度マトリクス並べ替えツール")
    print("="*60)
    
    # ファイル選択
    print("\nCSVファイルを選択してください...")
    csv_path = select_csv_file()
    
    if not csv_path:
        return
    
    original_filename = os.path.basename(csv_path)
    print(f"\n選択されたファイル: {original_filename}")
    
    # マトリクス読み込み
    df = load_similarity_matrix(csv_path)
    if df is None:
        return
    
    # 出力ディレクトリ作成（スクリプトと同じディレクトリ）
    output_dir = create_output_directory()
    base_name = os.path.splitext(original_filename)[0]
    
    original_labels = df.index.tolist()
    
    # 方法1: 最適葉順序付け（推奨）
    print("\n" + "-"*60)
    print("方法1: 最適葉順序付け（Optimal Leaf Ordering）")
    print("-"*60)
    reordered_df_opt, order_opt = reorder_by_optimal_leaf_ordering(df)
    save_reordered_matrix(reordered_df_opt, output_dir, base_name, 'optimal')
    save_order_info(order_opt, original_labels, output_dir, base_name, 'optimal')
    print_reorder_info(df, reordered_df_opt, order_opt)
    
    # 方法2: 平均連結法
    print("\n" + "-"*60)
    print("方法2: 階層的クラスタリング（平均連結法）")
    print("-"*60)
    reordered_df_avg, order_avg = reorder_by_hierarchical_clustering(df, method='average')
    save_reordered_matrix(reordered_df_avg, output_dir, base_name, 'average')
    save_order_info(order_avg, original_labels, output_dir, base_name, 'average')
    
    # 方法3: Ward法
    print("\n" + "-"*60)
    print("方法3: 階層的クラスタリング（Ward法）")
    print("-"*60)
    reordered_df_ward, order_ward = reorder_by_hierarchical_clustering(df, method='ward')
    save_reordered_matrix(reordered_df_ward, output_dir, base_name, 'ward')
    save_order_info(order_ward, original_labels, output_dir, base_name, 'ward')
    
    # サマリー作成
    print("\n" + "-"*60)
    save_summary(output_dir, base_name, df.shape, original_filename)
    
    print("\n" + "="*60)
    print("完了！")
    print("="*60)
    print(f"\n出力ディレクトリ: {output_dir}")
    print("\n生成されたファイル:")
    print("  ├─ README.txt (処理結果のサマリー)")
    print("  ├─ *_reordered_optimal.csv (推奨)")
    print("  ├─ *_order_optimal.txt")
    print("  ├─ *_reordered_average.csv")
    print("  ├─ *_order_average.txt")
    print("  ├─ *_reordered_ward.csv")
    print("  └─ *_order_ward.txt")
    print("\n推奨: 最適葉順序付け（optimal）の結果が最も見やすい配置になります。")
    print("="*60)

if __name__ == "__main__":
    main()