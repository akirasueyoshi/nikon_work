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
        title='類似度/関連度マトリクスCSVファイルを選択',
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
    script_dir = os.path.dirname(os.path.abspath(__file__))
    top_level_dir = os.path.join(script_dir, "reordered_matrices")
    os.makedirs(top_level_dir, exist_ok=True)
    
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

def detect_matrix_type(df):
    """
    マトリクスのタイプを検出
    - similarity: 類似度マトリクス（対角=1, 値の範囲0-1）
    - cooccurrence: 同時出現マトリクス（対角=0, 整数値）
    """
    diagonal = np.diag(df.values)
    values = df.values[np.triu_indices_from(df.values, k=1)]
    
    # 対角成分がすべて1で、値が0-1の範囲 → 類似度マトリクス
    if np.allclose(diagonal, 1.0) and values.min() >= 0 and values.max() <= 1:
        return "similarity"
    
    # 対角成分がすべて0で、整数値 → 同時出現マトリクス
    if np.allclose(diagonal, 0.0) and np.all(values == values.astype(int)):
        return "cooccurrence"
    
    # その他の場合は類似度と仮定
    return "similarity"

def convert_to_distance_matrix(df, matrix_type):
    """
    マトリクスを距離行列に変換
    
    Parameters:
    -----------
    df : DataFrame
        入力マトリクス
    matrix_type : str
        'similarity' or 'cooccurrence'
    
    Returns:
    --------
    distance_matrix : ndarray
        距離行列
    """
    if matrix_type == "similarity":
        # 類似度 → 距離: 1 - 類似度
        distance_matrix = 1 - df.values
    else:
        # 同時出現 → 距離: 最大値 - 同時出現回数
        max_value = df.values.max()
        if max_value == 0:
            # すべて0の場合は一様な距離
            distance_matrix = np.ones_like(df.values)
        else:
            distance_matrix = max_value - df.values
    
    # 対角成分を0に設定
    np.fill_diagonal(distance_matrix, 0)
    
    # 負の値をチェック（数値誤差で負になる場合がある）
    distance_matrix = np.maximum(distance_matrix, 0)
    
    return distance_matrix

def reorder_by_optimal_leaf_ordering(df, matrix_type):
    """
    最適葉順序付け (Optimal Leaf Ordering) による並べ替え
    """
    # 距離行列に変換
    distance_matrix = convert_to_distance_matrix(df, matrix_type)
    
    # 対称性を強制（数値誤差対策）
    distance_matrix = (distance_matrix + distance_matrix.T) / 2
    
    # 対称行列を圧縮形式に変換
    condensed_dist = squareform(distance_matrix, checks=False)
    
    # 距離が全て0の場合の対策
    if np.all(condensed_dist == 0):
        print("  警告: すべての距離が0です。元の順序を維持します。")
        return df, list(range(len(df)))
    
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

def reorder_by_hierarchical_clustering(df, matrix_type, method='average'):
    """
    階層的クラスタリングによる並べ替え
    """
    distance_matrix = convert_to_distance_matrix(df, matrix_type)
    distance_matrix = (distance_matrix + distance_matrix.T) / 2
    
    condensed_dist = squareform(distance_matrix, checks=False)
    
    if np.all(condensed_dist == 0):
        print("  警告: すべての距離が0です。元の順序を維持します。")
        return df, list(range(len(df)))
    
    linkage = hierarchy.linkage(condensed_dist, method=method)
    
    dendro = hierarchy.dendrogram(linkage, no_plot=True)
    order = dendro['leaves']
    
    reordered_df = df.iloc[order, order]
    
    return reordered_df, order

def save_reordered_matrix(reordered_df, output_dir, base_name, method_name):
    """並べ替え後のマトリクスをCSV保存"""
    output_filename = f"{base_name}_reordered_{method_name}.csv"
    output_path = os.path.join(output_dir, output_filename)
    
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

def save_summary(output_dir, base_name, df_shape, original_file, matrix_type):
    """処理結果のサマリーを保存"""
    summary_path = os.path.join(output_dir, "README.txt")
    
    matrix_type_str = "類似度マトリクス" if matrix_type == "similarity" else "同時出現マトリクス"
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("="*60 + "\n")
        f.write("マトリクス並べ替え結果\n")
        f.write("="*60 + "\n\n")
        f.write(f"処理日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n")
        f.write(f"元ファイル: {original_file}\n")
        f.write(f"マトリクスタイプ: {matrix_type_str}\n")
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
        f.write("距離変換方法:\n")
        f.write("-"*60 + "\n")
        if matrix_type == "similarity":
            f.write("類似度マトリクス: 距離 = 1 - 類似度\n")
        else:
            f.write("同時出現マトリクス: 距離 = 最大値 - 同時出現回数\n")
            f.write("（同時出現回数が多いほど距離が近い）\n")
        
        f.write("\n推奨:\n")
        f.write("最適葉順序付け（optimal）の結果が最も視覚的に見やすい配置です。\n")
    
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
    print("マトリクス並べ替えツール")
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
    
    # マトリクスタイプを検出
    matrix_type = detect_matrix_type(df)
    matrix_type_str = "類似度マトリクス" if matrix_type == "similarity" else "同時出現マトリクス"
    print(f"✓ マトリクスタイプ: {matrix_type_str}")
    
    # 出力ディレクトリ作成（スクリプトと同じディレクトリ）
    output_dir = create_output_directory()
    base_name = os.path.splitext(original_filename)[0]
    
    original_labels = df.index.tolist()
    
    # 方法1: 最適葉順序付け（推奨）
    print("\n" + "-"*60)
    print("方法1: 最適葉順序付け（Optimal Leaf Ordering）")
    print("-"*60)
    try:
        reordered_df_opt, order_opt = reorder_by_optimal_leaf_ordering(df, matrix_type)
        save_reordered_matrix(reordered_df_opt, output_dir, base_name, 'optimal')
        save_order_info(order_opt, original_labels, output_dir, base_name, 'optimal')
        print_reorder_info(df, reordered_df_opt, order_opt)
    except Exception as e:
        print(f"✗ エラー: {e}")
    
    # 方法2: 平均連結法
    print("\n" + "-"*60)
    print("方法2: 階層的クラスタリング（平均連結法）")
    print("-"*60)
    try:
        reordered_df_avg, order_avg = reorder_by_hierarchical_clustering(df, matrix_type, method='average')
        save_reordered_matrix(reordered_df_avg, output_dir, base_name, 'average')
        save_order_info(order_avg, original_labels, output_dir, base_name, 'average')
    except Exception as e:
        print(f"✗ エラー: {e}")
    
    # 方法3: Ward法
    print("\n" + "-"*60)
    print("方法3: 階層的クラスタリング（Ward法）")
    print("-"*60)
    try:
        reordered_df_ward, order_ward = reorder_by_hierarchical_clustering(df, matrix_type, method='ward')
        save_reordered_matrix(reordered_df_ward, output_dir, base_name, 'ward')
        save_order_info(order_ward, original_labels, output_dir, base_name, 'ward')
    except Exception as e:
        print(f"✗ エラー: {e}")
    
    # サマリー作成
    print("\n" + "-"*60)
    save_summary(output_dir, base_name, df.shape, original_filename, matrix_type)
    
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