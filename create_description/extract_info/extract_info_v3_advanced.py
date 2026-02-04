import pandas as pd
import os
import time
import json
from typing import List, Dict, Tuple
from azure.identity import InteractiveBrowserCredential, get_bearer_token_provider
from openai import AzureOpenAI

AZURE_OPENAI_ENDPOINT = 'https://aoai-je-exm.openai.azure.com/'
DEPLOYMENT_NAME = 'gpt-4o'

PROMPT = """
以降のユーザープロンプトで示されるテキストはExcelファイルをDataframe化したものです。
このExcelファイルは半導体露光装置における{basename}という機能仕様を表現した仕様書です。
この仕様書の記述から「明示的に記載されている」事実のみを以下の指示に従って抽出してください。
回答はMarkdown形式の記述としてください。

## 抽出する事実の種類
- 処理内容（何をするか）
- 入出力（何を受け取り何を返すか）
- 状態遷移
- 制約・前提条件
- 依存・関連する機能

## 抽出ルール
- 推測・一般知識・言い換えは禁止
- 原文の語彙をできるだけ維持すること
- 書かれていないことは「記載なし」とし、補完しない
- 1つの事実 = 1つの独立した主張(1〜3文程度)とする
- 出典（セクション名/見出し）を付与すること

## 出力フォーマット
- [処理内容] 
- [入出力] 
- [状態遷移] 
- [制約・前提条件] 
- [依存・関連する機能] 

"""

INPUT_DIR = r'resource'
OUTPUT_DIR = r'information'
PROGRESS_FILE = r'progress.json'

# レート制限対策の待機時間（秒）
WAIT_TIME_BETWEEN_SHEETS = 2
WAIT_TIME_BETWEEN_FILES = 5

# シートを分割する閾値（行数）
MAX_ROWS_PER_CHUNK = 100


def get_excel_files(dir_path: str) -> list[str]:
    """
    指定ディレクトリ配下のすべてのExcelファイルを再帰的に取得
    ルートディレクトリからの相対パス（拡張子なし）のリストを返す
    """
    excel_files = []
    
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
        return excel_files
    
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, dir_path)
                rel_path_without_ext = os.path.splitext(rel_path)[0]
                excel_files.append(rel_path_without_ext)
    
    return excel_files


def get_client() -> AzureOpenAI:
    credential = InteractiveBrowserCredential(
        tenant_id="4876a51c-4f2d-4d54-b712-e0b67d308e80"
    )
    
    token_provider = get_bearer_token_provider(
        credential,
        "https://cognitiveservices.azure.com/.default"
    )

    return AzureOpenAI(
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        azure_ad_token_provider=token_provider,
        api_version="2024-12-01-preview",
    )


def split_dataframe_into_chunks(df: pd.DataFrame, max_rows: int = MAX_ROWS_PER_CHUNK) -> List[Tuple[int, int, pd.DataFrame]]:
    """
    DataFrameを指定行数で分割
    
    Args:
        df: 分割するDataFrame
        max_rows: チャンクあたりの最大行数
    
    Returns:
        (開始行, 終了行, チャンク)のリスト
    """
    chunks = []
    total_rows = len(df)
    
    for start in range(0, total_rows, max_rows):
        end = min(start + max_rows, total_rows)
        chunk = df.iloc[start:end]
        chunks.append((start, end, chunk))
    
    return chunks


def extract_info_from_chunk(basename: str, sheet_name: str, chunk_info: str, 
                           chunk_df: pd.DataFrame, client: AzureOpenAI) -> str:
    """
    1つのチャンクから情報を抽出
    
    Args:
        basename: ファイルのベース名
        sheet_name: シート名
        chunk_info: チャンク情報（例: "行1-100"）
        chunk_df: チャンクのDataFrame
        client: AzureOpenAIクライアント
    
    Returns:
        抽出された情報（Markdown形式）
    """
    table_text = f"## シート: {sheet_name} ({chunk_info})\n\n{chunk_df.to_markdown()}"
    
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {"role": "system", "content": PROMPT.format(basename=basename)},
            {"role": "user", "content": table_text}
        ],
        max_tokens=10000
    )
    return response.choices[0].message.content


def extract_info_from_sheet(basename: str, sheet_name: str, df: pd.DataFrame, 
                           client: AzureOpenAI, use_chunking: bool = False) -> str:
    """
    1つのシートから情報を抽出（必要に応じてチャンク分割）
    
    Args:
        basename: ファイルのベース名
        sheet_name: シート名
        df: シートのDataFrame
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
    
    Returns:
        抽出された情報（Markdown形式）
    """
    if not use_chunking or len(df) <= MAX_ROWS_PER_CHUNK:
        # チャンク分割不要
        table_text = f"## シート: {sheet_name}\n\n{df.to_markdown()}"
        
        response = client.chat.completions.create(
            model=DEPLOYMENT_NAME,
            messages=[
                {"role": "system", "content": PROMPT.format(basename=basename)},
                {"role": "user", "content": table_text}
            ],
            max_tokens=10000
        )
        return response.choices[0].message.content
    else:
        # チャンク分割して処理
        chunks = split_dataframe_into_chunks(df)
        chunk_results = []
        
        print(f'    Splitting into {len(chunks)} chunks...')
        
        for i, (start, end, chunk_df) in enumerate(chunks, 1):
            chunk_info = f"行{start+1}-{end}"
            print(f'    Processing chunk {i}/{len(chunks)}: {chunk_info}')
            
            try:
                chunk_result = extract_info_from_chunk(
                    basename, sheet_name, chunk_info, chunk_df, client
                )
                chunk_results.append(f"### {chunk_info}\n\n{chunk_result}")
                
                if i < len(chunks):
                    time.sleep(WAIT_TIME_BETWEEN_SHEETS)
                    
            except Exception as e:
                print(f'    ⚠ Error processing chunk {chunk_info}: {str(e)}')
                chunk_results.append(f"### {chunk_info}\n\n**エラー**: {str(e)}")
        
        return "\n\n".join(chunk_results)


def extract_info(rel_path: str, resource_dir: str, client: AzureOpenAI, 
                use_chunking: bool = False) -> str:
    """
    Excelファイルから情報を抽出（シート単位で処理）
    
    Args:
        rel_path: resourceディレクトリからの相対パス（拡張子なし）
        resource_dir: resourceディレクトリのパス
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
    
    Returns:
        全シートから抽出された情報を統合したMarkdown文字列
    """
    basename = os.path.basename(rel_path)
    
    excel_path = os.path.join(resource_dir, f'{rel_path}.xlsx')
    if not os.path.exists(excel_path):
        excel_path = os.path.join(resource_dir, f'{rel_path}.xls')
    
    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    
    extracted_parts = [f"# {basename} - 抽出情報\n\n"]
    
    for i, (sheet_name, df) in enumerate(all_sheets.items(), 1):
        print(f'  Processing sheet {i}/{len(all_sheets)}: {sheet_name} ({len(df)} rows)')
        
        try:
            sheet_info = extract_info_from_sheet(basename, sheet_name, df, client, use_chunking)
            
            extracted_parts.append(f"## シート: {sheet_name}\n\n")
            extracted_parts.append(sheet_info)
            extracted_parts.append("\n\n---\n\n")
            
            if i < len(all_sheets):
                time.sleep(WAIT_TIME_BETWEEN_SHEETS)
            
        except Exception as e:
            print(f'  ⚠ Error processing sheet {sheet_name}: {str(e)}')
            extracted_parts.append(f"## シート: {sheet_name}\n\n")
            extracted_parts.append(f"**エラー**: {str(e)}\n\n---\n\n")
    
    return "".join(extracted_parts)


def load_progress() -> Dict:
    """処理進捗を読み込み"""
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"completed": [], "failed": []}


def save_progress(progress: Dict):
    """処理進捗を保存"""
    with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)


def main(use_chunking: bool = False, resume: bool = True):
    """
    メイン処理
    
    Args:
        use_chunking: 大きなシートをチャンク分割するか
        resume: 中断したところから再開するか
    """
    inputs = get_excel_files(INPUT_DIR)
    
    # 進捗情報を読み込み
    progress = load_progress() if resume else {"completed": [], "failed": []}
    
    # 出力済みのマークダウンファイルを取得
    output_mds = []
    if os.path.exists(OUTPUT_DIR):
        for root, dirs, files in os.walk(OUTPUT_DIR):
            for file in files:
                if file.endswith('.md'):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, OUTPUT_DIR)
                    rel_path_without_ext = os.path.splitext(rel_path)[0]
                    output_mds.append(rel_path_without_ext)
    
    # 未処理のファイルを抽出（完了済みと失敗済みを除外）
    completed_set = set(progress.get("completed", []))
    undescribeds = [f for f in inputs if f not in output_mds and f not in completed_set]
    
    print(f'Found {len(inputs)} Excel files')
    print(f'Already completed: {len(progress.get("completed", []))}')
    print(f'Previously failed: {len(progress.get("failed", []))}')
    print(f'Processing {len(undescribeds)} files...')
    print(f'Chunking mode: {"ON" if use_chunking else "OFF"}')
    
    client = get_client()
    
    for idx, undescribed in enumerate(undescribeds, 1):
        try:
            print(f'\n[{idx}/{len(undescribeds)}] 📄 Processing: {undescribed}')
            description = extract_info(undescribed, INPUT_DIR, client, use_chunking)
            
            output_path = os.path.join(OUTPUT_DIR, f'{undescribed}.md')
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(description)
            
            # 進捗を更新
            progress["completed"].append(undescribed)
            if undescribed in progress.get("failed", []):
                progress["failed"].remove(undescribed)
            save_progress(progress)
            
            print(f'✓ Information file created: {undescribed}')
            
            # ファイル間の待機
            if idx < len(undescribeds):
                time.sleep(WAIT_TIME_BETWEEN_FILES)
            
        except Exception as e:
            print(f'✗ Error processing {undescribed}: {str(e)}')
            if undescribed not in progress.get("failed", []):
                progress.setdefault("failed", []).append(undescribed)
            save_progress(progress)
    
    print(f'\n{"="*60}')
    print(f'Processing complete!')
    print(f'Total completed: {len(progress.get("completed", []))}')
    print(f'Total failed: {len(progress.get("failed", []))}')


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Extract information from Excel files')
    parser.add_argument('--chunking', action='store_true', 
                       help='Enable chunking for large sheets')
    parser.add_argument('--no-resume', action='store_true',
                       help='Start from scratch (ignore progress file)')
    
    args = parser.parse_args()
    
    main(use_chunking=args.chunking, resume=not args.no_resume)
