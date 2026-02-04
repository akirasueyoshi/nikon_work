import pandas as pd
import os
import time
from typing import Dict, List
from azure.identity import InteractiveBrowserCredential, get_bearer_token_provider
from openai import AzureOpenAI

AZURE_OPENAI_ENDPOINT = 'https://aoai-je-exm.openai.azure.com/'
DEPLOYMENT_NAME = 'gpt-4o'

# 段階1: 各シートの概要を作成
SHEET_SUMMARY_PROMPT = """
以下は半導体露光装置における{basename}という機能仕様書の「{sheet_name}」シートの内容です。
このシートの内容を簡潔に要約してください（200-300文字程度）。

要約には以下を含めてください：
- このシートの目的・役割
- 主要な情報の種類（パラメータ、フロー、状態遷移など）
- 重要なポイント
"""

# 段階2: 全体の解説を作成
FINAL_SUMMARY_PROMPT = """
以下は半導体露光装置における{basename}という機能仕様書の全シート概要です。
これらの情報を統合して、この機能の包括的な解説を作成してください。

## 解説に含めるべき内容
1. **機能概要**: この機能の目的と役割
2. **主要な処理フロー**: どのような処理が行われるか
3. **入出力**: 何を受け取り、何を出力するか
4. **重要なパラメータ**: キーとなる設定値や制約
5. **関連する機能**: 他の機能との関係性
6. **特記事項**: 注意すべき点や制約条件

回答はMarkdown形式で構造化してください。
"""

INPUT_DIR = r'resource'
OUTPUT_DIR = r'summary'

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


def split_dataframe_into_chunks(df: pd.DataFrame, max_rows: int = MAX_ROWS_PER_CHUNK) -> List[tuple]:
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


def summarize_sheet(basename: str, sheet_name: str, df: pd.DataFrame, client: AzureOpenAI,
                   use_chunking: bool = False) -> str:
    """
    1つのシートの概要を作成（必要に応じてチャンク分割）
    
    Args:
        basename: ファイルのベース名
        sheet_name: シート名
        df: シートのDataFrame
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
    
    Returns:
        シートの概要（テキスト）
    """
    # チャンク分割が不要な場合
    if not use_chunking or len(df) <= MAX_ROWS_PER_CHUNK:
        table_text = df.to_markdown()
        
        response = client.chat.completions.create(
            model=DEPLOYMENT_NAME,
            messages=[
                {"role": "system", "content": SHEET_SUMMARY_PROMPT.format(
                    basename=basename, 
                    sheet_name=sheet_name
                )},
                {"role": "user", "content": table_text}
            ],
            max_tokens=1000
        )
        return response.choices[0].message.content
    
    # チャンク分割して処理
    chunks = split_dataframe_into_chunks(df)
    chunk_summaries = []
    
    print(f'    Splitting into {len(chunks)} chunks...')
    
    for i, (start, end, chunk_df) in enumerate(chunks, 1):
        chunk_info = f"行{start+1}-{end}"
        print(f'    Processing chunk {i}/{len(chunks)}: {chunk_info}')
        
        try:
            table_text = chunk_df.to_markdown()
            
            response = client.chat.completions.create(
                model=DEPLOYMENT_NAME,
                messages=[
                    {"role": "system", "content": SHEET_SUMMARY_PROMPT.format(
                        basename=basename, 
                        sheet_name=f"{sheet_name} ({chunk_info})"
                    )},
                    {"role": "user", "content": table_text}
                ],
                max_tokens=1000
            )
            chunk_summaries.append(response.choices[0].message.content)
            
            if i < len(chunks):
                time.sleep(WAIT_TIME_BETWEEN_SHEETS)
                
        except Exception as e:
            print(f'    ⚠ Error processing chunk {chunk_info}: {str(e)}')
            chunk_summaries.append(f"[{chunk_info}] エラー: {str(e)}")
    
    # チャンク概要を統合
    return "\n\n".join(chunk_summaries)


def create_sheet_summaries(basename: str, all_sheets: Dict[str, pd.DataFrame], 
                          client: AzureOpenAI, use_chunking: bool = False) -> Dict[str, str]:
    """
    全シートの概要を作成
    
    Args:
        basename: ファイルのベース名
        all_sheets: {シート名: DataFrame}の辞書
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
    
    Returns:
        {シート名: 概要}の辞書
    """
    sheet_summaries = {}
    
    for i, (sheet_name, df) in enumerate(all_sheets.items(), 1):
        print(f'  Summarizing sheet {i}/{len(all_sheets)}: {sheet_name} ({len(df)} rows)')
        
        try:
            summary = summarize_sheet(basename, sheet_name, df, client, use_chunking)
            sheet_summaries[sheet_name] = summary
            
            # レート制限対策
            if i < len(all_sheets):
                time.sleep(WAIT_TIME_BETWEEN_SHEETS)
                
        except Exception as e:
            print(f'  ⚠ Error summarizing sheet {sheet_name}: {str(e)}')
            sheet_summaries[sheet_name] = f"エラー: {str(e)}"
    
    return sheet_summaries


def create_final_summary(basename: str, sheet_summaries: Dict[str, str], 
                        client: AzureOpenAI) -> str:
    """
    全シート概要から最終的な解説を作成
    
    Args:
        basename: ファイルのベース名
        sheet_summaries: {シート名: 概要}の辞書
        client: AzureOpenAIクライアント
    
    Returns:
        最終的な解説（Markdown形式）
    """
    # 全シート概要をテキストに整形
    summaries_text = []
    for sheet_name, summary in sheet_summaries.items():
        summaries_text.append(f"### {sheet_name}\n{summary}\n")
    
    combined_summaries = "\n".join(summaries_text)
    
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {"role": "system", "content": FINAL_SUMMARY_PROMPT.format(basename=basename)},
            {"role": "user", "content": combined_summaries}
        ],
        max_tokens=10000
    )
    return response.choices[0].message.content


def create_summary(rel_path: str, resource_dir: str, client: AzureOpenAI, 
                  use_chunking: bool = False) -> str:
    """
    Excelファイルから解説を作成（2段階方式）
    
    段階1: 各シートの概要を作成
    段階2: 全シート概要を統合して最終的な解説を作成
    
    Args:
        rel_path: resourceディレクトリからの相対パス（拡張子なし）
        resource_dir: resourceディレクトリのパス
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
    
    Returns:
        最終的な解説（Markdown形式）
    """
    basename = os.path.basename(rel_path)
    
    # Excelファイルのフルパスを構築
    excel_path = os.path.join(resource_dir, f'{rel_path}.xlsx')
    if not os.path.exists(excel_path):
        excel_path = os.path.join(resource_dir, f'{rel_path}.xls')
    
    # 全シートを読み込む
    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    print(f'  Found {len(all_sheets)} sheets')
    
    # 段階1: 各シートの概要を作成
    print(f'  Phase 1: Creating sheet summaries...')
    sheet_summaries = create_sheet_summaries(basename, all_sheets, client, use_chunking)
    
    # 段階2: 全体の解説を作成
    print(f'  Phase 2: Creating final summary...')
    time.sleep(WAIT_TIME_BETWEEN_SHEETS)
    final_summary = create_final_summary(basename, sheet_summaries, client)
    
    # 最終的なMarkdownを構築
    result_parts = [
        f"# {basename} - 機能解説\n\n",
        final_summary,
        "\n\n---\n\n",
        "## 各シート概要\n\n"
    ]
    
    for sheet_name, summary in sheet_summaries.items():
        result_parts.append(f"### {sheet_name}\n\n{summary}\n\n")
    
    return "".join(result_parts)


def main(use_chunking: bool = False):
    """
    メイン処理
    
    Args:
        use_chunking: 大きなシートをチャンク分割するか
    """
    inputs = get_excel_files(INPUT_DIR)
    
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
    
    # 未処理のファイルを抽出
    undescribeds = [f for f in inputs if f not in output_mds]
    
    print(f'Found {len(inputs)} Excel files')
    print(f'Already processed: {len(output_mds)}')
    print(f'Processing {len(undescribeds)} files...')
    print(f'Chunking mode: {"ON" if use_chunking else "OFF"}')
    
    if len(undescribeds) == 0:
        print('No files to process.')
        return
    
    client = get_client()
    
    for idx, undescribed in enumerate(undescribeds, 1):
        try:
            print(f'\n[{idx}/{len(undescribeds)}] 📄 Processing: {undescribed}')
            description = create_summary(undescribed, INPUT_DIR, client, use_chunking)
            
            # 出力ファイルのパスを構築
            output_path = os.path.join(OUTPUT_DIR, f'{undescribed}.md')
            
            # 出力ディレクトリを作成（存在しない場合）
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # マークダウンファイルを保存
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(description)
            
            print(f'✓ Summary file created: {undescribed}')
            
            # ファイル間の待機
            if idx < len(undescribeds):
                time.sleep(WAIT_TIME_BETWEEN_FILES)
            
        except Exception as e:
            print(f'✗ Error processing {undescribed}: {str(e)}')
    
    print(f'\n{"="*60}')
    print(f'Processing complete!')


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Create summary from Excel files')
    parser.add_argument('--chunking', action='store_true', 
                       help='Enable chunking for large sheets')
    
    args = parser.parse_args()
    
    main(use_chunking=args.chunking)
