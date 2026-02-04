import pandas as pd
import os
import time
from typing import Dict, List
from azure.identity import InteractiveBrowserCredential, get_bearer_token_provider
from openai import AzureOpenAI

AZURE_OPENAI_ENDPOINT = 'https://aoai-je-exm.openai.azure.com/'
DEPLOYMENT_NAME = 'gpt-4o'

# 各シートの詳細解説を作成
SHEET_DETAIL_PROMPT = """
以下は半導体露光装置における{basename}という機能仕様書の「{sheet_name}」シートの内容です。
このシートの内容について詳細な解説を作成してください。

## 解説に含めるべき内容
- シートの目的と役割
- 主要な情報の説明
- パラメータや設定値の意味
- 処理フローや状態遷移（該当する場合）
- 注意すべき点や制約

回答はMarkdown形式で、見出しを使って構造化してください。
"""

# 統合解説を作成
INTEGRATION_PROMPT = """
以下は半導体露光装置における{basename}という機能仕様書の各シート解説です。
これらの情報を統合して、機能全体の包括的な解説を冒頭に追加してください。

## 統合解説に含めるべき内容
1. **機能全体の概要**: この機能の目的と位置づけ
2. **主要な処理フロー**: 全体的な処理の流れ
3. **シート間の関係**: 各シートがどう関連しているか
4. **重要なポイント**: 理解しておくべき重要事項
5. **使用シーン**: どのような場面で使用されるか

統合解説を作成したら、その後に各シートの詳細解説を続けてください。
全体をMarkdown形式で出力してください。
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


def create_sheet_detail(basename: str, sheet_name: str, df: pd.DataFrame, 
                       client: AzureOpenAI, use_chunking: bool = False) -> str:
    """
    1つのシートの詳細解説を作成（必要に応じてチャンク分割）
    
    Args:
        basename: ファイルのベース名
        sheet_name: シート名
        df: シートのDataFrame
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
    
    Returns:
        シートの詳細解説（Markdown形式）
    """
    # チャンク分割が不要な場合
    if not use_chunking or len(df) <= MAX_ROWS_PER_CHUNK:
        table_text = df.to_markdown()
        
        response = client.chat.completions.create(
            model=DEPLOYMENT_NAME,
            messages=[
                {"role": "system", "content": SHEET_DETAIL_PROMPT.format(
                    basename=basename, 
                    sheet_name=sheet_name
                )},
                {"role": "user", "content": table_text}
            ],
            max_tokens=5000
        )
        return response.choices[0].message.content
    
    # チャンク分割して処理
    chunks = split_dataframe_into_chunks(df)
    chunk_details = []
    
    print(f'    Splitting into {len(chunks)} chunks...')
    
    for i, (start, end, chunk_df) in enumerate(chunks, 1):
        chunk_info = f"行{start+1}-{end}"
        print(f'    Processing chunk {i}/{len(chunks)}: {chunk_info}')
        
        try:
            table_text = chunk_df.to_markdown()
            
            response = client.chat.completions.create(
                model=DEPLOYMENT_NAME,
                messages=[
                    {"role": "system", "content": SHEET_DETAIL_PROMPT.format(
                        basename=basename, 
                        sheet_name=f"{sheet_name} ({chunk_info})"
                    )},
                    {"role": "user", "content": table_text}
                ],
                max_tokens=5000
            )
            chunk_details.append(f"### {chunk_info}\n\n{response.choices[0].message.content}")
            
            if i < len(chunks):
                time.sleep(WAIT_TIME_BETWEEN_SHEETS)
                
        except Exception as e:
            print(f'    ⚠ Error processing chunk {chunk_info}: {str(e)}')
            chunk_details.append(f"### {chunk_info}\n\n**エラー**: {str(e)}")
    
    # チャンク詳細を統合
    return "\n\n".join(chunk_details)


def create_sheet_details(basename: str, all_sheets: Dict[str, pd.DataFrame], 
                        client: AzureOpenAI, use_chunking: bool = False) -> Dict[str, str]:
    """
    全シートの詳細解説を作成
    
    Args:
        basename: ファイルのベース名
        all_sheets: {シート名: DataFrame}の辞書
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
    
    Returns:
        {シート名: 詳細解説}の辞書
    """
    sheet_details = {}
    
    for i, (sheet_name, df) in enumerate(all_sheets.items(), 1):
        print(f'  Creating detailed explanation for sheet {i}/{len(all_sheets)}: {sheet_name} ({len(df)} rows)')
        
        try:
            detail = create_sheet_detail(basename, sheet_name, df, client, use_chunking)
            sheet_details[sheet_name] = detail
            
            # レート制限対策
            if i < len(all_sheets):
                time.sleep(WAIT_TIME_BETWEEN_SHEETS)
                
        except Exception as e:
            print(f'  ⚠ Error processing sheet {sheet_name}: {str(e)}')
            sheet_details[sheet_name] = f"**エラー**: {str(e)}"
    
    return sheet_details


def integrate_explanations(basename: str, sheet_details: Dict[str, str], 
                          client: AzureOpenAI) -> str:
    """
    各シート解説を統合して全体解説を追加
    
    Args:
        basename: ファイルのベース名
        sheet_details: {シート名: 詳細解説}の辞書
        client: AzureOpenAIクライアント
    
    Returns:
        統合された解説（Markdown形式）
    """
    # 各シート解説をテキストに整形
    details_text = []
    for sheet_name, detail in sheet_details.items():
        details_text.append(f"## {sheet_name}\n\n{detail}\n\n---\n")
    
    combined_details = "\n".join(details_text)
    
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {"role": "system", "content": INTEGRATION_PROMPT.format(basename=basename)},
            {"role": "user", "content": combined_details}
        ],
        max_tokens=16000
    )
    return response.choices[0].message.content


def create_summary_detailed(rel_path: str, resource_dir: str, client: AzureOpenAI,
                           use_chunking: bool = False) -> str:
    """
    詳細解説方式でExcelファイルから解説を作成
    
    段階1: 各シートの詳細解説を作成
    段階2: 詳細解説を統合して全体解説を追加
    
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
    
    # 段階1: 各シートの詳細解説を作成
    print(f'  Phase 1: Creating detailed explanations for each sheet...')
    sheet_details = create_sheet_details(basename, all_sheets, client, use_chunking)
    
    # 段階2: 統合解説を作成
    print(f'  Phase 2: Integrating explanations...')
    time.sleep(WAIT_TIME_BETWEEN_SHEETS)
    integrated_summary = integrate_explanations(basename, sheet_details, client)
    
    return integrated_summary


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
            
            # 詳細モードで解説作成
            description = create_summary_detailed(undescribed, INPUT_DIR, client, use_chunking)
            
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
    
    parser = argparse.ArgumentParser(description='Create detailed summary from Excel files')
    parser.add_argument('--chunking', action='store_true', 
                       help='Enable chunking for large sheets')
    
    args = parser.parse_args()
    
    main(use_chunking=args.chunking)
