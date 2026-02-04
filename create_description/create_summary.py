import pandas as pd
import os
from azure.identity import InteractiveBrowserCredential, get_bearer_token_provider
from openai import AzureOpenAI

AZURE_OPENAI_ENDPOINT = 'https://aoai-je-exm.openai.azure.com/'
DEPLOYMENT_NAME = 'gpt-4o'

PROMPT = """
以降のユーザープロンプトで示されるテキストはExcelファイルをDataframe化したものです。
このExcelファイルは半導体露光装置における{basename}という機能仕様を表現した仕様書です。
この仕様書の記述から解説を作ってください。

"""

INPUT_DIR = r'resource'
OUTPUT_DIR = r'summary'


def get_excel_files(dir_path: str) -> list[str]:
    """
    指定ディレクトリ配下のすべてのExcelファイルを再帰的に取得
    ルートディレクトリからの相対パス（拡張子なし）のリストを返す
    """
    excel_files = []
    
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
        return excel_files
    
    # os.walkで再帰的にディレクトリを探索
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            # Excelファイルのみを対象
            if file.endswith(('.xlsx', '.xls')):
                # 絶対パスを取得
                full_path = os.path.join(root, file)
                # ルートディレクトリからの相対パスを取得
                rel_path = os.path.relpath(full_path, dir_path)
                # 拡張子を除いた相対パスを保存
                rel_path_without_ext = os.path.splitext(rel_path)[0]
                excel_files.append(rel_path_without_ext)
    
    return excel_files


def get_client() -> AzureOpenAI:
    # ブラウザで認証（初回のみブラウザが開く）
    credential = InteractiveBrowserCredential(
        tenant_id="4876a51c-4f2d-4d54-b712-e0b67d308e80"  # 必要に応じて指定
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


def create_summary(rel_path: str, resource_dir: str, client: AzureOpenAI) -> str:
    
    # ファイル名（ベース名）を取得
    basename = os.path.basename(rel_path)
    
    # Excelファイルのフルパスを構築
    excel_path = os.path.join(resource_dir, f'{rel_path}.xlsx')
    
    # .xlsxが存在しない場合は.xlsを試す
    if not os.path.exists(excel_path):
        excel_path = os.path.join(resource_dir, f'{rel_path}.xls')
    
    df = pd.read_excel(excel_path)
    table_text = df.to_markdown()
    
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {"role": "system", "content": PROMPT.format(basename=basename)},
            {"role": "user", "content": table_text}
        ],
        max_tokens=10000
    )
    return response.choices[0].message.content


def main():
    # 入力ディレクトリと出力ディレクトリからファイルリストを取得
    inputs = get_excel_files(INPUT_DIR)
    outputs = get_excel_files(OUTPUT_DIR)  # .mdファイルも取得したい場合は別関数が必要
    
    # 出力済みのマークダウンファイルを取得（拡張子なし）
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
    undescribeds = list(set(inputs).difference(output_mds))
    
    print(f'Found {len(inputs)} Excel files, {len(output_mds)} already processed.')
    print(f'Processing {len(undescribeds)} files...')
    
    client = get_client()
    
    for undescribed in undescribeds:
        try:
            description = create_summary(undescribed, INPUT_DIR, client)
            
            # 出力ファイルのパスを構築
            output_path = os.path.join(OUTPUT_DIR, f'{undescribed}.md')
            
            # 出力ディレクトリを作成（存在しない場合）
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # マークダウンファイルを保存
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(description)
            
            print(f'✓ Information file created: {undescribed}')
            
        except Exception as e:
            print(f'✗ Error processing {undescribed}: {str(e)}')


if __name__ == "__main__":
    main()
