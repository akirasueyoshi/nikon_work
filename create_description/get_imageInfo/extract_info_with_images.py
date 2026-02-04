import pandas as pd
import os
import time
import openpyxl
from PIL import Image
import io
import base64
from typing import List, Tuple, Dict, Optional
from azure.identity import InteractiveBrowserCredential, get_bearer_token_provider
from openai import AzureOpenAI

AZURE_OPENAI_ENDPOINT = 'https://aoai-je-exm.openai.azure.com/'
DEPLOYMENT_NAME = 'gpt-4o'  # gpt-4oはvision対応

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

## 画像について
- 画像が含まれている場合は、画像の内容も分析して情報を抽出してください
- 図表、フローチャート、状態遷移図などは特に重要です
- 画像から読み取れる情報も上記の事実として抽出してください

## 出力フォーマット
- [処理内容] 
- [入出力] 
- [状態遷移] 
- [制約・前提条件] 
- [依存・関連する機能] 

"""

INPUT_DIR = r'resource'
OUTPUT_DIR = r'information'

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


def extract_images_from_sheet(excel_path: str, sheet_name: str) -> List[Image.Image]:
    """
    特定のシートから画像を抽出
    
    Args:
        excel_path: Excelファイルのパス
        sheet_name: シート名
    
    Returns:
        PIL Imageのリスト
    """
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb[sheet_name]
        
        images = []
        if hasattr(sheet, '_images'):
            for img in sheet._images:
                image_data = img._data()
                pil_image = Image.open(io.BytesIO(image_data))
                images.append(pil_image)
        
        return images
    except Exception as e:
        print(f'    Warning: Could not extract images from {sheet_name}: {str(e)}')
        return []


def image_to_base64(image: Image.Image) -> str:
    """
    PIL ImageをBase64文字列に変換
    
    Args:
        image: PIL Image
    
    Returns:
        Base64エンコードされた文字列
    """
    buffered = io.BytesIO()
    # 大きな画像はリサイズ（トークン節約）
    max_size = (1024, 1024)
    image.thumbnail(max_size, Image.Resampling.LANCZOS)
    image.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return img_str


def create_message_content(text: str, images: List[Image.Image]) -> list:
    """
    テキストと画像を含むメッセージコンテンツを作成
    
    Args:
        text: テキスト内容
        images: PIL Imageのリスト
    
    Returns:
        Azure OpenAI APIに渡せるコンテンツ形式
    """
    content = [
        {
            "type": "text",
            "text": text
        }
    ]
    
    # 画像を追加（最大5枚まで）
    for img in images[:5]:
        base64_image = image_to_base64(img)
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{base64_image}"
            }
        })
    
    return content


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


def extract_info_from_sheet(basename: str, sheet_name: str, df: pd.DataFrame, 
                           excel_path: str, client: AzureOpenAI, 
                           use_chunking: bool = False,
                           include_images: bool = True) -> str:
    """
    1つのシートから情報を抽出（画像対応版）
    
    Args:
        basename: ファイルのベース名
        sheet_name: シート名
        df: シートのDataFrame
        excel_path: Excelファイルのパス（画像抽出用）
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
        include_images: 画像を含めるか
    
    Returns:
        抽出された情報（Markdown形式）
    """
    # 画像を抽出
    images = []
    if include_images:
        images = extract_images_from_sheet(excel_path, sheet_name)
        if images:
            print(f'    Found {len(images)} images in sheet')
    
    # チャンク分割が不要な場合
    if not use_chunking or len(df) <= MAX_ROWS_PER_CHUNK:
        table_text = f"## シート: {sheet_name}\n\n{df.to_markdown()}"
        
        # メッセージコンテンツを作成（画像含む）
        user_content = create_message_content(table_text, images)
        
        response = client.chat.completions.create(
            model=DEPLOYMENT_NAME,
            messages=[
                {"role": "system", "content": PROMPT.format(basename=basename)},
                {"role": "user", "content": user_content}
            ],
            max_tokens=10000
        )
        return response.choices[0].message.content
    
    # チャンク分割して処理
    chunks = split_dataframe_into_chunks(df)
    chunk_results = []
    
    print(f'    Splitting into {len(chunks)} chunks...')
    
    for i, (start, end, chunk_df) in enumerate(chunks, 1):
        chunk_info = f"行{start+1}-{end}"
        print(f'    Processing chunk {i}/{len(chunks)}: {chunk_info}')
        
        try:
            table_text = f"## シート: {sheet_name} ({chunk_info})\n\n{chunk_df.to_markdown()}"
            
            # 最初のチャンクにのみ画像を含める（全チャンクに含めるとトークン超過の恐れ）
            chunk_images = images if i == 1 else []
            user_content = create_message_content(table_text, chunk_images)
            
            response = client.chat.completions.create(
                model=DEPLOYMENT_NAME,
                messages=[
                    {"role": "system", "content": PROMPT.format(basename=basename)},
                    {"role": "user", "content": user_content}
                ],
                max_tokens=10000
            )
            chunk_results.append(f"### {chunk_info}\n\n{response.choices[0].message.content}")
            
            if i < len(chunks):
                time.sleep(WAIT_TIME_BETWEEN_SHEETS)
                
        except Exception as e:
            print(f'    ⚠ Error processing chunk {chunk_info}: {str(e)}')
            chunk_results.append(f"### {chunk_info}\n\n**エラー**: {str(e)}")
    
    return "\n\n".join(chunk_results)


def extract_info(rel_path: str, resource_dir: str, client: AzureOpenAI, 
                use_chunking: bool = False, include_images: bool = True) -> str:
    """
    Excelファイルから情報を抽出（シート単位で処理、画像対応）
    
    Args:
        rel_path: resourceディレクトリからの相対パス（拡張子なし）
        resource_dir: resourceディレクトリのパス
        client: AzureOpenAIクライアント
        use_chunking: チャンク分割を使用するか
        include_images: 画像を含めるか
    
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
            sheet_info = extract_info_from_sheet(
                basename, sheet_name, df, excel_path, client, use_chunking, include_images
            )
            
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


def main(use_chunking: bool = False, include_images: bool = True):
    """
    メイン処理
    
    Args:
        use_chunking: 大きなシートをチャンク分割するか
        include_images: 画像を含めるか
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
    print(f'Include images: {"ON" if include_images else "OFF"}')
    
    if len(undescribeds) == 0:
        print('No files to process.')
        return
    
    client = get_client()
    
    for idx, undescribed in enumerate(undescribeds, 1):
        try:
            print(f'\n[{idx}/{len(undescribeds)}] 📄 Processing: {undescribed}')
            description = extract_info(undescribed, INPUT_DIR, client, use_chunking, include_images)
            
            output_path = os.path.join(OUTPUT_DIR, f'{undescribed}.md')
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(description)
            
            print(f'✓ Information file created: {undescribed}')
            
            if idx < len(undescribeds):
                time.sleep(WAIT_TIME_BETWEEN_FILES)
            
        except Exception as e:
            print(f'✗ Error processing {undescribed}: {str(e)}')
    
    print(f'\n{"="*60}')
    print(f'Processing complete!')


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Extract information from Excel files with image support')
    parser.add_argument('--chunking', action='store_true', 
                       help='Enable chunking for large sheets')
    parser.add_argument('--no-images', action='store_true',
                       help='Disable image extraction')
    
    args = parser.parse_args()
    
    main(use_chunking=args.chunking, include_images=not args.no_images)
