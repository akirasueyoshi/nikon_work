import os
from azure.identity import InteractiveBrowserCredential, get_bearer_token_provider
from openai import AzureOpenAI

AZURE_OPENAI_ENDPOINT = 'https://aoai-je-exm.openai.azure.com/'
DEPLOYMENT_NAME = 'gpt-4o'

PROMPT = """
以降のユーザープロンプトで示されるテキストは半導体露光装置における{basename}という機能仕様書から「明示的に記載されている」事実のみを項目別に抽出したものです。
抽出した事実と抽出ルールは以下の通りです。

## 抽出した事実の種類
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

この記述が機能の要求レベル、仕様レベル、詳細仕様レベルかを出力フォーマットに従って振り分けてください。

## 出力フォーマット
### 要求
- [処理内容] 
- [入出力] 
- [状態遷移] 
- [制約・前提条件] 
- [依存・関連する機能] 
### 仕様
- [処理内容] 
- [入出力] 
- [状態遷移] 
- [制約・前提条件] 
- [依存・関連する機能] 
### 詳細仕様
- [処理内容] 
- [入出力] 
- [状態遷移] 
- [制約・前提条件] 
- [依存・関連する機能] 

回答はMarkdown形式の記述としてください。

"""

INPUT_DIR = r'information'
OUTPUT_DIR = r'organized_info'


def get_markdown_files(dir_path: str) -> list[str]:
    """
    指定ディレクトリ配下のすべてのMarkdownファイルを再帰的に取得
    ルートディレクトリからの相対パス（拡張子なし）のリストを返す
    """
    markdown_files = []
    
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
        return markdown_files
    
    # os.walkで再帰的にディレクトリを探索
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            # Markdownファイルのみを対象
            if file.endswith('.md'):
                # 絶対パスを取得
                full_path = os.path.join(root, file)
                # ルートディレクトリからの相対パスを取得
                rel_path = os.path.relpath(full_path, dir_path)
                # 拡張子を除いた相対パスを保存
                rel_path_without_ext = os.path.splitext(rel_path)[0]
                markdown_files.append(rel_path_without_ext)
    
    return markdown_files


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


def organize_info(rel_path: str, resource_dir: str, client: AzureOpenAI) -> str:
    """
    Markdownファイルから情報を抽出して整理
    
    Args:
        rel_path: informationディレクトリからの相対パス（拡張子なし）
        resource_dir: informationディレクトリのパス
        client: AzureOpenAIクライアント
    """
    # ファイル名（ベース名）を取得
    basename = os.path.basename(rel_path)
    
    # Markdownファイルのフルパスを構築
    input_path = os.path.join(resource_dir, f'{rel_path}.md')
    
    # Markdownファイルを読み込む
    with open(input_path, 'r', encoding='utf-8') as f:
        markdown_text = f.read()

    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {"role": "system", "content": PROMPT.format(basename=basename)},
            {"role": "user", "content": markdown_text}
        ],
        max_tokens=10000
    )
    return response.choices[0].message.content


def main():
    # 入力ディレクトリと出力ディレクトリからファイルリストを取得
    inputs = get_markdown_files(INPUT_DIR)
    outputs = get_markdown_files(OUTPUT_DIR)
    
    # 未処理のファイルを抽出
    undescribeds = list(set(inputs).difference(outputs))
    
    print(f'Found {len(inputs)} markdown files, {len(outputs)} already processed.')
    print(f'Processing {len(undescribeds)} files...')
    
    client = get_client()
    
    for undescribed in undescribeds:
        try:
            output = organize_info(undescribed, INPUT_DIR, client)
            
            # 出力ファイルのパスを構築
            output_path = os.path.join(OUTPUT_DIR, f'{undescribed}.md')
            
            # 出力ディレクトリを作成（存在しない場合）
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # 整理したMarkdownファイルを保存
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(output)
            
            print(f'✓ Organized information file created: {undescribed}')
            
        except Exception as e:
            print(f'✗ Error processing {undescribed}: {str(e)}')


if __name__ == "__main__":
    main()
