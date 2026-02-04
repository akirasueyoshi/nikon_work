import os
from azure.identity import InteractiveBrowserCredential, get_bearer_token_provider
from openai import AzureOpenAI

AZURE_OPENAI_ENDPOINT = 'https://aoai-je-exm.openai.azure.com/'
DEPLOYMENT_NAME = 'gpt-4o'

PROMPT = """
以降のユーザープロンプトで示されるテキストは半導体露光装置における{basename}という機能仕様書から「明示的に記載されている」事実のみを項目別に抽出したものです。
抽出した事実は以下の通りです。

以下は「{basename}」に関する事実群である (トピック・抽象度別に整理済み)。

## 整理された事実
{markdown_text}

## タスク
これらの事実を統合し、「{basename}」の現行仕様の解説ドキュメントを作成せよ。

## 解説の構成
1. 機能概要（目的、全体像）
2. インターフェース仕様（入出力）
3. 処理フロー
4. 制約・エラー処理
5. 依存関係・前提条件

## 制約
- 上記の事実に含まれない情報は記載しない
- 矛盾がある箇所は両論併記し、出典を明記する
- 推測や補完は行わない


"""

INPUT_DIR = r'organized_info'
OUTPUT_DIR = r'synthesized_info'


def get_directories_with_md_files(dir_path: str) -> list[str]:
    """
    指定ディレクトリ配下でマークダウンファイルを含むディレクトリを再帰的に取得
    ルートディレクトリからの相対パスのリストを返す
    """
    directories = []
    
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
        return directories
    
    # os.walkで再帰的にディレクトリを探索
    for root, dirs, files in os.walk(dir_path):
        # このディレクトリ内にマークダウンファイルがあるかチェック
        has_md = any(file.endswith('.md') for file in files)
        
        if has_md:
            # ルートディレクトリからの相対パスを取得
            rel_path = os.path.relpath(root, dir_path)
            # カレントディレクトリ自体の場合は空文字になるので、その場合は除外
            if rel_path != '.':
                directories.append(rel_path)
    
    return directories


def get_markdown_files_in_directory(dir_path: str) -> list[str]:
    """
    指定ディレクトリ直下のマークダウンファイルを取得
    """
    md_files = []
    
    if not os.path.exists(dir_path):
        return md_files
    
    for file in os.listdir(dir_path):
        if file.endswith('.md'):
            md_files.append(os.path.join(dir_path, file))
    
    return md_files


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


def synthesize_info(rel_dir_path: str, input_dir: str, client: AzureOpenAI) -> str:
    """
    ディレクトリ内のすべてのマークダウンファイルを統合して解説書を生成
    
    Args:
        rel_dir_path: organized_infoディレクトリからの相対ディレクトリパス
        input_dir: organized_infoディレクトリのパス
        client: AzureOpenAIクライアント
    """
    # ディレクトリ名（機能名）を取得
    basename = os.path.basename(rel_dir_path)
    
    # ディレクトリのフルパスを構築
    full_dir_path = os.path.join(input_dir, rel_dir_path)
    
    # ディレクトリ内のすべてのマークダウンファイルを取得
    md_files = get_markdown_files_in_directory(full_dir_path)
    
    if not md_files:
        print(f'  Warning: No markdown files found in {rel_dir_path}')
        return ""
    
    # すべてのマークダウンファイルを読み込んで統合
    combined_text = ""
    for md_file in sorted(md_files):  # ファイル名順にソート
        file_name = os.path.basename(md_file)
        with open(md_file, 'r', encoding='utf-8') as f:
            content = f.read()
            # ファイルごとに区切りを入れる
            combined_text += f"\n\n## ===== {file_name} =====\n\n{content}\n"
    
    print(f'  Processing {len(md_files)} files from {rel_dir_path}')
    
    # AIで解説書を生成
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        messages=[
            {"role": "system", "content": PROMPT.format(basename=basename, markdown_text=combined_text)},
            {"role": "user", "content": f"以下の{len(md_files)}個のファイルから統合された情報:\n\n{combined_text}"}
        ],
        max_tokens=10000
    )
    return response.choices[0].message.content


def main():
    # マークダウンファイルを含むディレクトリを取得
    input_dirs = get_directories_with_md_files(INPUT_DIR)
    
    # すでに処理済みの解説書を取得
    output_files = []
    if os.path.exists(OUTPUT_DIR):
        for root, dirs, files in os.walk(OUTPUT_DIR):
            for file in files:
                if file.endswith('.md'):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, OUTPUT_DIR)
                    # 拡張子を除いたパスを保存（ディレクトリパスに相当）
                    rel_path_without_ext = os.path.splitext(rel_path)[0]
                    output_files.append(rel_path_without_ext)
    
    # 未処理のディレクトリを抽出
    undescribeds = list(set(input_dirs).difference(output_files))
    
    print(f'Found {len(input_dirs)} directories with markdown files.')
    print(f'Already processed: {len(output_files)}')
    print(f'Processing {len(undescribeds)} directories...\n')
    
    client = get_client()
    
    for undescribed in sorted(undescribeds):
        try:
            print(f'Processing: {undescribed}')
            
            # 解説書を生成
            output = synthesize_info(undescribed, INPUT_DIR, client)
            
            if not output:
                continue
            
            # 出力ファイルのパスを構築
            # 親ディレクトリパス + ディレクトリ名.md
            parent_dir = os.path.dirname(undescribed)
            dir_name = os.path.basename(undescribed)
            
            if parent_dir:
                output_file_path = os.path.join(OUTPUT_DIR, parent_dir, f'{dir_name}.md')
            else:
                output_file_path = os.path.join(OUTPUT_DIR, f'{dir_name}.md')
            
            # 出力ディレクトリを作成（存在しない場合）
            output_dir = os.path.dirname(output_file_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # 解説書を保存
            with open(output_file_path, 'w', encoding='utf-8') as f:
                f.write(output)
            
            print(f'✓ Description file created: {output_file_path}\n')
            
        except Exception as e:
            print(f'✗ Error processing {undescribed}: {str(e)}\n')


if __name__ == "__main__":
    main()
