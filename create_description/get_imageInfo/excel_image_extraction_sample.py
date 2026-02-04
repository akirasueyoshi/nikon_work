"""
Excelファイルから画像を抽出する方法のサンプル
"""

import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
import io
import base64
from pathlib import Path


def extract_images_from_excel_openpyxl(excel_path: str, output_dir: str = "extracted_images"):
    """
    openpyxlを使用してExcelファイルから画像を抽出
    
    Args:
        excel_path: Excelファイルのパス
        output_dir: 画像の出力ディレクトリ
    
    Returns:
        {シート名: [(画像オブジェクト, 位置情報), ...]}の辞書
    """
    wb = openpyxl.load_workbook(excel_path)
    images_by_sheet = {}
    
    Path(output_dir).mkdir(exist_ok=True)
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        images_by_sheet[sheet_name] = []
        
        # シート内の画像を取得
        if hasattr(sheet, '_images'):
            for idx, img in enumerate(sheet._images):
                # 画像データを取得
                image_data = img._data()
                
                # PIL Imageに変換
                pil_image = Image.open(io.BytesIO(image_data))
                
                # ファイルとして保存
                image_filename = f"{sheet_name}_image_{idx}.png"
                image_path = Path(output_dir) / image_filename
                pil_image.save(image_path)
                
                # 位置情報を取得
                anchor = img.anchor
                position = {
                    'from_col': anchor._from.col if hasattr(anchor, '_from') else None,
                    'from_row': anchor._from.row if hasattr(anchor, '_from') else None,
                    'to_col': anchor.to.col if hasattr(anchor, 'to') else None,
                    'to_row': anchor.to.row if hasattr(anchor, 'to') else None,
                }
                
                images_by_sheet[sheet_name].append({
                    'image': pil_image,
                    'position': position,
                    'filename': image_filename
                })
                
                print(f"Extracted: {sheet_name} - {image_filename} at {position}")
    
    return images_by_sheet


def image_to_base64(image: Image.Image) -> str:
    """
    PIL ImageをBase64文字列に変換（LLMに渡す形式）
    
    Args:
        image: PIL Image
    
    Returns:
        Base64エンコードされた文字列
    """
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return img_str


def create_message_with_images(text_content: str, images: list[Image.Image]) -> list[dict]:
    """
    テキストと画像を含むメッセージを作成（Azure OpenAI形式）
    
    Args:
        text_content: テキスト内容
        images: PIL Imageのリスト
    
    Returns:
        Azure OpenAI APIに渡せるメッセージ形式
    """
    content = [
        {
            "type": "text",
            "text": text_content
        }
    ]
    
    # 画像を追加
    for img in images:
        base64_image = image_to_base64(img)
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{base64_image}"
            }
        })
    
    return content


# 使用例
if __name__ == "__main__":
    # Excelから画像を抽出
    excel_path = "sample.xlsx"
    images_by_sheet = extract_images_from_excel_openpyxl(excel_path)
    
    # 各シートの画像情報を表示
    for sheet_name, images in images_by_sheet.items():
        print(f"\n{sheet_name}: {len(images)} images")
        for img_info in images:
            print(f"  - {img_info['filename']} at {img_info['position']}")
