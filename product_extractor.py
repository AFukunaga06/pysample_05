import re
import pyperclip  # pip install pyperclip
import os
from datetime import datetime

class ProductInfoExtractor:
    def __init__(self):
        self.jan_pattern = r'JANコード[\s\t]*(\d{13}|\d{8})'
        self.weight_pattern = r'重量(\d+(?:\.\d+)?)[ｇg]'
        self.brand_pattern = r'ブランド名[\s\t]*([^\n\r]+)'
        self.product_name_pattern = r'商品名[\s\t]*([^\n\r]+)'
        self.spec_pattern = r'規格[\s\t]*([^\n\r]+)'
        self.size_pattern = r'商品サイズ[\s\t]*([^\n\r]+)'
    
    def extract_product_info(self, text):
        """
        テキストから商品情報を抽出する
        """
        # JANコードを検索
        jan_match = re.search(self.jan_pattern, text)
        if not jan_match:
            return None
        
        jan_code = jan_match.group(1)
        jan_index = text.find(jan_match.group(0))
        
        # JANコードから重量までのテキストを取得
        after_jan_text = text[jan_index:]
        
        # 重量を検索
        weight_match = re.search(self.weight_pattern, after_jan_text)
        if not weight_match:
            return None
        
        weight_index = after_jan_text.find(weight_match.group(0))
        extracted_text = after_jan_text[:weight_index + len(weight_match.group(0))]
        
        # 各項目を抽出
        brand_match = re.search(self.brand_pattern, extracted_text)
        product_name_match = re.search(self.product_name_pattern, extracted_text)
        spec_match = re.search(self.spec_pattern, extracted_text)
        size_match = re.search(self.size_pattern, extracted_text)
        
        return {
            'jan_code': jan_code,
            'brand_name': brand_match.group(1).strip() if brand_match else '',
            'product_name': product_name_match.group(1).strip() if product_name_match else '',
            'specification': spec_match.group(1).strip() if spec_match else '',
            'product_size': size_match.group(1).strip() if size_match else '',
            'weight': weight_match.group(1) + weight_match.group(0)[-1],  # 数値＋単位
            'extracted_text': extracted_text.strip()
        }
    
    def format_product_info(self, product_info):
        """
        商品情報を指定された形式で整形する
        """
        if not product_info:
            return "商品情報が見つかりませんでした。"
        
        result = f"JANコード\t{product_info['jan_code']}"
        
        if product_info['brand_name']:
            result += f"\nブランド名\t{product_info['brand_name']}"
        if product_info['product_name']:
            result += f"\n商品名\t{product_info['product_name']}"
        if product_info['specification']:
            result += f"\n規格\t{product_info['specification']}"
        if product_info['product_size']:
            result += f"\n商品サイズ\t{product_info['product_size']}"
        
        result += f"\n重量{product_info['weight']}"
        
        return result
    
    def read_from_clipboard(self):
        """
        クリップボードからテキストを読み取る
        """
        try:
            return pyperclip.paste()
        except Exception as e:
            print(f"クリップボードからの読み取りに失敗しました: {e}")
            return None
    
    def save_to_file(self, text, filename="sample0626_01.txt"):
        """
        テキストをファイルに保存する（現在のディレクトリ）
        """
        try:
            # 現在のスクリプトと同じディレクトリに保存
            current_dir = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(current_dir, filename)
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(text)
            
            print(f"ファイル '{filename}' を保存しました: {file_path}")
            return True
        except Exception as e:
            print(f"ファイル保存に失敗しました: {e}")
            return False
    
    def process_clipboard_to_file(self, filename="sample0626_01.txt"):
        """
        クリップボードから読み取り、商品情報を抽出してファイルに保存
        """
        print("クリップボードからテキストを読み取り中...")
        
        # クリップボードからテキストを読み取り
        clipboard_text = self.read_from_clipboard()
        
        if not clipboard_text:
            print("クリップボードにテキストがありません")
            return None
        
        print("クリップボードのテキスト:")
        print(clipboard_text)
        print("=" * 50)
        
        # 商品情報を抽出
        product_info = self.extract_product_info(clipboard_text)
        formatted_result = self.format_product_info(product_info)
        
        print("抽出された商品情報:")
        print(formatted_result)
        print("=" * 50)
        
        # ファイルに保存
        success = self.save_to_file(formatted_result, filename)
        
        if success:
            print(f"{filename}ファイルに保存完了！")
        
        return formatted_result
    
    def process_multiple_products(self, text, filename="sample0626_01.txt"):
        """
        複数の商品情報を処理してファイルに保存
        """
        # JANコードで分割
        jan_sections = re.split(r'(?=JANコード)', text)
        results = []
        
        for i, section in enumerate(jan_sections):
            if section.strip() and 'JANコード' in section:
                product_info = self.extract_product_info(section)
                if product_info:
                    formatted = self.format_product_info(product_info)
                    results.append(f"商品 {len(results) + 1}:\n{formatted}")
        
        if not results:
            # 単一商品として処理
            product_info = self.extract_product_info(text)
            formatted = self.format_product_info(product_info)
            self.save_to_file(formatted, filename)
            return [formatted]
        else:
            # 複数商品をまとめて保存
            combined_result = '\n\n' + ('=' * 50) + '\n\n'.join([''] + results)
            self.save_to_file(combined_result, filename)
            print(f"{len(results)}個の商品情報を処理しました")
            return results

# 使用例
if __name__ == "__main__":
    extractor = ProductInfoExtractor()
    
    # 方法1: クリップボードから自動処理
    print("方法1: クリップボードから自動処理")
    extractor.process_clipboard_to_file()
    
    print("\n" + "="*60 + "\n")
    
    # 方法2: テキストを直接指定
    print("方法2: テキストを直接指定")
    sample_text = """狭い部分や細かい草の除草作業が楽々できます。
千吉金賞 根カキ 2ﾎﾝ
JANコード	4977292622196
ブランド名	千吉金賞
商品名	根カキ
規格	2ﾎﾝ
商品サイズ	幅50×高さ210×奥行き30mm
重量120ｇ
メーカーHP
取扱説明書"""
    
    product_info = extractor.extract_product_info(sample_text)
    formatted = extractor.format_product_info(product_info)
    print(formatted)
    
    # ファイルに保存
    extractor.save_to_file(formatted, "sample0626_01.txt")

# 簡単な実行用関数
def quick_process():
    """簡単に実行できる関数"""
    extractor = ProductInfoExtractor()
    return extractor.process_clipboard_to_file()

# 使用方法:
# 1. 商品情報をコピー (Ctrl+C)
# 2. Pythonスクリプトを実行
# 3. 同じディレクトリにsample0626_01.txtが作成される