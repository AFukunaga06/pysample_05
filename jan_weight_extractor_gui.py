import customtkinter as ctk
from tkinter import messagebox
import pyautogui
import pyperclip
import time
import re
import os

class ProductInfoExtractor:
    def __init__(self):
        self.jan_pattern = r'JANコード[\s\t]*(\d{13}|\d{8})'
        self.weight_pattern = r'重量(\d+(?:\.\d+)?)[ｇg]'
        self.brand_pattern = r'ブランド名[\s\t]*([^\n\r]+)'
        self.product_name_pattern = r'商品名[\s\t]*([^\n\r]+)'
        self.spec_pattern = r'規格[\s\t]*([^\n\r]+)'
        self.size_pattern = r'商品サイズ[\s\t]*([^\n\r]+)'

    def extract_product_info(self, text):
        jan_match = re.search(self.jan_pattern, text)
        if not jan_match:
            return None
        jan_code = jan_match.group(1)
        start_pos = jan_match.start()
        tail = text[start_pos:]
        weight_match = re.search(self.weight_pattern, tail)
        if not weight_match:
            return None
        end_pos = weight_match.end()
        segment = tail[:end_pos]

        brand_match = re.search(self.brand_pattern, segment)
        product_name_match = re.search(self.product_name_pattern, segment)
        spec_match = re.search(self.spec_pattern, segment)
        size_match = re.search(self.size_pattern, segment)

        return {
            'jan_code': jan_code,
            'brand_name': brand_match.group(1).strip() if brand_match else '',
            'product_name': product_name_match.group(1).strip() if product_name_match else '',
            'specification': spec_match.group(1).strip() if spec_match else '',
            'product_size': size_match.group(1).strip() if size_match else '',
            'weight': weight_match.group(1) + weight_match.group(0)[-1],
            'extracted_text': segment.strip()
        }

    def format_product_info(self, info):
        if not info:
            return "商品情報が見つかりませんでした。"
        lines = [f"JANコード\t{info['jan_code']}"]
        if info['brand_name']:
            lines.append(f"ブランド名\t{info['brand_name']}")
        if info['product_name']:
            lines.append(f"商品名\t{info['product_name']}")
        if info['specification']:
            lines.append(f"規格\t{info['specification']}")
        if info['product_size']:
            lines.append(f"商品サイズ\t{info['product_size']}")
        lines.append(f"重量{info['weight']}")
        return "\n".join(lines)

class ExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("商品情報抽出ツール")
        self.geometry("400x300")
        self.extractor = ProductInfoExtractor()
        self._build_ui()

    def _build_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.instruction_label = ctk.CTkLabel(
            self, text="任意の範囲をドラッグ選択し、抽出します" , font=("Arial", 12)
        )
        self.instruction_label.pack(pady=10)

        self.select_button = ctk.CTkButton(
            self, text="範囲選択＆抽出", command=self.on_select_extract
        )
        self.select_button.pack(pady=5)

        self.result_box = ctk.CTkTextbox(self, width=380, height=150)
        self.result_box.pack(padx=10, pady=10)

    def on_select_extract(self):
        messagebox.showinfo(
            "操作方法",
            "抽出したい範囲をマウスでドラッグ選択してください。\n選択が完了したら OK を押してください。"
        )
        # 選択範囲をクリップボードにコピー
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.2)
        full_text = pyperclip.paste()

        info = self.extractor.extract_product_info(full_text)
        formatted = self.extractor.format_product_info(info)

        # テキストボックスに結果表示
        self.result_box.delete("0.0", "end")
        self.result_box.insert("0.0", formatted)

        # ファイル保存
        path = os.path.join(os.path.dirname(__file__), "input.txt")
        with open(path, 'w', encoding='utf-8') as f:
            f.write(formatted)

        messagebox.showinfo(
            "完了",
            f"抽出結果を保存しました:\n{path}"
        )

if __name__ == '__main__':
    app = ExtractorApp()
    app.mainloop()
