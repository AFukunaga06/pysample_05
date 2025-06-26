#!/usr/bin/env python3
"""
CustomTkinter OCR Application with JAN Code Detection
Fixed for headless environment compatibility
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import pyperclip
import time
import os
from PIL import Image, ImageTk
import tkinter as tk
import requests
import re
import json

# OCR関連ライブラリの動的インポート
try:
    import pytesseract
    import cv2
    import numpy as np
    OCR_AVAILABLE = True
    print("OCR機能が利用可能です")
    
    # Windows環境でTesseract-OCRのパスを設定
    if os.path.exists(r'C:\Program Files\Tesseract-OCR\tesseract.exe'):
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        print("Tesseract-OCRパスを設定しました")
    
except ImportError as e:
    OCR_AVAILABLE = False
    print(f"OCR機能が利用できません: {e}")
    print("必要なライブラリ: pip install pytesseract opencv-python")

# pyautoguiの動的インポート
try:
    import pyautogui
    SCREENSHOT_AVAILABLE = True
    print("スクリーンショット機能が利用可能です")
except ImportError:
    SCREENSHOT_AVAILABLE = False
    print("スクリーンショット機能が利用できません")
    print("必要なライブラリ: pip install pyautogui")

class CustomTkinterOCRApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("OCR機能テスト - JANコード商品情報抽出")
        
        # ウィンドウサイズを設定
        window_width = 450
        window_height = 700
        
        try:
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x_pos = (screen_width - window_width) // 2
            y_pos = (screen_height - window_height) // 2
        except:
            x_pos = 100
            y_pos = 100
        
        self.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
        self.resizable(True, True)
        
        # 変数の初期化
        self.screenshot = None
        self.selection_start = None
        self.selection_end = None
        self.left_monitor_region = None
        self.api_key = ""
        
        self.setup_ui()
        self.check_dependencies()

    def check_dependencies(self):
        """依存関係をチェックして状態を表示"""
        status_text = "システム状態:\n"
        
        if OCR_AVAILABLE:
            status_text += "✓ OCR機能: 利用可能\n"
        else:
            status_text += "✗ OCR機能: 利用不可\n"
            
        if SCREENSHOT_AVAILABLE:
            status_text += "✓ スクリーンショット: 利用可能\n"
        else:
            status_text += "✗ スクリーンショット: 利用不可\n"
            
        # Tesseract確認
        if OCR_AVAILABLE:
            try:
                # テスト用の小さな画像でTesseractをテスト
                test_img = Image.new('RGB', (100, 50), color='white')
                pytesseract.image_to_string(test_img)
                status_text += "✓ Tesseract-OCR: 正常動作\n"
                
                # 日本語サポート確認
                try:
                    langs = pytesseract.get_languages()
                    if 'jpn' in langs:
                        status_text += "✓ 日本語サポート: 利用可能\n"
                    else:
                        status_text += "⚠ 日本語サポート: 未確認\n"
                except:
                    status_text += "⚠ 日本語サポート: 未確認\n"
                    
            except Exception as e:
                status_text += f"✗ Tesseract-OCR: エラー\n({str(e)[:30]}...)\n"
        
        self.status_label.configure(text=status_text)

    def setup_ui(self):
        main_frame = ctk.CTkScrollableFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # タイトル
        title_label = ctk.CTkLabel(main_frame, text="OCR機能テスト", font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=10)

        # システム状態表示
        self.status_label = ctk.CTkLabel(
            main_frame, 
            text="システム状態を確認中...", 
            font=ctk.CTkFont(size=10),
            justify="left"
        )
        self.status_label.pack(pady=5)

        api_frame = ctk.CTkFrame(main_frame)
        api_frame.pack(pady=10, padx=10, fill="x")
        
        api_label = ctk.CTkLabel(api_frame, text="API設定", font=ctk.CTkFont(size=14, weight="bold"))
        api_label.pack(pady=5)
        
        self.api_entry = ctk.CTkEntry(api_frame, placeholder_text="Barcode Lookup API Key", width=300)
        self.api_entry.pack(pady=5)
        
        api_button = ctk.CTkButton(api_frame, text="API Key設定", command=self.set_api_key)
        api_button.pack(pady=5)

        # 使用方法の説明
        usage_frame = ctk.CTkFrame(main_frame)
        usage_frame.pack(pady=10, padx=10, fill="x")
        
        usage_label = ctk.CTkLabel(usage_frame, text="使用方法", font=ctk.CTkFont(size=14, weight="bold"))
        usage_label.pack(pady=5)
        
        usage_text = ctk.CTkLabel(
            usage_frame, 
            text="1. API Keyを設定\n2. スクリーンショット撮影または画像読み込み\n3. 範囲選択＆OCR実行\n4. JANコード検出と商品情報取得", 
            font=ctk.CTkFont(size=10),
            justify="left"
        )
        usage_text.pack(pady=5)

        # スクリーンショット機能
        screenshot_frame = ctk.CTkFrame(main_frame)
        screenshot_frame.pack(pady=10, padx=10, fill="x")
        
        screenshot_label = ctk.CTkLabel(screenshot_frame, text="スクリーンショット機能", font=ctk.CTkFont(size=14, weight="bold"))
        screenshot_label.pack(pady=5)

        screenshot_button = ctk.CTkButton(
            screenshot_frame,
            text="スクリーンショット撮影",
            command=self.take_screenshot,
            width=250,
            height=35
        )
        screenshot_button.pack(pady=5)

        # 画像読み込み機能
        load_frame = ctk.CTkFrame(main_frame)
        load_frame.pack(pady=10, padx=10, fill="x")
        
        load_label = ctk.CTkLabel(load_frame, text="画像読み込み機能", font=ctk.CTkFont(size=14, weight="bold"))
        load_label.pack(pady=5)

        load_button = ctk.CTkButton(
            load_frame,
            text="画像ファイルを読み込み",
            command=self.load_image_file,
            width=250,
            height=35
        )
        load_button.pack(pady=5)

        # OCR実行機能
        ocr_frame = ctk.CTkFrame(main_frame)
        ocr_frame.pack(pady=10, padx=10, fill="x")
        
        ocr_label = ctk.CTkLabel(ocr_frame, text="OCR実行機能", font=ctk.CTkFont(size=14, weight="bold"))
        ocr_label.pack(pady=5)

        select_button = ctk.CTkButton(
            ocr_frame,
            text="範囲選択＆OCR実行",
            command=self.select_area_and_ocr,
            width=250,
            height=35
        )
        select_button.pack(pady=5)

        result_frame = ctk.CTkFrame(main_frame)
        result_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        result_label = ctk.CTkLabel(result_frame, text="OCR結果", font=ctk.CTkFont(size=14, weight="bold"))
        result_label.pack(pady=5)

        self.result_text = ctk.CTkTextbox(result_frame, height=200, width=400)
        self.result_text.pack(pady=5, padx=10, fill="both", expand=True)

        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(pady=10, padx=10, fill="x")

        copy_button = ctk.CTkButton(button_frame, text="結果をコピー", command=self.copy_result)
        copy_button.pack(side="left", padx=5)

        save_button = ctk.CTkButton(button_frame, text="sample_01.txtに保存", command=self.save_to_sample_file)
        save_button.pack(side="left", padx=5)

        clear_button = ctk.CTkButton(button_frame, text="結果をクリア", command=self.clear_result)
        clear_button.pack(side="left", padx=5)

    def set_api_key(self):
        """API Keyを設定"""
        self.api_key = self.api_entry.get().strip()
        if self.api_key:
            messagebox.showinfo("完了", "API Keyが設定されました。")
        else:
            messagebox.showwarning("警告", "API Keyを入力してください。")

    def lookup_jan_code_info(self, jan_code):
        """JANコードから商品情報を取得"""
        if not self.api_key:
            return None
            
        try:
            url = f"https://api.barcodelookup.com/v3/products?barcode={jan_code}&formatted=y&key={self.api_key}"
            response = requests.get(url, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                if data.get('products') and len(data['products']) > 0:
                    product = data['products'][0]
                    return {
                        'jan_code': jan_code,
                        'brand': product.get('brand', ''),
                        'title': product.get('title', ''),
                        'category': product.get('category', ''),
                        'weight': product.get('weight', ''),
                        'dimensions': product.get('dimension', ''),
                        'description': product.get('description', '')
                    }
            return None
        except Exception as e:
            print(f"API呼び出しエラー: {e}")
            return None

    def extract_weight_from_text(self, text):
        """テキストから重量情報を抽出"""
        weight_patterns = [
            r'重量[：:\s]*(\d+(?:\.\d+)?)\s*[gG]',
            r'(\d+(?:\.\d+)?)\s*[gG](?:ラム)?',
            r'(\d+(?:\.\d+)?)\s*グラム',
            r'重さ[：:\s]*(\d+(?:\.\d+)?)\s*[gG]'
        ]
        
        for pattern in weight_patterns:
            matches = re.findall(pattern, text)
            if matches:
                return f"{matches[0]}g"
        return ""

    def extract_dimensions_from_text(self, text):
        """テキストから寸法情報を抽出"""
        dimension_patterns = [
            r'幅\d+.*?高さ\d+.*?奥行き\d+.*?mm',
            r'\d+×\d+×\d+\s*mm',
            r'サイズ[：:\s]*([^。\n]+)',
            r'寸法[：:\s]*([^。\n]+)'
        ]
        
        for pattern in dimension_patterns:
            matches = re.findall(pattern, text)
            if matches:
                return matches[0]
        return ""

    def detect_jan_code_in_text(self, text):
        """テキストからJANコードを検出"""
        jan_patterns = [
            r'\b(\d{13})\b',
            r'\b(\d{8})\b',
            r'JAN[：:\s]*(\d{8,13})',
            r'JANコード[：:\s]*(\d{8,13})'
        ]
        
        for pattern in jan_patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                if len(match) in [8, 13]:
                    return match
        return None

    def take_screenshot(self):
        """スクリーンショットを撮影（headless環境対応）"""
        if not SCREENSHOT_AVAILABLE:
            messagebox.showerror("エラー", "スクリーンショット機能が利用できません。\npyautoguiをインストールしてください。")
            return
            
        try:
            self.withdraw()
            time.sleep(0.5)
            
            try:
                self.screenshot = pyautogui.screenshot()
                
                try:
                    screen_width = pyautogui.size().width
                    screen_height = pyautogui.size().height
                    
                    if screen_width > 1920:  # 複数モニターの可能性
                        choice = messagebox.askyesnocancel(
                            "ディスプレイ選択", 
                            f"メインディスプレイ ({screen_width}x{screen_height}) を検出しました。\n\n左モニターを使用しますか？\n（はい=左モニター、いいえ=メインディスプレイ、キャンセル=中止）"
                        )
                        
                        if choice is None:  # キャンセル
                            self.deiconify()
                            return
                        elif choice:  # 左モニター
                            try:
                                self.screenshot = pyautogui.screenshot(region=(-1920, 0, 1920, 1080))
                                self.left_monitor_region = (-1920, 0, 1920, 1080)
                            except Exception as e:
                                print(f"左モニタースクリーンショットエラー: {e}")
                                self.screenshot = pyautogui.screenshot()
                                self.left_monitor_region = None
                        
                except Exception as e:
                    print(f"マルチディスプレイ処理エラー: {e}")
                    
            except Exception as e:
                print(f"スクリーンショットエラー: {e}")
                self.screenshot = None
                self.deiconify()
                messagebox.showerror("エラー", f"スクリーンショット撮影に失敗しました:\n{str(e)}")
                return
            
            if self.screenshot is None:
                self.deiconify()
                messagebox.showerror("エラー", "スクリーンショットの撮影に失敗しました。")
                return
                
            self.deiconify()
            messagebox.showinfo("完了", "スクリーンショットを撮影しました。\n「範囲選択＆OCR実行」ボタンで範囲を選択してください。")
            
        except Exception as e:
            self.deiconify()
            messagebox.showerror("エラー", f"スクリーンショット撮影エラー:\n{str(e)}")

    def load_image_file(self):
        """画像ファイルを読み込み"""
        try:
            file_path = filedialog.askopenfilename(
                title="画像ファイルを選択",
                filetypes=[
                    ("画像ファイル", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff"),
                    ("すべてのファイル", "*.*")
                ]
            )
            
            if file_path:
                self.screenshot = Image.open(file_path)
                messagebox.showinfo("完了", f"画像ファイルを読み込みました:\n{os.path.basename(file_path)}")
                
        except Exception as e:
            messagebox.showerror("エラー", f"画像読み込みエラー:\n{str(e)}")

    def select_area_and_ocr(self):
        """範囲選択とOCR実行"""
        if not OCR_AVAILABLE:
            messagebox.showerror("エラー", "OCR機能が利用できません。")
            return
            
        if not self.screenshot:
            messagebox.showerror("エラー", "スクリーンショットまたは画像が読み込まれていません。")
            return
            
        self.open_selection_window()

    def open_selection_window(self):
        """選択ウィンドウを開く（安全性向上）"""
        if not self.screenshot:
            messagebox.showerror("エラー", "スクリーンショットが撮影されていません。")
            return
            
        try:
            if not hasattr(self.screenshot, 'width') or not hasattr(self.screenshot, 'height'):
                messagebox.showerror("エラー", "スクリーンショットデータが無効です。")
                return
                
            self.selection_window = tk.Toplevel(self)
            self.selection_window.title("範囲選択")
            self.selection_window.attributes('-topmost', True)
            
            try:
                screen_width = self.selection_window.winfo_screenwidth()
                screen_height = self.selection_window.winfo_screenheight()
            except:
                screen_width = 1024
                screen_height = 768
            
            max_width = int(screen_width * 0.8)
            max_height = int(screen_height * 0.8)
            
            img_ratio = min(max_width / self.screenshot.width, max_height / self.screenshot.height)
            new_width = int(self.screenshot.width * img_ratio)
            new_height = int(self.screenshot.height * img_ratio)
            
            self.display_image = self.screenshot.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.photo = ImageTk.PhotoImage(self.display_image)
            
            self.canvas = tk.Canvas(self.selection_window, width=new_width, height=new_height)
            self.canvas.pack()
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            
            self.scale_ratio = img_ratio
            
            self.canvas.bind("<Button-1>", self.on_selection_start)
            self.canvas.bind("<B1-Motion>", self.on_selection_drag)
            self.canvas.bind("<ButtonRelease-1>", self.on_selection_end)
            
            self.selection_window.geometry(f"{new_width}x{new_height}")
            self.selection_window.resizable(False, False)
            
        except Exception as e:
            messagebox.showerror("エラー", f"選択ウィンドウの作成に失敗しました:\n{str(e)}")
            if hasattr(self, 'selection_window'):
                self.selection_window.destroy()

    def on_selection_start(self, event):
        """選択開始"""
        self.selection_start = (event.x, event.y)

    def on_selection_drag(self, event):
        """選択中のドラッグ"""
        if self.selection_start:
            self.canvas.delete("selection")
            self.canvas.create_rectangle(
                self.selection_start[0], self.selection_start[1],
                event.x, event.y,
                outline="red", width=2, tags="selection"
            )

    def on_selection_end(self, event):
        """選択終了"""
        self.selection_end = (event.x, event.y)
        self.selection_window.destroy()
        self.perform_ocr()

    def perform_ocr(self):
        """選択範囲のOCR処理を実行（安全性向上）"""
        if not self.screenshot or not self.selection_start or not self.selection_end:
            messagebox.showerror("エラー", "範囲が選択されていません。")
            return
            
        try:
            if not hasattr(self.screenshot, 'crop'):
                messagebox.showerror("エラー", "スクリーンショットデータが無効です。")
                return
                
            x1 = int(min(self.selection_start[0], self.selection_end[0]) / self.scale_ratio)
            y1 = int(min(self.selection_start[1], self.selection_end[1]) / self.scale_ratio)
            x2 = int(max(self.selection_start[0], self.selection_end[0]) / self.scale_ratio)
            y2 = int(max(self.selection_start[1], self.selection_end[1]) / self.scale_ratio)
            
            cropped_image = self.screenshot.crop((x1, y1, x2, y2))
            
            opencv_image = cv2.cvtColor(np.array(cropped_image), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            kernel = np.ones((1,1), np.uint8)
            processed = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, kernel)
            processed = cv2.medianBlur(processed, 3)
            
            _, binary = cv2.threshold(processed, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            try:
                text = pytesseract.image_to_string(
                    binary, 
                    config=r'--oem 3 --psm 6 -l jpn+eng',
                    lang='jpn+eng'
                )
            except:
                try:
                    text = pytesseract.image_to_string(
                        binary, 
                        config=r'--oem 3 --psm 6 -l jpn',
                        lang='jpn'
                    )
                except:
                    text = pytesseract.image_to_string(binary, config=r'--oem 3 --psm 6')
            
            self.result_text.delete("1.0", "end")
            
            jan_code = self.detect_jan_code_in_text(text)
            if jan_code:
                product_info = self.lookup_jan_code_info(jan_code)
                if product_info:
                    formatted_result = self.format_product_info(product_info, text)
                    self.result_text.insert("1.0", formatted_result)
                    pyperclip.copy(formatted_result)
                    messagebox.showinfo("完了", f"JANコード {jan_code} の商品情報を取得しました。")
                else:
                    ocr_result = self.format_ocr_result(text, jan_code)
                    self.result_text.insert("1.0", ocr_result)
                    pyperclip.copy(ocr_result)
                    messagebox.showinfo("完了", f"JANコード {jan_code} を検出しましたが、API情報取得に失敗しました。OCR結果を表示します。")
            else:
                self.result_text.insert("1.0", text.strip())
                pyperclip.copy(text.strip())
                messagebox.showinfo("完了", "OCR処理が完了しました。JANコードは検出されませんでした。")
                
        except Exception as e:
            messagebox.showerror("エラー", f"OCR処理エラー:\n{str(e)}")

    def format_product_info(self, product_info, ocr_text):
        """API取得した商品情報を指定フォーマットで整形"""
        jan_code = product_info.get('jan_code', '')
        brand = product_info.get('brand', '')
        title = product_info.get('title', '')
        category = product_info.get('category', '')
        
        weight = product_info.get('weight', '')
        if not weight:
            weight = self.extract_weight_from_text(ocr_text)
        
        dimensions = product_info.get('dimensions', '')
        if not dimensions:
            dimensions = self.extract_dimensions_from_text(ocr_text)
        
        formatted_lines = []
        formatted_lines.append(f"JANコード\t{jan_code}")
        formatted_lines.append(f"ブランド名\t{brand}")
        formatted_lines.append(f"商品名\t{title}")
        formatted_lines.append(f"規格\t{category}")
        formatted_lines.append(f"商品サイズ\t{dimensions}")
        
        if weight:
            if weight.endswith('g'):
                formatted_lines.append(f"重量{weight}")
            else:
                formatted_lines.append(f"重量{weight}g")
        else:
            formatted_lines.append("重量")
        
        return '\n'.join(formatted_lines)

    def format_ocr_result(self, text, jan_code=None):
        """OCR結果を指定フォーマットで整形"""
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        formatted_lines = []
        if jan_code:
            formatted_lines.append(f"JANコード\t{jan_code}")
        
        brand = ""
        product_name = ""
        specification = ""
        
        for line in lines:
            if any(keyword in line for keyword in ['千吉', 'ブランド', '金貴', '株式会社', '藤原産業', 'E-Value']):
                brand = line
            elif any(keyword in line for keyword in ['根力', 'マルチドライバー', 'セット', 'ドライバー']):
                product_name = line
            elif any(keyword in line for keyword in ['本', 'セット', '個', 'ホン', 'コ', 'ERD']):
                specification = line
        
        formatted_lines.append(f"ブランド名\t{brand}")
        formatted_lines.append(f"商品名\t{product_name}")
        formatted_lines.append(f"規格\t{specification}")
        
        dimensions = self.extract_dimensions_from_text(text)
        formatted_lines.append(f"商品サイズ\t{dimensions}")
        
        weight = self.extract_weight_from_text(text)
        if weight:
            formatted_lines.append(f"重量{weight}")
        else:
            formatted_lines.append("重量")
        
        return '\n'.join(formatted_lines)

    def copy_result(self):
        """結果をクリップボードにコピー"""
        text = self.result_text.get("1.0", "end-1c")
        if text.strip():
            pyperclip.copy(text.strip())
            messagebox.showinfo("完了", "結果をクリップボードにコピーしました。")
        else:
            messagebox.showwarning("警告", "コピーするテキストがありません。")

    def save_to_sample_file(self):
        """結果をsample_01.txtに保存"""
        text = self.result_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning("警告", "保存するテキストがありません。")
            return
            
        try:
            with open('sample_01.txt', 'w', encoding='utf-8') as f:
                f.write(text.strip())
            messagebox.showinfo("完了", "sample_01.txtに保存しました。")
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル保存エラー:\n{str(e)}")

    def clear_result(self):
        """結果をクリア"""
        self.result_text.delete("1.0", "end")

if __name__ == '__main__':
    print("=== CustomTkinter OCR機能テスト - JANコード商品情報抽出 ===")
    print("起動中...")
    
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    
    app = CustomTkinterOCRApp()
    app.mainloop()
