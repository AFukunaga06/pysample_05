#!/usr/bin/env python3
"""
CustomTkinter OCR Application - Free OCR-Only Version
Optimized for JAN code detection without paid API services
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import pyperclip
import time
import os
from PIL import Image, ImageTk
import tkinter as tk
import re

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

class CustomTkinterOCRFreeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("OCR機能テスト - JANコード検出（無料版）")
        
        self.geometry("600x800")
        self.minsize(500, 600)
        
        self.screenshot = None
        self.selection_start = None
        self.selection_end = None
        self.scale_ratio = 1.0
        
        self.setup_ui()
        self.check_dependencies()

    def check_dependencies(self):
        """依存関係の確認とシステム状態表示"""
        status_text = "=== システム状態 ===\n"
        
        if OCR_AVAILABLE:
            status_text += "✓ OCR機能: 利用可能\n"
            
            try:
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
        else:
            status_text += "✗ OCR機能: 利用不可\n"
        
        if SCREENSHOT_AVAILABLE:
            status_text += "✓ スクリーンショット: 利用可能\n"
        else:
            status_text += "✗ スクリーンショット: 利用不可\n"
        
        status_text += "✓ 無料版: API不要で完全動作\n"
        
        self.status_label.configure(text=status_text)

    def setup_ui(self):
        main_frame = ctk.CTkScrollableFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # タイトル
        title_label = ctk.CTkLabel(main_frame, text="OCR機能テスト（無料版）", font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=10)

        # システム状態表示
        self.status_label = ctk.CTkLabel(
            main_frame, 
            text="システム状態を確認中...", 
            font=ctk.CTkFont(size=10),
            justify="left"
        )
        self.status_label.pack(pady=5)

        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(pady=10, padx=10, fill="x")
        
        info_label = ctk.CTkLabel(info_frame, text="無料版の特徴", font=ctk.CTkFont(size=14, weight="bold"))
        info_label.pack(pady=5)
        
        info_text = ctk.CTkLabel(
            info_frame, 
            text="• JANコード自動検出\n• 重量・サイズ情報抽出\n• sample_01.txt出力\n• API不要で完全動作", 
            font=ctk.CTkFont(size=10),
            justify="left"
        )
        info_text.pack(pady=5)

        # 使用方法の説明
        usage_frame = ctk.CTkFrame(main_frame)
        usage_frame.pack(pady=10, padx=10, fill="x")
        
        usage_label = ctk.CTkLabel(usage_frame, text="使用方法", font=ctk.CTkFont(size=14, weight="bold"))
        usage_label.pack(pady=5)
        
        usage_text = ctk.CTkLabel(
            usage_frame, 
            text="1. スクリーンショット撮影または画像読み込み\n2. 範囲選択＆OCR実行\n3. JANコード・重量・サイズ自動検出\n4. sample_01.txt保存", 
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

    def extract_weight_from_text(self, text):
        """テキストから重量情報を抽出（強化版）"""
        weight_patterns = [
            r'重量[：:\s]*(\d+(?:\.\d+)?)\s*[gG]',
            r'(\d+(?:\.\d+)?)\s*[gG](?:ラム)?',
            r'(\d+(?:\.\d+)?)\s*グラム',
            r'重さ[：:\s]*(\d+(?:\.\d+)?)\s*[gG]',
            r'質量[：:\s]*(\d+(?:\.\d+)?)\s*[gG]',
            r'(\d+(?:\.\d+)?)\s*g\b',
            r'(\d+(?:\.\d+)?)\s*G\b'
        ]
        
        for pattern in weight_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                return f"{matches[0]}g"
        return ""

    def extract_dimensions_from_text(self, text):
        """テキストから寸法情報を抽出（強化版）"""
        dimension_patterns = [
            r'幅\d+.*?高さ\d+.*?奥行き\d+.*?mm',
            r'\d+×\d+×\d+\s*mm',
            r'サイズ[：:\s]*([^。\n]+)',
            r'寸法[：:\s]*([^。\n]+)',
            r'大きさ[：:\s]*([^。\n]+)',
            r'(\d+)\s*[×xX]\s*(\d+)\s*[×xX]\s*(\d+)\s*mm',
            r'W\s*(\d+)\s*[×xX]\s*H\s*(\d+)\s*[×xX]\s*D\s*(\d+)\s*mm'
        ]
        
        for pattern in dimension_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                if isinstance(matches[0], tuple):
                    return f"幅{matches[0][0]}×高さ{matches[0][1]}×奥行き{matches[0][2]}mm"
                else:
                    return matches[0]
        return ""

    def detect_jan_code_in_text(self, text):
        """テキストからJANコードを検出（強化版）"""
        jan_patterns = [
            r'JAN[：:\s]*(\d{13})',
            r'JANコード[：:\s]*(\d{13})',
            r'\b(\d{13})\b',
            r'JAN[：:\s]*(\d{8})',
            r'JANコード[：:\s]*(\d{8})',
            r'\b(\d{8})\b',
            r'バーコード[：:\s]*(\d{8,13})',
            r'商品コード[：:\s]*(\d{8,13})'
        ]
        
        for pattern in jan_patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                if len(match) in [8, 13]:
                    return match
        return None

    def extract_brand_from_text(self, text):
        """テキストからブランド名を抽出"""
        brand_patterns = [
            r'ブランド[：:\s]*([^\n。]+)',
            r'メーカー[：:\s]*([^\n。]+)',
            r'製造元[：:\s]*([^\n。]+)',
            r'(株式会社[^\n。]+)',
            r'([A-Za-z\-]+(?:\s+[A-Za-z\-]+)*)\s*(?:株式会社|Co\.|Ltd\.|Inc\.)',
            r'(E-Value|千吉|藤原産業|金貴)'
        ]
        
        for pattern in brand_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                return matches[0].strip()
        return ""

    def extract_product_name_from_text(self, text):
        """テキストから商品名を抽出"""
        product_patterns = [
            r'商品名[：:\s]*([^\n。]+)',
            r'製品名[：:\s]*([^\n。]+)',
            r'品名[：:\s]*([^\n。]+)',
            r'(マルチドライバーセット|ドライバーセット|工具セット)',
            r'([^\n]*セット[^\n]*)',
            r'([^\n]*ドライバー[^\n]*)'
        ]
        
        for pattern in product_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                return matches[0].strip()
        return ""

    def extract_specification_from_text(self, text):
        """テキストから規格情報を抽出"""
        spec_patterns = [
            r'規格[：:\s]*([^\n。]+)',
            r'型番[：:\s]*([^\n。]+)',
            r'品番[：:\s]*([^\n。]+)',
            r'(ERD-\d+)',
            r'([A-Z]{2,4}-\d+)',
            r'(\d+本セット|\d+個セット|\d+ピース)'
        ]
        
        for pattern in spec_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                return matches[0].strip()
        return ""

    def take_screenshot(self):
        """スクリーンショット撮影（安全性向上）"""
        if not SCREENSHOT_AVAILABLE:
            messagebox.showerror("エラー", "スクリーンショット機能が利用できません。")
            return
            
        try:
            self.withdraw()
            time.sleep(0.5)
            
            screenshot = pyautogui.screenshot()
            self.screenshot = screenshot
            
            self.deiconify()
            messagebox.showinfo("完了", "スクリーンショットを撮影しました。\n範囲選択＆OCR実行ボタンを押してください。")
            
        except Exception as e:
            self.deiconify()
            messagebox.showerror("エラー", f"スクリーンショット撮影エラー:\n{str(e)}")

    def load_image_file(self):
        """画像ファイル読み込み"""
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
                messagebox.showinfo("完了", "画像ファイルを読み込みました。\n範囲選択＆OCR実行ボタンを押してください。")
                
        except Exception as e:
            messagebox.showerror("エラー", f"画像読み込みエラー:\n{str(e)}")

    def select_area_and_ocr(self):
        """範囲選択ウィンドウを開く"""
        if not self.screenshot:
            messagebox.showerror("エラー", "先にスクリーンショットを撮影するか、画像ファイルを読み込んでください。")
            return
        
        self.open_selection_window()

    def open_selection_window(self):
        """範囲選択ウィンドウを開く（安全性向上）"""
        try:
            if not self.screenshot or not hasattr(self.screenshot, 'size'):
                messagebox.showerror("エラー", "スクリーンショットデータが無効です。")
                return
                
            selection_window = tk.Toplevel(self)
            selection_window.title("範囲選択")
            selection_window.attributes('-topmost', True)
            
            screen_width = selection_window.winfo_screenwidth()
            screen_height = selection_window.winfo_screenheight()
            
            img_width, img_height = self.screenshot.size
            self.scale_ratio = min(screen_width * 0.8 / img_width, screen_height * 0.8 / img_height, 1.0)
            
            display_width = int(img_width * self.scale_ratio)
            display_height = int(img_height * self.scale_ratio)
            
            resized_screenshot = self.screenshot.resize((display_width, display_height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(resized_screenshot)
            
            canvas = tk.Canvas(selection_window, width=display_width, height=display_height)
            canvas.pack()
            canvas.create_image(0, 0, anchor=tk.NW, image=photo)
            setattr(canvas, 'image', photo)
            
            self.selection_start = None
            self.selection_end = None
            self.selection_rect = None
            
            canvas.bind("<Button-1>", lambda e: self.on_selection_start(e, canvas))
            canvas.bind("<B1-Motion>", lambda e: self.on_selection_drag(e, canvas))
            canvas.bind("<ButtonRelease-1>", lambda e: self.on_selection_end(e, canvas, selection_window))
            
        except Exception as e:
            messagebox.showerror("エラー", f"選択ウィンドウエラー:\n{str(e)}")

    def on_selection_start(self, event, canvas):
        self.selection_start = (event.x, event.y)

    def on_selection_drag(self, event, canvas):
        if self.selection_start:
            if hasattr(self, 'selection_rect') and self.selection_rect:
                canvas.delete(self.selection_rect)
            
            self.selection_rect = canvas.create_rectangle(
                self.selection_start[0], self.selection_start[1],
                event.x, event.y,
                outline="red", width=2
            )

    def on_selection_end(self, event, canvas, window):
        self.selection_end = (event.x, event.y)
        window.destroy()
        self.perform_ocr()

    def perform_ocr(self):
        """選択範囲のOCR処理を実行（無料版最適化）"""
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
                formatted_result = self.format_ocr_result_enhanced(text, jan_code)
                self.result_text.insert("1.0", formatted_result)
                pyperclip.copy(formatted_result)
                messagebox.showinfo("完了", f"JANコード {jan_code} を検出しました。商品情報を抽出しました。")
            else:
                formatted_result = self.format_ocr_result_enhanced(text)
                self.result_text.insert("1.0", formatted_result)
                pyperclip.copy(formatted_result)
                messagebox.showinfo("完了", "OCR処理が完了しました。JANコードは検出されませんでした。")
                
        except Exception as e:
            messagebox.showerror("エラー", f"OCR処理エラー:\n{str(e)}")

    def format_ocr_result_enhanced(self, text, jan_code=None):
        """OCR結果を指定フォーマットで整形（強化版）"""
        formatted_lines = []
        
        if jan_code:
            formatted_lines.append(f"JANコード\t{jan_code}")
        else:
            formatted_lines.append("JANコード\t")
        
        brand = self.extract_brand_from_text(text)
        formatted_lines.append(f"ブランド名\t{brand}")
        
        product_name = self.extract_product_name_from_text(text)
        formatted_lines.append(f"商品名\t{product_name}")
        
        specification = self.extract_specification_from_text(text)
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
    print("=== CustomTkinter OCR機能テスト - JANコード検出（無料版） ===")
    print("起動中...")
    
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    
    app = CustomTkinterOCRFreeApp()
    app.mainloop()
