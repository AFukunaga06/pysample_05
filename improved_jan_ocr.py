#!/usr/bin/env python3
"""
改良版JAN Code OCR Program
商品パッケージのレイアウトに特化したOCR改善版
"""

import cv2
import numpy as np
import pytesseract
import pyperclip
from PIL import Image, ImageGrab, ImageEnhance, ImageFilter
from pynput import mouse
from screeninfo import get_monitors
import re
import threading
import time
import sys
import os

class ImprovedJANCodeOCR:
    def __init__(self):
        self.running = False
        self.click_listener = None
        
        try:
            self.monitors = get_monitors()
            print(f"Detected {len(self.monitors)} monitors:")
            for i, monitor in enumerate(self.monitors):
                print(f"  Monitor {i+1}: {monitor.width}x{monitor.height} at ({monitor.x}, {monitor.y})")
        except Exception as e:
            print(f"Could not detect monitors (headless environment): {e}")
            from collections import namedtuple
            Monitor = namedtuple('Monitor', ['x', 'y', 'width', 'height'])
            self.monitors = [Monitor(0, 0, 1920, 1080)]
            print("Using default monitor configuration for testing")
        
        # 商品パッケージに特化したTesseract設定
        self.tesseract_configs = [
            '--oem 3 --psm 3 -l jpn+eng',   # 完全に自動的なページ分割（デフォルト）
            '--oem 3 --psm 6 -l jpn+eng',   # 単一の均一なテキストブロック
            '--oem 3 --psm 12 -l jpn+eng',  # スパースなテキスト用
            '--oem 3 --psm 11 -l jpn+eng',  # スパースなテキスト用（PSM 12より厳密）
            '--oem 3 --psm 8 -l jpn+eng',   # 単一語として扱う
            '--oem 3 --psm 7 -l jpn+eng',   # 単一のテキスト行として扱う
        ]
        
    def advanced_preprocess_image(self, image):
        """商品パッケージ用の高度な前処理"""
        opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
        
        processed_images = []
        
        # 1. オリジナル画像
        processed_images.append(('original', Image.fromarray(gray)))
        
        # 2. コントラスト強化（複数レベル）
        for contrast_level in [1.5, 2.0, 2.5]:
            enhanced = ImageEnhance.Contrast(Image.fromarray(gray)).enhance(contrast_level)
            processed_images.append((f'contrast_{contrast_level}', enhanced))
        
        # 3. 適応的閾値処理
        adaptive_thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                               cv2.THRESH_BINARY, 11, 2)
        processed_images.append(('adaptive_thresh', Image.fromarray(adaptive_thresh)))
        
        # 4. OTSU閾値処理
        _, otsu_thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        processed_images.append(('otsu_thresh', Image.fromarray(otsu_thresh)))
        
        # 5. ノイズ除去
        denoised = cv2.fastNlMeansDenoising(gray)
        processed_images.append(('denoised', Image.fromarray(denoised)))
        
        # 6. モルフォロジー演算（テキストの連結改善）
        kernel = np.ones((2,2), np.uint8)
        morphed = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, kernel)
        processed_images.append(('morphed', Image.fromarray(morphed)))
        
        # 7. スケーリング（複数サイズ）
        height, width = gray.shape
        for scale in [1.5, 2.0, 3.0]:
            scaled = cv2.resize(gray, (int(width*scale), int(height*scale)), 
                              interpolation=cv2.INTER_CUBIC)
            processed_images.append((f'scaled_{scale}', Image.fromarray(scaled)))
        
        # 8. シャープネス強化
        sharpened = ImageEnhance.Sharpness(Image.fromarray(gray)).enhance(2.0)
        processed_images.append(('sharpened', sharpened))
        
        return processed_images
        
    def extract_comprehensive_patterns(self, text):
        """包括的なパターン抽出（全角文字対応）"""
        extracted_info = []
        
        # ブランド名パターン
        brand_patterns = [
            r'ブランド名[：:\s]*([^\n\s]+)',
            r'メーカー[：:\s]*([^\n\s]+)',
            r'BRAND[：:\s]*([^\n\s]+)',
        ]
        
        # 商品名パターン
        product_patterns = [
            r'商品名[：:\s]*([^\n]+)',
            r'製品名[：:\s]*([^\n]+)',
            r'品名[：:\s]*([^\n]+)',
        ]
        
        # 規格パターン
        spec_patterns = [
            r'規格[：:\s]*([^\n]+)',
            r'型番[：:\s]*([^\n]+)',
            r'モデル[：:\s]*([^\n]+)',
            r'品番[：:\s]*([^\n]+)',
        ]
        
        # 重量パターン（全角・半角両対応）
        weight_patterns = [
            r'重量[：:\s]*(\d+)[gｇ]',
            r'重さ[：:\s]*(\d+)[gｇ]',
            r'内容量[：:\s]*(\d+)[gｇ]',
            r'(\d{2,4})[gｇ](?![xX×])',  # 数字＋g（寸法でない）
            r'重量.*?(\d+).*?[gｇ]',
            r'NET\s*WT[：:\s]*(\d+)[gｇ]',
        ]
        
        # サイズパターン（全角・半角両対応）
        size_patterns = [
            r'商品サイズ[：:\s]*([^\n]+)',
            r'サイズ[：:\s]*([^\n]+)',
            r'寸法[：:\s]*([^\n]+)',
            r'幅\d+.*?高さ\d+.*?奥行き\d+.*?[mｍ][mｍ]',
            r'\d+[xX×]\d+[xX×]?\d*[mｍ][mｍ]',
        ]
        
        # パターンマッチング実行
        all_patterns = [
            ('ブランド名', brand_patterns),
            ('商品名', product_patterns),
            ('規格', spec_patterns),
            ('重量', weight_patterns),
            ('商品サイズ', size_patterns),
        ]
        
        for category, patterns in all_patterns:
            for pattern in patterns:
                matches = re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE)
                for match in matches:
                    if match.groups():
                        value = match.group(1).strip()
                        if value:  # 空でない場合のみ
                            # 重量の場合は数値チェック
                            if category == '重量':
                                weight_num = re.search(r'\d+', value)
                                if weight_num and 10 <= int(weight_num.group()) <= 10000:
                                    extracted_info.append(f"{category}: {value}")
                            else:
                                extracted_info.append(f"{category}: {value}")
                    else:
                        extracted_info.append(f"{category}: {match.group(0)}")
        
        return extracted_info
    
    def find_jan_codes_in_image(self, image):
        """JAN code検出の改良版"""
        jan_codes = []
        
        # 複数の前処理を試す
        processed_images = self.advanced_preprocess_image(image)
        
        for name, proc_image in processed_images[:3]:  # 最初の3つを試す
            for config in self.tesseract_configs[:3]:  # 最初の3つを試す
                try:
                    ocr_data = pytesseract.image_to_data(proc_image, config=config, 
                                                       output_type=pytesseract.Output.DICT)
                    
                    for i in range(len(ocr_data['text'])):
                        text = ocr_data['text'][i].strip()
                        # JAN codeパターン（13桁）
                        if re.match(r'^\d{13}$', text):
                            x = ocr_data['left'][i]
                            y = ocr_data['top'][i]
                            w = ocr_data['width'][i]
                            h = ocr_data['height'][i]
                            jan_codes.append({
                                'text': text,
                                'bbox': (x, y, w, h),
                                'left_click_area': (x - 100, y, 100, h)  # より広い範囲
                            })
                            print(f"JAN code found: {text} using {name} with {config}")
                            return jan_codes  # 見つかったら即座に返す
                            
                except Exception as e:
                    continue
        
        return jan_codes
    
    def extract_product_info_comprehensive(self, image, jan_bbox=None):
        """商品情報の包括的抽出"""
        # 検索範囲の拡大
        if jan_bbox:
            x, y, w, h = jan_bbox
            search_x = max(0, x - 500)  # より広範囲
            search_y = max(0, y - 500)
            search_w = min(image.width - search_x, w + 1000)
            search_h = min(image.height - search_y, h + 1000)
            search_area = image.crop((search_x, search_y, search_x + search_w, search_y + search_h))
        else:
            search_area = image
        
        processed_images = self.advanced_preprocess_image(search_area)
        
        all_extracted_info = []
        all_ocr_texts = []
        
        # より多くの前処理とOCR設定の組み合わせを試す
        for name, proc_image in processed_images:
            for config in self.tesseract_configs:
                try:
                    text = pytesseract.image_to_string(proc_image, config=config)
                    all_ocr_texts.append(f"=== {name} with {config} ===\n{text}\n")
                    
                    product_info = self.extract_comprehensive_patterns(text)
                    if product_info:
                        all_extracted_info.extend(product_info)
                        
                except Exception as e:
                    continue
        
        # 重複削除
        unique_info = []
        seen = set()
        for info in all_extracted_info:
            if info not in seen:
                unique_info.append(info)
                seen.add(info)
        
        return unique_info, "\n".join(all_ocr_texts)
    
    def capture_screen_area(self, x, y, width, height):
        """画面キャプチャ"""
        try:
            screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))
            return screenshot
        except Exception as e:
            print(f"Error capturing screen area: {e}")
            return None
    
    def find_monitor_for_point(self, x, y):
        """クリック位置のモニター特定"""
        for monitor in self.monitors:
            if (monitor.x <= x < monitor.x + monitor.width and 
                monitor.y <= y < monitor.y + monitor.height):
                return monitor
        return self.monitors[0]
    
    def on_click(self, x, y, button, pressed):
        """マウスクリックイベントハンドラ"""
        if not pressed or button != mouse.Button.left:
            return
        
        print(f"Click detected at ({x}, {y})")
        
        monitor = self.find_monitor_for_point(x, y)
        screenshot = self.capture_screen_area(monitor.x, monitor.y, monitor.width, monitor.height)
        if screenshot is None:
            return
        
        relative_x = x - monitor.x
        relative_y = y - monitor.y
        
        jan_codes = self.find_jan_codes_in_image(screenshot)
        print(f"Found {len(jan_codes)} JAN codes")
        
        for jan_code in jan_codes:
            left_area = jan_code['left_click_area']
            if (left_area[0] <= relative_x <= left_area[0] + left_area[2] and
                left_area[1] <= relative_y <= left_area[1] + left_area[3]):
                
                print(f"Click detected on left side of JAN code: {jan_code['text']}")
                
                product_info, debug_text = self.extract_product_info_comprehensive(
                    screenshot, jan_code['bbox'])
                
                if product_info:
                    # フォーマット化された結果
                    result_text = f"JANコード: {jan_code['text']}\n"
                    for info in product_info:
                        result_text += f"{info}\n"
                    
                    print("抽出された商品情報:")
                    print(result_text)
                    
                    try:
                        pyperclip.copy(result_text.strip())
                        print("情報をクリップボードにコピーしました！")
                    except Exception as e:
                        print(f"クリップボードにコピーできませんでした: {e}")
                    
                    # デバッグファイル保存
                    timestamp = int(time.time())
                    debug_file = f'/home/ubuntu/improved_ocr_{jan_code["text"]}_{timestamp}.txt'
                    with open(debug_file, 'w', encoding='utf-8') as f:
                        f.write(f"JAN Code: {jan_code['text']}\n")
                        f.write(f"Click position: ({x}, {y})\n\n")
                        f.write("抽出された情報:\n")
                        f.write(result_text)
                        f.write("\n\n=== デバッグ情報 ===\n")
                        f.write(debug_text)
                    print(f"デバッグ情報を保存: {debug_file}")
                    
                else:
                    print("商品情報が抽出できませんでした")
                
                break
        else:
            print("JANコードの左側をクリックしてください")
    
    def start_listening(self):
        """マウス監視開始"""
        print("JAN Code OCR監視を開始...")
        print("JANコードの左側をクリックして商品情報を抽出してください。")
        print("Ctrl+Cで停止。")
        
        self.running = True
        self.click_listener = mouse.Listener(on_click=self.on_click)
        self.click_listener.start()
        
        try:
            while self.running:
                time.sleep(0.1)
        except KeyboardInterrupt:
            print("\n停止中...")
            self.stop_listening()
    
    def stop_listening(self):
        """マウス監視停止"""
        self.running = False
        if self.click_listener:
            self.click_listener.stop()
    
    def test_with_image(self, image_path):
        """画像テスト機能"""
        try:
            image = Image.open(image_path)
            print(f"テスト画像: {image_path}")
            print(f"画像サイズ: {image.size}")
            
            jan_codes = self.find_jan_codes_in_image(image)
            print(f"発見したJANコード数: {len(jan_codes)}")
            
            for jan_code in jan_codes:
                print(f"JAN Code: {jan_code['text']}")
                
                product_info, debug_text = self.extract_product_info_comprehensive(
                    image, jan_code['bbox'])
                
                if product_info:
                    result_text = f"JANコード: {jan_code['text']}\n"
                    for info in product_info:
                        result_text += f"{info}\n"
                    
                    print("=== 抽出結果 ===")
                    print(result_text)
                    
                    # デバッグファイル保存
                    debug_file = f'/home/ubuntu/test_debug_{jan_code["text"]}.txt'
                    with open(debug_file, 'w', encoding='utf-8') as f:
                        f.write(result_text)
                        f.write("\n\n=== OCRデバッグ ===\n")
                        f.write(debug_text)
                    print(f"デバッグファイル: {debug_file}")
                    
                    return result_text
                else:
                    print("商品情報が抽出できませんでした")
            
            return None
                
        except Exception as e:
            print(f"エラー: {e}")
            import traceback
            traceback.print_exc()
            return None

def main():
    ocr = ImprovedJANCodeOCR()
    
    if len(sys.argv) > 1:
        image_path = sys.argv[1]
        result = ocr.test_with_image(image_path)
    else:
        ocr.start_listening()

if __name__ == "__main__":
    main()