#!/usr/bin/env python3
"""
マルチディスプレイ対応強化版JAN Code OCR Program
より堅牢なマルチモニター座標処理
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

class MultiDisplayJANCodeOCR:
    def __init__(self):
        self.running = False
        self.click_listener = None
        self.setup_monitors()
        
        # 商品パッケージに特化したTesseract設定
        self.tesseract_configs = [
            '--oem 3 --psm 3 -l jpn+eng',
            '--oem 3 --psm 6 -l jpn+eng',
            '--oem 3 --psm 12 -l jpn+eng',
            '--oem 3 --psm 11 -l jpn+eng',
            '--oem 3 --psm 8 -l jpn+eng',
            '--oem 3 --psm 7 -l jpn+eng',
        ]
    
    def setup_monitors(self):
        """モニター設定の初期化と検証"""
        try:
            self.monitors = get_monitors()
            print(f"=== マルチディスプレイ設定 ===")
            print(f"検出されたモニター数: {len(self.monitors)}")
            
            # 座標軸の範囲を計算
            min_x = min(m.x for m in self.monitors)
            max_x = max(m.x + m.width for m in self.monitors)
            min_y = min(m.y for m in self.monitors)
            max_y = max(m.y + m.height for m in self.monitors)
            
            print(f"\n=== 仮想デスクトップ座標軸 ===")
            print(f"X軸範囲: {min_x} ～ {max_x} (幅: {max_x - min_x})")
            print(f"Y軸範囲: {min_y} ～ {max_y} (高さ: {max_y - min_y})")
            
            # 各モニターの詳細情報を表示
            for i, monitor in enumerate(self.monitors):
                print(f"\nモニター {i+1}:")
                print(f"  解像度: {monitor.width}x{monitor.height}")
                print(f"  左上座標: ({monitor.x}, {monitor.y})")
                print(f"  右下座標: ({monitor.x + monitor.width}, {monitor.y + monitor.height})")
                print(f"  X座標範囲: [{monitor.x} ～ {monitor.x + monitor.width})")
                print(f"  Y座標範囲: [{monitor.y} ～ {monitor.y + monitor.height})")
                
                # モニターの位置関係を表示
                if monitor.x < 0:
                    print(f"  ※ このモニターはプライマリモニターの左側にあります")
                elif monitor.x > 0:
                    print(f"  ※ このモニターはプライマリモニターの右側にあります")
                else:
                    print(f"  ※ このモニターはプライマリモニターと同じX座標です")
                
                if monitor.y < 0:
                    print(f"  ※ このモニターはプライマリモニターの上側にあります")
                elif monitor.y > 0:
                    print(f"  ※ このモニターはプライマリモニターの下側にあります")
                else:
                    print(f"  ※ このモニターはプライマリモニターと同じY座標です")
            
            # プライマリモニターの特定
            self.primary_monitor = self.monitors[0]
            for monitor in self.monitors:
                if monitor.x == 0 and monitor.y == 0:
                    self.primary_monitor = monitor
                    break
            
            print(f"\nプライマリモニター: {self.primary_monitor.width}x{self.primary_monitor.height} at ({self.primary_monitor.x}, {self.primary_monitor.y})")
            
            # 座標軸の視覚的表示
            self.display_coordinate_map()
            
            # モニター配置の検証
            self.validate_monitor_layout()
            
        except Exception as e:
            print(f"モニター検出エラー (ヘッドレス環境): {e}")
            from collections import namedtuple
            Monitor = namedtuple('Monitor', ['x', 'y', 'width', 'height'])
            self.monitors = [Monitor(0, 0, 1920, 1080)]
            self.primary_monitor = self.monitors[0]
            print("デフォルトモニター設定を使用")
    
    def display_coordinate_map(self):
        """座標軸の視覚的マップを表示"""
        print(f"\n=== モニター配置マップ ===")
        
        # 各モニターの相対位置を簡易表示
        min_x = min(m.x for m in self.monitors)
        min_y = min(m.y for m in self.monitors)
        
        # 正規化された座標でマップ作成
        monitor_map = {}
        for i, monitor in enumerate(self.monitors):
            norm_x = (monitor.x - min_x) // 100  # 100ピクセル単位で正規化
            norm_y = (monitor.y - min_y) // 100
            monitor_map[(norm_x, norm_y)] = f"M{i+1}"
        
        # 簡易マップ表示
        max_norm_x = max(x for x, y in monitor_map.keys())
        max_norm_y = max(y for x, y in monitor_map.keys())
        
        print("簡易配置図 (M1=モニター1, M2=モニター2...):")
        for y in range(max_norm_y + 1):
            line = ""
            for x in range(max_norm_x + 1):
                if (x, y) in monitor_map:
                    line += f"[{monitor_map[(x, y)]}]"
                else:
                    line += "    "
            print(line)
        
        print()

    def test_clipboard_functionality(self):
        """クリップボード機能のテスト"""
        print("\n=== クリップボード機能テスト ===")
        test_text = "クリップボードテスト"
        
        try:
            pyperclip.copy(test_text)
            copied_text = pyperclip.paste()
            if copied_text == test_text:
                print("✓ クリップボード機能は正常です")
                return True
            else:
                print(f"⚠ クリップボードの内容が一致しません: '{copied_text}' != '{test_text}'")
                return False
        except Exception as e:
            print(f"✗ クリップボードエラー: {e}")
            print(f"エラータイプ: {type(e).__name__}")
            return False

    def safe_clipboard_copy(self, text):
        """安全なクリップボードコピー（詳細エラー情報付き）"""
        print(f"\n=== クリップボードコピー試行 ===")
        print(f"コピー対象テキスト長: {len(text)} 文字")
        print(f"テキストプレビュー: {text[:100]}{'...' if len(text) > 100 else ''}")
        
        try:
            # クリップボードにコピー
            pyperclip.copy(text)
            print("✓ pyperclip.copy() 成功")
            
            # 検証のため読み戻し
            try:
                copied_back = pyperclip.paste()
                if copied_back == text:
                    print("✓ クリップボード検証成功: 完全一致")
                    return True
                else:
                    print(f"⚠ クリップボード検証警告: 長さの違い (元:{len(text)} vs コピー:{len(copied_back)})")
                    if len(copied_back) == 0:
                        print("  -> クリップボードが空です")
                    elif len(copied_back) < len(text):
                        print("  -> 一部のテキストが失われました")
                    else:
                        print("  -> 予期しない追加テキストがあります")
                    return False
                    
            except Exception as verify_error:
                print(f"⚠ クリップボード検証エラー: {verify_error}")
                print("  -> コピーは成功したが検証に失敗しました")
                return True  # コピー自体は成功したと判断
                
        except ImportError as import_error:
            print(f"✗ pyperclipモジュールエラー: {import_error}")
            print("  解決方法: pip install pyperclip")
            return False
            
        except Exception as copy_error:
            print(f"✗ クリップボードコピーエラー: {copy_error}")
            print(f"  エラータイプ: {type(copy_error).__name__}")
            
            # 環境固有のエラー情報
            if "DISPLAY" in str(copy_error):
                print("  -> X11 DISPLAYエラー: GUI環境が必要です")
            elif "Wayland" in str(copy_error):
                print("  -> Waylandエラー: wl-clipboardが必要な可能性があります")
            elif "xclip" in str(copy_error) or "xsel" in str(copy_error):
                print("  -> X11クリップボードツールが不足しています")
                print("     Ubuntu/Debian: sudo apt install xclip xsel")
                print("     CentOS/RHEL: sudo yum install xclip xsel")
            
            return False
    
    def find_monitor_for_point(self, x, y):
        """クリック位置のモニター特定（改良版）"""
        print(f"座標 ({x}, {y}) のモニターを特定中...")
        
        # 完全一致検索
        for i, monitor in enumerate(self.monitors):
            if (monitor.x <= x < monitor.x + monitor.width and 
                monitor.y <= y < monitor.y + monitor.height):
                print(f"  -> モニター{i+1}で発見: {monitor.width}x{monitor.height} at ({monitor.x}, {monitor.y})")
                return monitor
        
        # 完全一致しない場合の詳細ログ
        print(f"警告: 座標 ({x}, {y}) がどのモニターにも該当しません")
        for i, monitor in enumerate(self.monitors):
            distance_x = min(abs(x - monitor.x), abs(x - (monitor.x + monitor.width)))
            distance_y = min(abs(y - monitor.y), abs(y - (monitor.y + monitor.height)))
            print(f"  モニター{i+1}からの距離: X={distance_x}, Y={distance_y}")
        
        # 最も近いモニターを選択
        closest_monitor = self.primary_monitor
        min_distance = float('inf')
        
        for monitor in self.monitors:
            center_x = monitor.x + monitor.width // 2
            center_y = monitor.y + monitor.height // 2
            distance = ((x - center_x) ** 2 + (y - center_y) ** 2) ** 0.5
            
            if distance < min_distance:
                min_distance = distance
                closest_monitor = monitor
        
        print(f"  -> 最も近いモニターを使用: {closest_monitor.width}x{closest_monitor.height}")
        return closest_monitor
    
    def capture_screen_area_safe(self, x, y, width, height):
        """安全な画面キャプチャ（境界チェック付き）"""
        try:
            print(f"画面キャプチャ: ({x}, {y}) サイズ {width}x{height}")
            
            # 座標の妥当性チェック
            if width <= 0 or height <= 0:
                print(f"エラー: 無効なサイズ {width}x{height}")
                return None
            
            # キャプチャ実行
            screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))
            
            if screenshot.size[0] != width or screenshot.size[1] != height:
                print(f"警告: 期待サイズ {width}x{height} != 実際サイズ {screenshot.size}")
            
            print(f"キャプチャ成功: {screenshot.size}")
            return screenshot
            
        except Exception as e:
            print(f"画面キャプチャエラー: {e}")
            print(f"  試行座標: ({x}, {y}) - ({x + width}, {y + height})")
            return None
    
    def advanced_preprocess_image(self, image):
        """高度な画像前処理"""
        opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
        
        processed_images = []
        processed_images.append(('original', Image.fromarray(gray)))
        
        # コントラスト強化
        for contrast_level in [1.5, 2.0, 2.5]:
            enhanced = ImageEnhance.Contrast(Image.fromarray(gray)).enhance(contrast_level)
            processed_images.append((f'contrast_{contrast_level}', enhanced))
        
        # 閾値処理
        adaptive_thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                               cv2.THRESH_BINARY, 11, 2)
        processed_images.append(('adaptive_thresh', Image.fromarray(adaptive_thresh)))
        
        _, otsu_thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        processed_images.append(('otsu_thresh', Image.fromarray(otsu_thresh)))
        
        # ノイズ除去
        denoised = cv2.fastNlMeansDenoising(gray)
        processed_images.append(('denoised', Image.fromarray(denoised)))
        
        # スケーリング
        height, width = gray.shape
        for scale in [1.5, 2.0, 3.0]:
            scaled = cv2.resize(gray, (int(width*scale), int(height*scale)), 
                              interpolation=cv2.INTER_CUBIC)
            processed_images.append((f'scaled_{scale}', Image.fromarray(scaled)))
        
        return processed_images
    
    def find_jan_codes_in_image(self, image):
        """JAN code検出"""
        jan_codes = []
        processed_images = self.advanced_preprocess_image(image)
        
        for name, proc_image in processed_images[:3]:
            for config in self.tesseract_configs[:3]:
                try:
                    ocr_data = pytesseract.image_to_data(proc_image, config=config, 
                                                       output_type=pytesseract.Output.DICT)
                    
                    for i in range(len(ocr_data['text'])):
                        text = ocr_data['text'][i].strip()
                        if re.match(r'^\d{13}$', text):
                            x = ocr_data['left'][i]
                            y = ocr_data['top'][i]
                            w = ocr_data['width'][i]
                            h = ocr_data['height'][i]
                            jan_codes.append({
                                'text': text,
                                'bbox': (x, y, w, h),
                                'left_click_area': (x - 100, y, 100, h)
                            })
                            print(f"JAN code検出: {text} ({name}, {config})")
                            return jan_codes
                            
                except Exception as e:
                    continue
        
        return jan_codes
    
    def extract_comprehensive_patterns(self, text):
        """包括的パターン抽出"""
        extracted_info = []
        
        patterns = {
            'ブランド名': [r'ブランド名[：:\s]*([^\n\s]+)', r'メーカー[：:\s]*([^\n\s]+)'],
            '商品名': [r'商品名[：:\s]*([^\n]+)', r'製品名[：:\s]*([^\n]+)'],
            '規格': [r'規格[：:\s]*([^\n]+)', r'型番[：:\s]*([^\n]+)'],
            '重量': [r'重量[：:\s]*(\d+)[gｇ]', r'重さ[：:\s]*(\d+)[gｇ]'],
            '商品サイズ': [r'商品サイズ[：:\s]*([^\n]+)', r'サイズ[：:\s]*([^\n]+)'],
        }
        
        for category, pattern_list in patterns.items():
            for pattern in pattern_list:
                matches = re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE)
                for match in matches:
                    if match.groups():
                        value = match.group(1).strip()
                        if value:
                            if category == '重量':
                                weight_num = re.search(r'\d+', value)
                                if weight_num and 10 <= int(weight_num.group()) <= 10000:
                                    extracted_info.append(f"{category}: {value}")
                            else:
                                extracted_info.append(f"{category}: {value}")
        
        return extracted_info
    
    def extract_product_info_comprehensive(self, image, jan_bbox=None):
        """商品情報の包括的抽出"""
        if jan_bbox:
            x, y, w, h = jan_bbox
            search_x = max(0, x - 500)
            search_y = max(0, y - 500)
            search_w = min(image.width - search_x, w + 1000)
            search_h = min(image.height - search_y, h + 1000)
            search_area = image.crop((search_x, search_y, search_x + search_w, search_y + search_h))
        else:
            search_area = image
        
        processed_images = self.advanced_preprocess_image(search_area)
        all_extracted_info = []
        all_ocr_texts = []
        
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
    
    def on_click(self, x, y, button, pressed):
        """マウスクリックイベントハンドラ（改良版）"""
        if not pressed or button != mouse.Button.left:
            return
        
        print(f"\n=== クリック検出 ===")
        print(f"グローバル座標: ({x}, {y})")
        
        # モニター特定
        monitor = self.find_monitor_for_point(x, y)
        
        # 相対座標計算
        relative_x = x - monitor.x
        relative_y = y - monitor.y
        print(f"相対座標: ({relative_x}, {relative_y})")
        print(f"使用モニター: {monitor.width}x{monitor.height} at ({monitor.x}, {monitor.y})")
        
        # 画面キャプチャ
        screenshot = self.capture_screen_area_safe(monitor.x, monitor.y, monitor.width, monitor.height)
        if screenshot is None:
            print("画面キャプチャに失敗しました")
            return
        
        # JAN code検索
        jan_codes = self.find_jan_codes_in_image(screenshot)
        print(f"検出されたJANコード数: {len(jan_codes)}")
        
        # クリック領域判定
        for jan_code in jan_codes:
            left_area = jan_code['left_click_area']
            print(f"JAN {jan_code['text']} クリック領域: ({left_area[0]}, {left_area[1]}) サイズ {left_area[2]}x{left_area[3]}")
            
            if (left_area[0] <= relative_x <= left_area[0] + left_area[2] and
                left_area[1] <= relative_y <= left_area[1] + left_area[3]):
                
                print(f"✓ JAN {jan_code['text']} の左側クリック検出")
                
                # 商品情報抽出
                product_info, debug_text = self.extract_product_info_comprehensive(
                    screenshot, jan_code['bbox'])
                
                if product_info:
                    result_text = f"JANコード: {jan_code['text']}\n"
                    for info in product_info:
                        result_text += f"{info}\n"
                    
                    print("\n=== 抽出された商品情報 ===")
                    print(result_text)
                    
                    # 安全なクリップボードコピー
                    copy_success = self.safe_clipboard_copy(result_text.strip())
                    if copy_success:
                        print("✓ 情報をクリップボードにコピー完了")
                    else:
                        print("✗ クリップボードへのコピーに失敗しました")
                        print("  -> 代替手段: 結果をファイルに保存します")
                        
                        # ファイルとして保存
                        fallback_file = f'/home/ubuntu/ocr_result_{jan_code["text"]}.txt'
                        try:
                            with open(fallback_file, 'w', encoding='utf-8') as f:
                                f.write(result_text)
                            print(f"  -> 結果ファイル保存: {fallback_file}")
                        except Exception as file_error:
                            print(f"  -> ファイル保存も失敗: {file_error}")
                    
                    # デバッグファイル保存
                    timestamp = int(time.time())
                    debug_file = f'/home/ubuntu/multi_display_ocr_{jan_code["text"]}_{timestamp}.txt'
                    try:
                        with open(debug_file, 'w', encoding='utf-8') as f:
                            f.write(f"マルチディスプレイOCR結果\n")
                            f.write(f"実行時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                            f.write(f"JAN Code: {jan_code['text']}\n")
                            f.write(f"グローバル座標: ({x}, {y})\n")
                            f.write(f"相対座標: ({relative_x}, {relative_y})\n")
                            f.write(f"使用モニター: {monitor.width}x{monitor.height} at ({monitor.x}, {monitor.y})\n")
                            f.write(f"クリップボードコピー: {'成功' if copy_success else '失敗'}\n\n")
                            f.write("抽出情報:\n")
                            f.write(result_text)
                            f.write("\n\n=== OCRデバッグ情報 ===\n")
                            f.write(debug_text)
                        print(f"✓ デバッグファイル保存: {debug_file}")
                    except Exception as debug_error:
                        print(f"✗ デバッグファイル保存失敗: {debug_error}")
                else:
                    print("✗ 商品情報が抽出できませんでした")
                
                return
        
        print("JANコードの左側がクリックされませんでした")
    
    def start_listening(self):
        """マウス監視開始"""
        print("\n=== マルチディスプレイ対応JAN Code OCR開始 ===")
        print("JANコードの左側をクリックしてください")
        print("\n使用方法:")
        print("1. JANコード（13桁の数字）を見つける")
        print("2. その数字の左側の領域をマウスで左クリック")
        print("3. 商品情報が自動抽出されクリップボードにコピーされます")
        print("4. Ctrl+Cで停止")
        
        # 現在のマウス位置を表示
        try:
            from pynput.mouse import Listener
            print(f"\n現在のマウス位置をリアルタイム表示中...")
            print("（クリックすると処理開始）")
        except:
            pass
        
        self.running = True
        self.click_listener = mouse.Listener(on_click=self.on_click)
        self.click_listener.start()
        
        try:
            while self.running:
                time.sleep(0.1)
        except KeyboardInterrupt:
            print("\n=== プログラム停止中 ===")
            self.stop_listening()
            print("✓ 停止完了")
    
    def stop_listening(self):
        """マウス監視停止"""
        self.running = False
        if self.click_listener:
            self.click_listener.stop()
    
    def test_with_image(self, image_path):
        """画像テスト機能"""
        try:
            image = Image.open(image_path)
            print(f"テスト画像: {image_path} ({image.size})")
            
            jan_codes = self.find_jan_codes_in_image(image)
            print(f"検出JAN数: {len(jan_codes)}")
            
            for jan_code in jan_codes:
                product_info, debug_text = self.extract_product_info_comprehensive(
                    image, jan_code['bbox'])
                
                if product_info:
                    result_text = f"JANコード: {jan_code['text']}\n"
                    for info in product_info:
                        result_text += f"{info}\n"
                    
                    print("=== テスト結果 ===")
                    print(result_text)
                    return result_text
            
            return None
                
        except Exception as e:
            print(f"テストエラー: {e}")
            import traceback
            traceback.print_exc()
            return None

def main():
    ocr = MultiDisplayJANCodeOCR()
    
    if len(sys.argv) > 1:
        image_path = sys.argv[1]
        result = ocr.test_with_image(image_path)
    else:
        ocr.start_listening()

if __name__ == "__main__":
    main()