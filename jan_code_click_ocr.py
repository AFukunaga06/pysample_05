#!/usr/bin/env python3
"""
JAN Code Click OCR Program
Detects clicks on the left side of JAN codes and uses enhanced OCR to extract weight information.
Supports multi-display environments.
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

class JANCodeClickOCR:
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
            self.monitors = [Monitor(0, 0, 1920, 1080)]  # Default monitor
            print("Using default monitor configuration for testing")
        
        self.tesseract_configs = [
            '--oem 3 --psm 6 -l jpn+eng',
            '--oem 3 --psm 8 -l jpn+eng', 
            '--oem 3 --psm 7 -l jpn+eng',
            '--oem 3 --psm 11 -l jpn+eng',
            '--oem 3 --psm 13 -l jpn+eng',
        ]
        
    def preprocess_image(self, image):
        """Apply various preprocessing techniques to improve OCR"""
        opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
        
        processed_images = []
        
        processed_images.append(('original', Image.fromarray(gray)))
        
        enhanced = ImageEnhance.Contrast(Image.fromarray(gray)).enhance(2.0)
        processed_images.append(('contrast', enhanced))
        
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        processed_images.append(('threshold', Image.fromarray(thresh)))
        
        denoised = cv2.fastNlMeansDenoising(gray)
        processed_images.append(('denoised', Image.fromarray(denoised)))
        
        height, width = gray.shape
        scaled = cv2.resize(gray, (width*2, height*2), interpolation=cv2.INTER_CUBIC)
        processed_images.append(('scaled', Image.fromarray(scaled)))
        
        return processed_images
        
    def find_jan_codes_in_image(self, image):
        """Find JAN code patterns in the image using OCR"""
        jan_codes = []
        
        for config in self.tesseract_configs:
            try:
                ocr_data = pytesseract.image_to_data(image, config=config, output_type=pytesseract.Output.DICT)
                
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
                            'left_click_area': (x - 50, y, 50, h)  # Area to the left of JAN code
                        })
                        break
                
                if jan_codes:
                    break
            except:
                continue
        
        return jan_codes
    
    def extract_patterns_from_text(self, text):
        """Extract weight and size patterns from OCR text"""
        extracted_info = []
        
        weight_patterns = [
            r'重量\s*(\d+)\s*g',      # 重量380g
            r'重量(\d+)g',            # 重量380g (no space)
            r'重量\s*(\d+)',          # 重量380 (without g)
            r'重さ\s*(\d+)\s*g',      # 重さ380g
            r'重さ\s*(\d+)',          # 重さ380 (without g)
            r'内容量\s*(\d+)\s*g',    # 内容量380g
            r'(\d+)\s*g(?!\s*[xX×])', # 380g (not followed by x for dimensions)
            r'重量.*?(\d+).*?g',      # 重量...380...g (flexible)
            r'(\d{2,4})\s*g',         # 2-4 digit number followed by g
        ]
        
        for pattern in weight_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                if match.group(1) if len(match.groups()) > 0 else match.group(0):
                    weight_value = match.group(1) if len(match.groups()) > 0 else match.group(0)
                    if weight_value.isdigit() and 10 <= int(weight_value) <= 10000:  # Reasonable weight range
                        extracted_info.append(f"重量{weight_value}g")
                else:
                    extracted_info.append(match.group(0))
        
        size_patterns = [
            r'\d+[xX×]\d+[xX×]?\d*[mM][mM]',  # 125X1.2X22MM
            r'サイズ[：:]\s*[^\n]+',
            r'商品サイズ[：:]\s*[^\n]+',
            r'商品サイズ\s+[^\n]+',           # 商品サイズ with space
            r'規格\s+[^\n]+',                # 規格 specification
            r'幅\d+.*?高さ\d+.*?奥行き\d+.*?mm', # 幅130x高さ160x奥行き28mm
            r'柱\d+.*?高さ\d+.*?内行き\d+.*?mm', # 柱130x高さ160x内行き28mm
        ]
        
        for pattern in size_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                extracted_info.append(match.group(0))
        
        return extracted_info
    
    def extract_weight_info_comprehensive(self, image, jan_bbox=None):
        """Extract weight information using multiple preprocessing techniques"""
        if jan_bbox:
            x, y, w, h = jan_bbox
            search_x = max(0, x - 300)
            search_y = max(0, y - 300)
            search_w = min(image.width - search_x, w + 600)
            search_h = min(image.height - search_y, h + 600)
            search_area = image.crop((search_x, search_y, search_x + search_w, search_y + search_h))
        else:
            search_area = image
        
        processed_images = self.preprocess_image(search_area)
        
        all_extracted_info = []
        
        for name, proc_image in processed_images:
            for config in self.tesseract_configs:
                try:
                    text = pytesseract.image_to_string(proc_image, config=config)
                    
                    weight_info = self.extract_patterns_from_text(text)
                    if weight_info:
                        all_extracted_info.extend(weight_info)
                        
                except Exception as e:
                    continue
        
        unique_info = []
        seen = set()
        for info in all_extracted_info:
            if info not in seen:
                unique_info.append(info)
                seen.add(info)
        
        return unique_info
    
    def capture_screen_area(self, x, y, width, height):
        """Capture a specific area of the screen"""
        try:
            screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))
            return screenshot
        except Exception as e:
            print(f"Error capturing screen area: {e}")
            return None
    
    def find_monitor_for_point(self, x, y):
        """Find which monitor contains the given point"""
        for monitor in self.monitors:
            if (monitor.x <= x < monitor.x + monitor.width and 
                monitor.y <= y < monitor.y + monitor.height):
                return monitor
        return self.monitors[0]  # Default to primary monitor
    
    def on_click(self, x, y, button, pressed):
        """Handle mouse click events"""
        if not pressed or button != mouse.Button.left:
            return
        
        print(f"Click detected at ({x}, {y})")
        
        monitor = self.find_monitor_for_point(x, y)
        print(f"Click is on monitor: {monitor.width}x{monitor.height} at ({monitor.x}, {monitor.y})")
        
        screenshot = self.capture_screen_area(monitor.x, monitor.y, monitor.width, monitor.height)
        if screenshot is None:
            print("Failed to capture screenshot")
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
                
                weight_info = self.extract_weight_info_comprehensive(screenshot, jan_code['bbox'])
                
                if weight_info:
                    result_text = f"JAN Code: {jan_code['text']}\n"
                    result_text += "Extracted Information:\n"
                    for info in weight_info:
                        result_text += f"- {info}\n"
                    
                    print("Extracted information:")
                    print(result_text)
                    
                    try:
                        pyperclip.copy(result_text.strip())
                        print("Information copied to clipboard!")
                    except Exception as e:
                        print(f"Could not copy to clipboard: {e}")
                    
                    timestamp = int(time.time())
                    debug_file = f'/home/ubuntu/ocr_result_{jan_code["text"]}_{timestamp}.txt'
                    with open(debug_file, 'w', encoding='utf-8') as f:
                        f.write(f"Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                        f.write(f"JAN Code: {jan_code['text']}\n")
                        f.write(f"Click position: ({x}, {y})\n")
                        f.write(f"Monitor: {monitor.width}x{monitor.height} at ({monitor.x}, {monitor.y})\n\n")
                        f.write("Extracted Information:\n")
                        f.write(result_text)
                    print(f"Debug info saved to {debug_file}")
                    
                else:
                    print("No weight/size information found")
                
                break
        else:
            print("Click was not on the left side of any JAN code")
    
    def start_listening(self):
        """Start listening for mouse clicks"""
        print("Starting JAN Code OCR listener...")
        print("Click on the left side of a JAN code to extract weight information.")
        print("Press Ctrl+C to stop.")
        
        self.running = True
        self.click_listener = mouse.Listener(on_click=self.on_click)
        self.click_listener.start()
        
        try:
            while self.running:
                time.sleep(0.1)
        except KeyboardInterrupt:
            print("\nStopping...")
            self.stop_listening()
    
    def stop_listening(self):
        """Stop listening for mouse clicks"""
        self.running = False
        if self.click_listener:
            self.click_listener.stop()
    
    def test_with_image(self, image_path):
        """Test OCR functionality with a static image"""
        try:
            image = Image.open(image_path)
            print(f"Testing with image: {image_path}")
            print(f"Image size: {image.size}")
            
            jan_codes = self.find_jan_codes_in_image(image)
            print(f"Found {len(jan_codes)} JAN codes:")
            
            for jan_code in jan_codes:
                print(f"  JAN Code: {jan_code['text']} at {jan_code['bbox']}")
                
                weight_info = self.extract_weight_info_comprehensive(image, jan_code['bbox'])
                
                if weight_info:
                    result_text = f"JAN Code: {jan_code['text']}\n"
                    result_text += "Extracted Information:\n"
                    for info in weight_info:
                        result_text += f"- {info}\n"
                    
                    print("Final result:")
                    print(result_text)
                    
                    try:
                        pyperclip.copy(result_text.strip())
                        print("Result copied to clipboard!")
                    except:
                        print("Could not copy to clipboard (no display)")
                    
                    return result_text
                else:
                    print("No weight/size information found")
            
            return None
                
        except Exception as e:
            print(f"Error testing with image: {e}")
            import traceback
            traceback.print_exc()
            return None

def main():
    ocr = JANCodeClickOCR()
    
    if len(sys.argv) > 1:
        image_path = sys.argv[1]
        result = ocr.test_with_image(image_path)
    else:
        ocr.start_listening()

if __name__ == "__main__":
    main()
