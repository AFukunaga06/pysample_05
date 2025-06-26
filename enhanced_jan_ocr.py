#!/usr/bin/env python3
"""
Enhanced JAN Code OCR Program with image preprocessing
Improves OCR accuracy for weight information extraction
"""

import cv2
import numpy as np
import pytesseract
import pyperclip
from PIL import Image, ImageEnhance, ImageFilter
import re
import sys
import os

class EnhancedJANCodeOCR:
    def __init__(self):
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
                            'left_click_area': (x - 50, y, 50, h)
                        })
                        break
                
                if jan_codes:
                    break
            except:
                continue
        
        return jan_codes
    
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
        all_ocr_texts = []
        
        for name, proc_image in processed_images:
            for config in self.tesseract_configs:
                try:
                    text = pytesseract.image_to_string(proc_image, config=config)
                    all_ocr_texts.append(f"=== {name} with {config} ===\n{text}\n")
                    
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
        
        return unique_info, "\n".join(all_ocr_texts)
    
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
    
    def test_with_image(self, image_path):
        """Test OCR functionality with comprehensive preprocessing"""
        try:
            image = Image.open(image_path)
            print(f"Testing with image: {image_path}")
            print(f"Image size: {image.size}")
            
            jan_codes = self.find_jan_codes_in_image(image)
            print(f"Found {len(jan_codes)} JAN codes:")
            
            all_results = []
            
            for jan_code in jan_codes:
                print(f"  JAN Code: {jan_code['text']} at {jan_code['bbox']}")
                
                weight_info, full_text = self.extract_weight_info_comprehensive(image, jan_code['bbox'])
                
                print(f"  Extracted information: {weight_info}")
                
                if weight_info:
                    result_text = f"JAN Code: {jan_code['text']}\n"
                    result_text += "Extracted Information:\n"
                    for info in weight_info:
                        result_text += f"- {info}\n"
                    all_results.append(result_text)
                    
                    print("Result for this JAN code:")
                    print(result_text)
                
                debug_filename = f'/home/ubuntu/enhanced_debug_{jan_code["text"]}.txt'
                with open(debug_filename, 'w', encoding='utf-8') as f:
                    f.write(f"JAN Code: {jan_code['text']}\n")
                    f.write(f"Bounding box: {jan_code['bbox']}\n\n")
                    f.write("All OCR attempts:\n")
                    f.write(full_text)
                print(f"Debug info saved to {debug_filename}")
            
            if not jan_codes:
                print("No JAN codes found. Running comprehensive OCR on full image...")
                weight_info, full_text = self.extract_weight_info_comprehensive(image)
                
                if weight_info:
                    result_text = "Extracted Information (no JAN code found):\n"
                    for info in weight_info:
                        result_text += f"- {info}\n"
                    all_results.append(result_text)
                    print("Extracted information:")
                    print(result_text)
                
                with open('/home/ubuntu/enhanced_full_debug.txt', 'w', encoding='utf-8') as f:
                    f.write("Comprehensive OCR results:\n")
                    f.write(full_text)
            
            if all_results:
                final_result = "\n".join(all_results)
                try:
                    pyperclip.copy(final_result)
                    clipboard_msg = "Result copied to clipboard!"
                except:
                    clipboard_msg = "Could not copy to clipboard (no display)"
                
                print("\n=== FINAL RESULT ===")
                print(final_result)
                print(clipboard_msg)
                return final_result
            else:
                print("No weight/size information extracted")
                return None
                
        except Exception as e:
            print(f"Error testing with image: {e}")
            import traceback
            traceback.print_exc()
            return None

def main():
    ocr = EnhancedJANCodeOCR()
    
    if len(sys.argv) > 1:
        image_path = sys.argv[1]
        result = ocr.test_with_image(image_path)
    else:
        print("Usage: python3 enhanced_jan_ocr.py <image_path>")

if __name__ == "__main__":
    main()
