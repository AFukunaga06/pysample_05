#!/usr/bin/env python3

import sys
import os
sys.path.insert(0, '.')

from ctkmain_updated0627 import extract_jan_to_weight, format_extracted_text_for_input

def test_complete_integration():
    """Test the complete integration functionality"""
    test_text = 'JANコード 4977292361613 ブランド名 ＳＫ１１ 商品名 電動サンダー用ナイロンタワシ 規格 #120 商品サイズ 幅95×高さ185×奥行き7mm 重量25ｇ'

    print('Testing complete integration...')
    print(f'Input: {test_text}')

    extracted = extract_jan_to_weight(test_text)
    print(f'Extracted: {extracted}')

    if extracted:
        formatted = format_extracted_text_for_input(extracted)
        print('Formatted output:')
        print(repr(formatted))
        print('Visual output:')
        print(formatted)
        
        if os.path.exists('input.txt'):
            os.remove('input.txt')
        
        with open('input.txt', 'a', encoding='utf-8') as file:
            file.write(formatted + '\n\n')
        
        print('✓ Successfully saved to input.txt')
        
        with open('input.txt', 'r', encoding='utf-8') as file:
            content = file.read()
        print('File contents:')
        print(repr(content))
        
        return True
    else:
        print('✗ Extraction failed')
        return False

if __name__ == '__main__':
    success = test_complete_integration()
    print(f"\nIntegration test {'PASSED' if success else 'FAILED'}")
