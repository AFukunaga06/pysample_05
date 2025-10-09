#!/usr/bin/env python3

import re

def format_extracted_text_for_input(text):
    """抽出されたテキストをinput.txt用にフォーマットする"""
    try:
        jan_pattern = r'\b\d{13}\b|\b\d{8}\b'
        jan_match = re.search(jan_pattern, text)
        
        if jan_match:
            jan_code = jan_match.group()
            remaining_text = re.sub(r'JANコード\s*' + re.escape(jan_code), '', text).strip()
            
            parts = remaining_text.split()
            formatted_lines = [f"JANコード\t{jan_code}"]
            
            i = 0
            while i < len(parts) - 1:
                field_name = parts[i]
                field_value = parts[i + 1]
                formatted_lines.append(f"{field_name}\t{field_value}")
                i += 2
            
            if i < len(parts):
                formatted_lines.append(parts[i])
            
            return '\n'.join(formatted_lines)
        else:
            formatted = re.sub(r'\s+', '\t', text.strip())
            return formatted
            
    except Exception as e:
        print(f"テキスト形式変換中にエラー: {str(e)}")
        return text.strip()

def test_final_format():
    test_text = 'JANコード 4977292361613 ブランド名 ＳＫ１１ 商品名 電動サンダー用ナイロンタワシ 規格 #120 商品サイズ 幅95×高さ185×奥行き7mm 重量25ｇ'
    
    formatted = format_extracted_text_for_input(test_text)
    expected = "JANコード\t4977292361613\nブランド名\tＳＫ１１\n商品名\t電動サンダー用ナイロンタワシ\n規格\t#120\n商品サイズ\t幅95×高さ185×奥行き7mm\n重量25ｇ"
    
    print("Formatted output:")
    print(repr(formatted))
    print("\nExpected output:")
    print(repr(expected))
    print(f"\nMatch: {formatted == expected}")
    
    return formatted == expected

if __name__ == '__main__':
    success = test_final_format()
    print(f"\nTest {'PASSED' if success else 'FAILED'}")
