import sys
import re
from datetime import datetime

def process_input_file():
    with open('input.txt', 'r', encoding='utf-8') as f:
        input_data = f.read()

    lines = input_data.split('\n')

    output_lines = []
    output_line = []

    jan_code = ""
    brand = ""
    product_name = ""
    spec = ""
    width = ""
    height = ""
    depth = ""
    weight = ""

    seen_jan_codes = set()

    for line in lines:
        if line:
            key_value = line.split('\t')
            key = key_value[0]
            if len(key_value) > 1:
                if key.startswith("重量"):
                    weight_match = re.search(r'\d+(\.\d{1,2})?', key)
                    value = weight_match.group() if weight_match else ""
                else:
                    value = '\t'.join(key_value[1:])
            else:
                if key.startswith("重量"):
                    weight_match = re.search(r'\d+(\.\d{1,2})?', key)
                    value = weight_match.group() if weight_match else ""
                else:
                    value = ""

            if key == "JANコード":
                if jan_code:
                    if brand == "廃番":
                        output_line = [brand, '', jan_code, datetime.now().strftime('%Y/%m/%d')]
                        output_line.extend([''] * 6)  # Add empty fields for the remaining columns
                    else:
                        output_line = ['', '', jan_code, datetime.now().strftime('%Y/%m/%d'), brand, product_name]
                        output_line.extend([spec if spec else "", width, height, depth, weight])
                    output_lines.append('\t'.join(output_line))
                    output_line = []

                if value in seen_jan_codes:
                    output_lines.append(f"\t\t警告: JANコードが重複してます - {value}")
                jan_code = value
                seen_jan_codes.add(value)
                brand = ""
                product_name = ""
                spec = ""
                width = ""
                height = ""
                depth = ""
                weight = ""
            elif key == "ブランド名":
                brand = value
            elif key == "商品名":
                product_name = value
            elif key == "規格":
                spec = value
            elif key == "商品サイズ":
                dimensions = value.split('×')
                width = dimensions[0].replace('幅', '').strip()
                height = dimensions[1].replace('高さ', '').strip()
                depth = dimensions[2].replace('奥行き', '').replace('mm', '').strip()
            elif key.startswith("重量"):
                weight = value.strip()

    if jan_code:
        if brand == "廃番":
            output_line = [brand, '', jan_code, datetime.now().strftime('%Y/%m/%d')]
            output_line.extend([''] * 6)  # Add empty fields for the remaining columns
        else:
            output_line = ['', '', jan_code, datetime.now().strftime('%Y/%m/%d'), brand, product_name]
            output_line.extend([spec if spec else "", width, height, depth, weight])
        output_lines.append('\t'.join(output_line))

    # Modify the lines according to 'outputa.py'
    for i, line in enumerate(output_lines):
        if line.startswith("廃番\t"):
            output_lines[i] = line.replace("廃番\t", "廃番", 1)

    output_data = '\n'.join(output_lines)
    sys.stdout.write(output_data)
    sys.stdout.flush()


if __name__ == "__main__":
    process_input_file()
