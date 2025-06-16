def extract_jan_codes(input_file):
    with open(input_file, "r", encoding="utf-8") as infile:
        for line in infile:
            if line.startswith("JANコード"):  # 行が "JANコード" で始まる場合のみ処理
                parts = line.split()
                if len(parts) >= 2:  # 空白区切りの要素が2つ以上ある場合のみ処理
                    jan_code = parts[1].strip()
                    print(jan_code)  # コマンドプロンプトに出力

# 入力ファイルのパスを指定
input_file = "input.txt"

# 関数を実行
extract_jan_codes(input_file)