import customtkinter
import gspread
import pyautogui as p
import os
import subprocess
import tkinter.messagebox as messagebox
import time
import pyperclip
from collections import Counter
import sys
import re
from google.oauth2.service_account import Credentials

# スプレッドシートID
spreadsheet_id = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'

# JSONキーファイルへのパス
KEY_FILE_PATH = r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json'

# 認証情報を取得
creds = Credentials.from_service_account_file(KEY_FILE_PATH, scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'])

# gspreadクライアントを作成
client = gspread.authorize(creds)

# スプレッドシートを開く
worksheet = client.open_by_key(spreadsheet_id).sheet1

# スクリプトのファイル名を取得
file_name = os.path.basename(sys.argv[0])

# 安全にファイル操作を行うためのデコレーター関数
def safe_file_operation(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{str(e)}ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")
    return wrapper

# JANコードの重複と廃番をチェックし、結果を出力する関数
@safe_file_operation
def check_and_count_jan_codes(output):
    with open('input.txt', 'r', encoding='utf-8') as f:
        data_str = f.read()
    jan_codes = re.findall(r'JANコード\t(\d+)', data_str)
    discontinued_codes = [code for code in jan_codes if 'ブランド名\t廃番' in data_str.split(f'JANコード\t{code}')[1].split('JANコード')[0]]
    duplicate_jan_codes = [code for code, count in Counter(jan_codes).items() if count > 1]
    output.delete("1.0", customtkinter.END)
    if duplicate_jan_codes:
        for duplicate in duplicate_jan_codes:
            output.insert(customtkinter.END, f"JANコード {duplicate} が重複しています\n")
    else:
        output.insert(customtkinter.END, "重複はありません\n")
    for code in discontinued_codes:
        output.insert(customtkinter.END, f"JANコード {code} は廃番です\n")
    output.insert(customtkinter.END, f"JANコードは上から{len(jan_codes)}番目です\n")
    if jan_codes:
        output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")

    # データ抜けのチェックも同時に実行
    check_data_differences(output)

# データの差分をチェックする関数
@safe_file_operation
def check_data_differences(output):
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as file1, open('checkd02.txt', 'r', encoding='utf-8') as file2:
            lines1 = file1.readlines()
            lines2 = file2.readlines()
            differences = []
            for i, (line1, line2) in enumerate(zip(lines1, lines2)):
                jan_code1 = re.search(r'\d+', line1)
                jan_code2 = re.search(r'\d+', line2)
                if jan_code1 and jan_code2 and jan_code1.group() != jan_code2.group():
                    differences.append(f"{i + 1}番目が違います")
            output.delete("1.0", customtkinter.END)
            if differences:
                output.insert(customtkinter.END, "\n".join(differences))
            else:
                output.insert(customtkinter.END, "すべてのJANコードが一致しています")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt または checkd02.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

# クリップボードの内容を貼り付けて処理を実行する関数
def paste_and_execute():
    clipboard_data = pyperclip.paste()
    if not clipboard_data:
        messagebox.showwarning("警告", "クリップボードにデータがありません。")
        return
    process_clipboard_data("", clipboard_data)
    window.output.insert(customtkinter.END, clipboard_data + "\n")

# バッチファイルを実行し、ファイルを開かずにデータをチェックする関数
def execute_batch_and_skip_opening_file(batch_file):
    @safe_file_operation
    def execute():
        subprocess.run([batch_file])
        time.sleep(1)
        check_empty_checkd01()  # checkd01.txtが空かどうかを確認
        check_data_differences(window.output)  # バッチファイル実行後にデータ抜けチェックを自動実行
    execute()

# メインウィンドウのクラス
class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        self.geometry("800x600")
        self.setup_ui()

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)
        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        buttons = [
            ("重複と項目抜けのチェック", lambda: check_and_count_jan_codes(self.output), 1, 0, "#008000"),  # 濃い緑色
            ("JANコードのコピー", self.jancode_copy, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", paste_and_execute, 2, 1, "#008000"),  # 濃い緑色
            ("データ抜けのチェック", lambda: execute_batch_and_skip_opening_file("Type1x.bat"), 2, 2, "#FF0000"),
            ("チェックシートを開く", lambda: open_file(r'C:\pysample_01\checksheet12.bat'), 3, 0, "#0078D4"),
            ("座標軸とコピー", self.open_coordinate_subform, 2, 0, "#FF0000"),
            ("商品情報入力シートを開く", lambda: open_file('syouhin_n.bat'), 3, 1, "#0078D4"),
            ("藤原産業を開く", lambda: open_file('fujiwarasanngyou.bat'), 3, 2, "#0078D4"),
            ("input.txtのチェック", self.check_input_file, 4, 0, "#0078D4"),
            ("サブフォーム廃番処理", self.open_subform, 4, 1, "#0078D4"),
            ("Type2x.bat実行とoutput.txt表示", lambda: execute_batch_and_open_file("Type2x.bat", "output.txt"), 4, 2, "#0078D4"),
            ("checkd01.txtを開く", lambda: open_file('checkd01.txt'), 5, 1, "#0078D4"),
            ("checkd02.txtを開く", lambda: open_file('checkd02.txt'), 5, 2, "#0078D4"),
            ("クリップボードのクリア", clear_files, 5, 0, "#0078D4"),
            ("input.txtを開く", lambda: open_file('input.txt'), 1, 2, "#FF0000"),
        ]
        for text, command, row, col, color in buttons:
            btn = customtkinter.CTkButton(self.frame, text=text, command=command, fg_color=color)
            btn.grid(row=row, column=col, padx=5, pady=2, sticky="ew")
        for i in range(4):
            self.frame.columnconfigure(i, weight=1)

    # JANコードのコピーを実行する関数
    @safe_file_operation
    def jancode_copy(self):
        subprocess.Popen(['python', 'jancopy0906_01.py'])

    # サブフォームを開く関数
    @safe_file_operation
    def open_subform(self):
        subprocess.Popen(['python', 'ckt0412sab01.py'])

    # 座標サブフォームを開く関数
    @safe_file_operation
    def open_coordinate_subform(self):
        subprocess.Popen(['python', 'Copycoord0902_01.py'])

    # input.txtファイルをチェックする関数
    def check_input_file(self):
        try:
            with open('input.txt', 'r', encoding='utf-8') as f:
                lines = f.readlines()
            invalid_an_code_lines = [i for i, line in enumerate(lines, 1) if 'ANコード' in line and 'J' not in line]
            thirteen_digit_lines = [i for i, line in enumerate(lines, 1) if re.match(r'^\d{13}$', line.strip())]
            message = (
                f"「ANコード」に「J」が含まれていない行: {', '.join(map(str, invalid_an_code_lines)) or 'なし'}\n"
                f"13桁の数値のみを含む行: {', '.join(map(str, thirteen_digit_lines)) or 'なし'}"
            )
            messagebox.showinfo("チェック結果", message)
        except FileNotFoundError:
            messagebox.showerror("エラー", "input.txtファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルのチェック中にエラーが発生しました： {str(e)}")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
