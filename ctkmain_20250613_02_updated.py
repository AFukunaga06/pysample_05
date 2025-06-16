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

spreadsheet_id = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
KEY_FILE_PATH = r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json'

creds = Credentials.from_service_account_file(KEY_FILE_PATH, scopes=[
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
])
client = gspread.authorize(creds)
worksheet = client.open_by_key(spreadsheet_id).sheet1
file_name = os.path.basename(sys.argv[0])

if not os.path.exists('input.txt'):
    with open('input.txt', 'w', encoding='utf-8') as f:
        pass

def show_under_construction():
    messagebox.showinfo("お知らせ", "この機能は現在工事中です。\nしばらくお待ちください。")

# 安全にファイル操作を行うデコレーター
def safe_file_operation(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{str(e)} ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")
    return wrapper

@safe_file_operation
def process_clipboard_data(jancode, clipboard_data):
    lines = clipboard_data.split('\n')
    output_data = "\n".join(line.strip() for line in lines if line.strip())
    if jancode:
        output_data = f"JANコード\t{jancode}\n{output_data}"
    pyperclip.copy(output_data)
    messagebox.showinfo("コピー完了", "JANコードとクリップボードのデータをクリップボードにコピーしました。")

# 他の補助関数定義（open_file, execute_batch_and_open_file, check_empty_checkd01 など）
...

# 重複チェックおよびデータ抜けチェック関数
@safe_file_operation
def check_duplicate_and_data_gaps(output):
    # テキストボックスをクリア
    output.delete("1.0", customtkinter.END)
    # ■ 重複と項目抜けのチェック結果見出し
    output.insert(customtkinter.END, "■ 重複と項目抜けのチェック結果：\n")
    # input.txt を読み込んで JAN コード列を取得
    with open('input.txt', 'r', encoding='utf-8') as f:
        data_str = f.read()
    jan_codes = re.findall(r'JANコード\t(\d+)', data_str)
    discontinued_codes = [
        code for code in jan_codes
        if 'ブランド名\t廃番' in data_str.split(f'JANコード\t{code}')[1].split('JANコード')[0]
    ]
    duplicate_jan_codes = [code for code, count in Counter(jan_codes).items() if count > 1]

    # 重複チェック
    if duplicate_jan_codes:
        for duplicate in duplicate_jan_codes:
            output.insert(customtkinter.END, f"JANコード {duplicate} が重複しています\n")
    else:
        output.insert(customtkinter.END, "重複はありません\n")

    # 廃番チェック
    for code in discontinued_codes:
        output.insert(customtkinter.END, f"JANコード {code} は廃番です\n")

    # コード数と最新コード表示
    output.insert(customtkinter.END, f"JANコードは上から{len(jan_codes)}番目です\n", "red_text")
    if jan_codes:
        output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")
    output.tag_config("red_text", foreground="red")

    # ■ データ抜けのチェック結果見出し
    output.insert(customtkinter.END, "\n■ データ抜けのチェック結果：\n")
    data_gap_result = get_data_differences_result()
    output.insert(customtkinter.END, data_gap_result)

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        window_width = 800
        window_height = 600
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.resizable(False, False)
        self.setup_ui()

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)

        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        buttons = [
            ("重複と項目抜けのチェック", lambda: check_duplicate_and_data_gaps(self.output), 1, 0, "#008000"),
            ("JANコード等のコピー", self.jancode_copy, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", paste_and_execute, 2, 1, "#008000"),
            ("データ抜けのチェック", lambda: execute_batch_and_skip_opening_file("Type1x.bat"), 2, 2, "#FF0000"),
            # 他のボタン定義...
        ]
        for (text, cmd, r, c, color) in buttons:
            btn = customtkinter.CTkButton(self.frame, text=text, command=cmd, fg_color=color)
            btn.grid(row=r, column=c, padx=5, pady=5, sticky="ew")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
