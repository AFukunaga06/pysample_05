import customtkinter
import gspread
import os
import subprocess
import tkinter.messagebox as messagebox
import time
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

# メインウィンドウのクラス
class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - JAN Code and Missing Data Check")
        self.geometry("800x600")
        self.setup_ui()

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)
        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        
        # JANコードの重複と廃番とデータ抜けのチェックボタン
        self.check_all_button = customtkinter.CTkButton(self.frame, text="重複と廃番とデータ抜けのチェック", command=self.check_all, fg_color="#0078D4")
        self.check_all_button.grid(row=1, column=0, padx=5, pady=2, sticky="ew")

    def check_all(self):
        self.check_and_count_jan_codes(self.output)
        self.check_empty_checkd01()

    # JANコードの重複と廃番をチェックし、結果を出力する関数
    def check_and_count_jan_codes(self, output):
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

    # データ抜けのチェックを行い、checkd01.txtが空の場合にメッセージを表示する関数
    def check_empty_checkd01(self):
        try:
            with open('checkd01.txt', 'r', encoding='utf-8') as file:
                content = file.read().strip()
                if not content:
                    messagebox.showinfo("情報", "checkd01.txtは空です")
        except FileNotFoundError:
            messagebox.showwarning("ファイルが見つかりません", "checkd01.txt が見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
