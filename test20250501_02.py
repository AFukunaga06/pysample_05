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

# スプレッドシートIDと認証ファイルパス
spreadsheet_id = '17Le1KA9nzMREt0Qp9_elM1OF1q9aSp-GDBZRPOntNI8'
KEY_FILE_PATH = r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json'

# Google Sheets 認証
creds = Credentials.from_service_account_file(
    KEY_FILE_PATH,
    scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
)
client = gspread.authorize(creds)
# スプレッドシートを開く（キーが正しくない場合はエラー表示）
try:
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.sheet1
except gspread.exceptions.SpreadsheetNotFound:
    messagebox.showerror("エラー", f"スプレッドシート（キー：{spreadsheet_id}）が見つかりません。キーを確認してください。" )
    worksheet = None
worksheet = client.open_by_key(spreadsheet_id).sheet1

# スクリプト名取得
file_name = os.path.basename(sys.argv[0])

# 必要ファイルの初期化
for fname in ['input.txt', 'checkd01.txt', 'checkd02.txt']:
    if not os.path.exists(fname):
        with open(fname, 'w', encoding='utf-8'):
            pass

# 安全なファイル操作デコレータ
def safe_file_operation(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{str(e)} ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")
    return wrapper

# JANコード重複・廃番チェック
def check_and_count_jan_codes(output):
    with open('input.txt', 'r', encoding='utf-8') as f:
        data_str = f.read()
    jan_codes = re.findall(r'JANコード\t(\d+)', data_str)
    discontinued_codes = [
        code for code in jan_codes
        if 'ブランド名\t廃番' in data_str.split(f'JANコード\t{code}')[1].split('JANコード')[0]
    ]
    duplicate_jan_codes = [code for code, cnt in Counter(jan_codes).items() if cnt > 1]

    output.insert(customtkinter.END, "--- 重複と項目抜け（廃番）チェック 結果 ---\n")
    if duplicate_jan_codes:
        for d in duplicate_jan_codes:
            output.insert(customtkinter.END, f"JANコード {d} が重複しています\n")
    else:
        output.insert(customtkinter.END, "重複はありません\n")
    for c in discontinued_codes:
        output.insert(customtkinter.END, f"JANコード {c} は廃番です\n")
    output.insert(customtkinter.END, f"JANコードは上から{len(jan_codes)}番目です\n")
    if jan_codes:
        output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")
    output.insert(customtkinter.END, "\n")

# データ抜けチェック
def check_empty_checkd01(output):
    with open('checkd01.txt', 'r', encoding='utf-8') as f:
        content = f.read().strip()
    output.insert(customtkinter.END, "--- データ抜けチェック（checkd01.txt） ---\n")
    output.insert(customtkinter.END, "checkd01.txt は空です\n" if not content else "checkd01.txt にデータがあります\n")
    output.insert(customtkinter.END, "\n")

# 差分チェック
def check_data_differences(output):
    with open('checkd01.txt', 'r', encoding='utf-8') as f1, open('checkd02.txt', 'r', encoding='utf-8') as f2:
        l1, l2 = f1.readlines(), f2.readlines()
    diffs = [f"{i+1}番目が違います" for i, (a, b) in enumerate(zip(l1, l2))
             if re.search(r'\d+', a) and re.search(r'\d+', b) and re.search(r'\d+', a).group() != re.search(r'\d+', b).group()]

    output.insert(customtkinter.END, "--- データ差分チェック ---\n")
    output.insert(customtkinter.END, "\n".join(diffs) + "\n" if diffs else "すべてのJANコードが一致しています\n")
    output.insert(customtkinter.END, "\n")

# バッチ実行後まとめてチェック実行
def execute_batch_and_run_checks(batch_file, output):
    @safe_file_operation
    def run_all():
        subprocess.run([batch_file])
        time.sleep(1)
        check_empty_checkd01(output)
        check_data_differences(output)
    run_all()

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Application - {file_name}")
        w, h = 800, 600
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.resizable(False, False)
        self.setup_ui()

    def setup_ui(self):
        frm = customtkinter.CTkFrame(self)
        frm.pack(fill="both", expand=True, padx=20, pady=20)

        self.output = customtkinter.CTkTextbox(frm, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, pady=5, sticky="ew")

        # 統合チェック
        def combined_check():
            self.output.delete("1.0", customtkinter.END)
            check_and_count_jan_codes(self.output)
            execute_batch_and_run_checks("Type1x.bat", self.output)

        buttons = [
            ("重複と項目抜けのチェック", combined_check, 1, 0),
            ("JANコード等のコピー", self.jancode_copy, 1, 1),
            ("データ抜けのチェック", lambda: execute_batch_and_run_checks("Type1x.bat", self.output), 2, 2),
            ("サブフォーム廃番処理", self.open_subform, 4, 1),
            # 必要に応じて他のボタンを追加
        ]
        for txt, cmd, r, c in buttons:
            customtkinter.CTkButton(frm, text=txt, command=cmd).grid(row=r, column=c, padx=5, pady=2, sticky="ew")

        for i in range(4): frm.columnconfigure(i, weight=1)

    @safe_file_operation
    def jancode_copy(self):
        # ここに元のJANコードコピー処理をそのまま実装
        env = os.environ.copy()
        subprocess.Popen(['python', 'jancoordcopy1122_01.py'], env=env)

    @safe_file_operation
    def open_subform(self):
        subprocess.Popen(['python', 'ckt0412sab01.py'])

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
