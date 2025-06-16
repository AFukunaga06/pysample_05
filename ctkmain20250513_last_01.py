import customtkinter
import tkinter.font as tkfont
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
    'https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'
])
client = gspread.authorize(creds)
worksheet = client.open_by_key(spreadsheet_id).sheet1
file_name = os.path.basename(sys.argv[0])

if not os.path.exists('input.txt'):
    with open('input.txt', 'w', encoding='utf-8') as f:
        pass

def show_under_construction():
    messagebox.showinfo("お知らせ", "この機能は現在工事中です。\nしばらくお待ちください。")

def safe_file_operation(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{str(e)} ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました：{str(e)}")
    return wrapper

@safe_file_operation
def process_clipboard_data(jancode, clipboard_data):
    lines = clipboard_data.split('\n')
    output_data = "\n".join(line.strip() for line in lines if line.strip())
    if jancode:
        output_data = f"JANコード\t{jancode}\n{output_data}"
    with open('input.txt', 'a', encoding='utf-8') as file:
        file.write(output_data + '\n\n')
    print('新しいデータがinput.txtに追加されました。')

@safe_file_operation
def check_and_count_jan_codes(output: customtkinter.CTkTextbox):
    # ファイル読み込み
    with open('input.txt', 'r', encoding='utf-8') as f:
        data_str = f.read()
    jan_codes = re.findall(r'JANコード\t(\d+)', data_str)
    discontinued_codes = [
        code for code in jan_codes
        if 'ブランド名\t廃番' in data_str.split(f'JANコード\t{code}')[1].split('JANコード')[0]
    ]
    duplicate_jan_codes = [
        code for code, cnt in Counter(jan_codes).items() if cnt > 1
    ]

    # 一旦クリア
    output.delete("1.0", customtkinter.END)

    # 重複チェック結果
    if duplicate_jan_codes:
        for dup in duplicate_jan_codes:
            output.insert(customtkinter.END, f"JANコード {dup} が重複しています\n")
    else:
        output.insert(customtkinter.END, "重複はありません\n")

    # 廃番チェック結果
    for code in discontinued_codes:
        output.insert(customtkinter.END, f"JANコード {code} は廃番です\n")

    # ─────────────────────────────────────────
    # ★ ここから「赤字＋太字」設定
    # CTkTextbox の内部 tkinter.Text ウィジェットを取得
    text_widget = output.textbox

    # 現在のフォントをコピーして太字に
    bold_font = tkfont.Font(font=text_widget.cget("font"))
    bold_font.configure(weight="bold")

    # 内部の Text に対してタグを設定
    text_widget.tag_config("red_text", foreground="red", font=bold_font)

    # 挿入時にタグ名 "red_text" を指定
    output.insert(
        customtkinter.END,
        f"JANコードは上から{len(jan_codes)}番目です\n",
        "red_text"
    )
    # ★ ここまで「赤字＋太字」設定
    # ─────────────────────────────────────────

    # 最後のコード表示
    if jan_codes:
        output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")

def paste_and_execute():
    clipboard_data = pyperclip.paste()
    if not clipboard_data:
        messagebox.showwarning("警告", "クリップボードにデータがありません。")
        return
    process_clipboard_data("", clipboard_data)
    window.output.insert(customtkinter.END, clipboard_data + "\n")

@safe_file_operation
def open_file(file_path):
    if not os.path.exists(file_path):
        messagebox.showerror("エラー", f"{file_path} が存在しません。")
        return
    os.startfile(file_path)

def execute_batch_and_open_file(batch_file, output_file):
    @safe_file_operation
    def execute():
        subprocess.run([batch_file])
        time.sleep(1)
        open_file(output_file)
        check_data_differences(window.output)
    execute()

def check_empty_checkd01():
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as f:
            if not f.read().strip():
                messagebox.showinfo("情報", "checkd01.txt は空です")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました：{str(e)}")

def execute_batch_and_skip_opening_file(batch_file):
    @safe_file_operation
    def execute():
        subprocess.run([batch_file])
        time.sleep(1)
        check_empty_checkd01()
        check_data_differences(window.output)
    execute()

def clear_files():
    if messagebox.askyesno("確認", "本当にクリアして良いですか？"):
        try:
            for fn in ["checkd01.txt", "checkd02.txt", "output02.txt", "output.txt", "input.txt"]:
                if os.path.exists(fn):
                    with open(fn, 'w', encoding='utf-8') as f:
                        f.write('')
            messagebox.showinfo("クリア完了", "ファイルのデータをクリアしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルをクリアする際にエラーが発生しました：{str(e)}")

@safe_file_operation
def check_data_differences(output):
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as f1, \
             open('checkd02.txt', 'r', encoding='utf-8') as f2:
            lines1, lines2 = f1.readlines(), f2.readlines()
        diffs = []
        for i, (l1, l2) in enumerate(zip(lines1, lines2), start=1):
            c1 = re.search(r'\d+', l1)
            c2 = re.search(r'\d+', l2)
            if c1 and c2 and c1.group() != c2.group():
                diffs.append(f"{i}番目が違います")
        output.delete("1.0", customtkinter.END)
        output.insert(customtkinter.END, "\n".join(diffs) if diffs else "すべてのJANコードが一致しています")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt または checkd02.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました：{str(e)}")

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        ww, wh = 800, 600
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{ww}x{wh
