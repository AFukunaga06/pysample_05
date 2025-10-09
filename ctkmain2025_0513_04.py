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

# Google Sheets 認証情報
spreadsheet_id = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
KEY_FILE_PATH = r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json'
creds = Credentials.from_service_account_file(
    KEY_FILE_PATH,
    scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
)
client = gspread.authorize(creds)
worksheet = client.open_by_key(spreadsheet_id).sheet1

file_name = os.path.basename(sys.argv[0])

# input.txt がなければ作成
if not os.path.exists('input.txt'):
    with open('input.txt', 'w', encoding='utf-8'):
        pass

def show_under_construction():
    messagebox.showinfo("お知らせ", "この機能は現在工事中です。\nしばらくお待ちください。")

def safe_file_operation(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{e} ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました：{e}")
    return wrapper

@safe_file_operation
def process_clipboard_data(jancode, clipboard_data):
    lines = [line.strip() for line in clipboard_data.split('\n') if line.strip()]
    output_data = "\n".join(lines)
    if jancode:
        output_data = f"JANコード\t{jancode}\n{output_data}"
    with open('input.txt', 'a', encoding='utf-8') as f:
        f.write(output_data + "\n\n")
    print("新しいデータがinput.txtに追加されました。")

@safe_file_operation
def check_and_count_jan_codes(output: customtkinter.CTkTextbox):
    # ファイル読み込み
    with open('input.txt', 'r', encoding='utf-8') as f:
        data = f.read()
    jan_codes = re.findall(r'JANコード\t(\d+)', data)
    discontinued = [
        code for code in jan_codes
        if 'ブランド名\t廃番' in data.split(f'JANコード\t{code}')[1]
                             .split('JANコード')[0]
    ]
    duplicates = [c for c, cnt in Counter(jan_codes).items() if cnt > 1]

    # 出力クリア
    output.delete("1.0", customtkinter.END)

    # 重複チェック
    if duplicates:
        for d in duplicates:
            output.insert(customtkinter.END, f"JANコード {d} が重複しています\n")
    else:
        output.insert(customtkinter.END, "重複はありません\n")

    # 廃番チェック
    for code in discontinued:
        output.insert(customtkinter.END, f"JANコード {code} は廃番です\n")

    # ─────────────────────────────────────────
    # CTkTextbox の内部テキストウィジェットに対してタグ設定
    text_widget = output.textbox
    bold_font = tkfont.Font(font=text_widget.cget("font"))
    bold_font.configure(weight="bold")
    text_widget.tag_config("red_text", foreground="red", font=bold_font)

    # 赤字太字で挿入
    output.insert(
        customtkinter.END,
        f"JANコードは上から{len(jan_codes)}番目です\n",
        "red_text"
    )
    # ─────────────────────────────────────────

    # 最後のJANコード表示
    if jan_codes:
        output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")

def paste_and_execute():
    clip = pyperclip.paste()
    if not clip:
        messagebox.showwarning("警告", "クリップボードにデータがありません。")
        return
    process_clipboard_data("", clip)
    window.output.insert(customtkinter.END, clip + "\n")

@safe_file_operation
def open_file(path):
    if not os.path.exists(path):
        messagebox.showerror("エラー", f"{path} が存在しません。")
        return
    os.startfile(path)

def execute_batch_and_open_file(batch, out_file):
    @safe_file_operation
    def run():
        subprocess.run([batch])
        time.sleep(1)
        open_file(out_file)
        check_data_differences(window.output)
    run()

def check_empty_checkd01():
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as f:
            if not f.read().strip():
                messagebox.showinfo("情報", "checkd01.txt は空です")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました：{e}")

def execute_batch_and_skip_opening_file(batch):
    @safe_file_operation
    def run():
        subprocess.run([batch])
        time.sleep(1)
        check_empty_checkd01()
        check_data_differences(window.output)
    run()

def clear_files():
    if messagebox.askyesno("確認", "本当にクリアして良いですか？"):
        try:
            for fn in ["checkd01.txt", "checkd02.txt", "output02.txt", "output.txt", "input.txt"]:
                if os.path.exists(fn):
                    with open(fn, 'w', encoding='utf-8'):
                        pass
            messagebox.showinfo("クリア完了", "ファイルのデータをクリアしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"クリア中にエラーが発生しました：{e}")

@safe_file_operation
def check_data_differences(output):
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as f1, \
             open('checkd02.txt', 'r', encoding='utf-8') as f2:
            l1, l2 = f1.readlines(), f2.readlines()
        diffs = []
        for i, (a, b) in enumerate(zip(l1, l2), start=1):
            m1 = re.search(r'\d+', a)
            m2 = re.search(r'\d+', b)
            if m1 and m2 and m1.group() != m2.group():
                diffs.append(f"{i}番目が違います")
        output.delete("1.0", customtkinter.END)
        output.insert(customtkinter.END, "\n".join(diffs) if diffs else "すべてのJANコードが一致しています")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt または checkd02.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました：{e}")

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)

        # ウィンドウサイズと中央配置
        ww, wh = 800, 600
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}")

        self.resizable(False, False)
        self.setup_ui()

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)

        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        buttons = [
            ("重複と項目抜けのチェック", lambda: check_and_count_jan_codes(self.output), 1, 0, "#008000"),
            ("JANコード等のコピー",        self.jancode_copy,                     1, 1, "#FF0000"),
            ("テキストに貼り付けと実行",  paste_and_execute,                    2, 1, "#008000"),
            ("データ抜けのチェック",      lambda: execute_batch_and_skip_opening_file("Type1x.bat"), 2, 2, "#FF0000"),
            ("チェックシートを開く",      lambda: open_file(r'C:\pysample_01\checksheet12.bat'), 3, 0, "#0078D4"),
            ("座標軸とコピー",            show_under_construction,              2, 0, "#FF0000"),
            ("商品情報入力シートを開く",  lambda: open_file('syouhin_n.bat'),      3, 1, "#0078D4"),
            ("藤原産業を開く",            lambda: open_file('fujiwarasanngyou.bat'), 3, 2, "#0078D4"),
            ("input.txtのチェック",       self.check_input_file,                 4, 0, "#0078D4"),
            ("サブフォーム廃番処理",      self.open_subform,                     4, 1, "#0078D4"),
            ("Type2x.bat実行とoutput.txt表示", lambda: execute_batch_and_open_file("Type2x.bat", "output.txt"), 4, 2, "#0078D4"),
            ("checkd01.txtを開く",       lambda: open_file('checkd01.txt'),     5, 1, "#0078D4"),
            ("checkd02.txtを開く",       lambda: open_file('checkd02.txt'),     5, 2, "#0078D4"),
            ("クリップボードのクリア",    clear_files,                          5, 0, "#0078D4"),
            ("input.txtを開く",          lambda: open_file('input.txt'),       1, 2, "#FF0000"),
        ]
        for txt, cmd, r, c, col in buttons:
            btn = customtkinter.CTkButton(self.frame, text=txt, command=cmd, fg_color=col)
            btn.grid(row=r, column=c, padx=5, pady=2, sticky="ew")

        for i in range(4):
            self.frame.columnconfigure(i, weight=1)

        # 下部の注意ラベル
        self.info_label = customtkinter.CTkLabel(
            self.frame,
            text="※商品情報入力シートに必ず名前を明示すること",
            text_color="red",
            font=("メイリオ", 18, "bold")
        )
        self.info_label.grid(row=6, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        self.info_label2 = customtkinter.CTkLabel(
            self.frame,
            text="※チェックシートの内容をcheckd01.txtにコピーする",
            text_color="red",
            font=("メイリオ", 18, "bold")
        )
        self.info_label2.grid(row=7, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

    @safe_file_operation
    def jancode_copy(self):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        ww, wh = 220, 500
        x = (sw - ww) // 2 - 400 - 8
        y = (sh - wh) // 2
        env = os.environ.copy()
        env['WINDOW_POSITION'] = f"{x},{y}"
        env['PARENT_WINDOW'] = f"{self.winfo_x()},{self.winfo_y()}"
        subprocess.Popen(['python', 'jancoordcopy1122_01.py'], env=env)

    @safe_file_operation
    def open_subform(self):
        subprocess.Popen(['python', 'ckt0412sab01.py'])

    def check_input_file(self):
        try:
            with open('input.txt', 'r', encoding='utf-8') as f:
                lines = f.readlines()
            invalid = [i for i, l in enumerate(lines, 1) if 'ANコード' in l and 'J' not in l]
            thirteen = [i for i, l in enumerate(lines, 1) if re.match(r'^\d{13}$', l.strip())]
            msg = (
                f"「ANコード」に「J」が含まれていない行: {', '.join(map(str, invalid)) or 'なし'}\n"
                f"13桁の数値のみを含む行: {', '.join(map(str, thirteen)) or 'なし'}"
            )
            messagebox.showinfo("チェック結果", msg)
        except FileNotFoundError:
            messagebox.showerror("エラー", "input.txt ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"チェック中にエラーが発生しました：{e}")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
