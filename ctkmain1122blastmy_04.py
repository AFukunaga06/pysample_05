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
from tkinter import ttk

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

# 工事中メッセージを表示する関数
def show_under_construction():
    messagebox.showinfo("お知らせ", "この機能は現在工事中です。\nしばらくお待ちください。")

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

# クリップボードのデータを処理し、JANコードとともにファイルに書き込む関数
@safe_file_operation
def process_clipboard_data(jancode, clipboard_data):
    lines = clipboard_data.split('\n')
    output_data = "\n".join(line.strip() for line in lines if line.strip())
    if jancode:
        output_data = f"JANコード\t{jancode}\n{output_data}"
    with open('input.txt', 'a', encoding='utf-8') as file:
        file.write(output_data + '\n\n')
    print('新しいデータがinput.txtに追加されました。')

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

# クリップボードの内容を貼り付けて処理を実行する関数
def paste_and_execute():
    clipboard_data = pyperclip.paste()
    if not clipboard_data:
        messagebox.showwarning("警告", "クリップボードにデータがありません。")
        return
    process_clipboard_data("", clipboard_data)
    window.output.insert(customtkinter.END, clipboard_data + "\n")

# ファイルを開く関数
@safe_file_operation
def open_file(file_path):
    os.startfile(file_path)

# バッチファイルを実行し、出力ファイルを開く関数
def execute_batch_and_open_file(batch_file, output_file):
    @safe_file_operation
    def execute():
        subprocess.run([batch_file])
        time.sleep(1)
        open_file(output_file)
        check_data_differences(window.output)
    execute()
    
# データ抜けのチェックを行い、checkd01.txtが空の場合にメッセージを表示する関数
def check_empty_checkd01():
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as file:
            content = file.read().strip()
            if not content:
                messagebox.showinfo("情報", "checkd01.txtは空です")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

# バッチファイルを実行し、ファイルを開かずにデータをチェックする関数
def execute_batch_and_skip_opening_file(batch_file):
    @safe_file_operation
    def execute():
        subprocess.run([batch_file])
        time.sleep(1)
        check_empty_checkd01()
        check_data_differences(window.output)
    execute()

# 複数のファイルの内容をクリアする関数
def clear_files():
    if messagebox.askyesno("確認", "本当にクリアして良いですか？"):
        try:
            for file in ["checkd01.txt", "checkd02.txt", "output02.txt", "output.txt", "input.txt"]:
                with open(file, 'w', encoding='utf-8') as f:
                    f.write('')
            messagebox.showinfo("クリア完了", "ファイルのデータをクリアしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルをクリアする際にエラーが発生しました： {str(e)}")

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

# メインウィンドウのクラス
class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        
        # ウィンドウサイズを設定
        window_width = 800
        window_height = 600
        
        # 画面中央に配置
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        
        # ジオメトリを設定
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        # リサイズを許可（横方向のみ）
        self.resizable(True, False)
        
        # 最小ウィンドウサイズを設定
        self.minsize(600, window_height)
        
        # 左右のフレームを作成
        self.left_frame = customtkinter.CTkFrame(self)
        self.left_frame.pack(side='left', fill='both', expand=True, padx=(20,10), pady=20)
        
        # Sashを追加
        self.sash = ttk.Separator(self, orient='vertical')
        self.sash.pack(side='left', fill='y', padx=2)
        
        self.right_frame = customtkinter.CTkFrame(self)
        self.right_frame.pack(side='left', fill='both', expand=True, padx=(10,20), pady=20)
        
        self.setup_ui()

    def setup_ui(self):
        # 左フレームのUI
        self.output = customtkinter.CTkTextbox(self.left_frame, width=360, height=200)
        self.output.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 左側のボタン設定
        buttons_left = [
            ("重複と項目抜けのチェック", lambda: check_and_count_jan_codes(self.output), "#008000"),
            ("JANコード等のコピー", self.jancode_copy, "#FF0000"),
            ("座標軸とコピー", show_under_construction, "#FF0000"),
            ("チェックシートを開く", lambda: open_file(r'C:\pysample_01\checksheet12.bat'), "#0078D4"),
            ("input.txtのチェック", self.check_input_file, "#0078D4"),
            ("クリップボードのクリア", clear_files, "#0078D4"),
            ("Type2x.bat実行とoutput.txt表示", lambda: execute_batch_and_open_file("Type2x.bat", "output.txt"), "#0078D4"),
            ("藤原産業を開く", lambda: open_file('fujiwarasanngyou.bat'), "#0078D4"),
        ]
        
        # 右側のボタン設定
        buttons_right = [
            ("input.txtを開く", lambda: open_file('input.txt'), "#FF0000"),
            ("テキストに貼り付けと実行", paste_and_execute, "#008000"),
            ("データ抜けのチェック", lambda: execute_batch_and_skip_opening_file("Type1x.bat"), "#FF0000"),
            ("商品情報入力シートを開く", lambda: open_file('syouhin_n.bat'), "#0078D4"),
            ("サブフォーム廃番処理", self.open_subform, "#0078D4"),
            ("checkd01.txtを開く", lambda: open_file('checkd01.txt'), "#0078D4"),
            ("checkd02.txtを開く", lambda: open_file('checkd02.txt'), "#0078D4"),
        ]
        
        # 左側のボタンを配置
        for text, command, color in buttons_left:
            btn = customtkinter.CTkButton(self.left_frame, text=text, command=command, fg_color=color)
            btn.pack(fill='x', padx=5, pady=2)
        
        # 右側のボタンを配置
        for text, command, color in buttons_right:
            btn = customtkinter.CTkButton(self.right_frame, text=text, command=command, fg_color=color)
            btn.pack(fill='x', padx=5, pady=2)

    # JANコード等のコピーを実行する関数
    @safe_file_operation
    def jancode_copy(self):
        # サブプロセスの起動位置を指定
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 220
        window_height = 500
        x_position = (screen_width - window_width) // 2 - 400 - 8
        y_position = (screen_height - window_height) // 2
        
        # 環境変数に位置情報を追加
        env = os.environ.copy()
        env['WINDOW_POSITION'] = f"{x_position},{y_position}"
        env['PARENT_WINDOW'] = f"{self.winfo_x()},{self.winfo_y()}"
        
        # サブプロセスを起動
        subprocess.Popen(['python', 'jancoordcopy1122_01.py'], env=env)

    @safe_file_operation
    def open_subform(self):
        subprocess.Popen(['python', 'ckt0412sab01.py'])

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