import customtkinter
import gspread
import pyperclip
import os
import subprocess
import tkinter.messagebox as messagebox
import sys
import re
import time
from google.oauth2.service_account import Credentials

spreadsheet_id = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
KEY_FILE_PATH = r'C:\pysample_01\samplep20240906-KEY.json'
file_name = os.path.basename(sys.argv[0])

def safe_file_operation(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{str(e)}ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"予期せぬエラーが発生しました: {str(e)}")
    return wrapper

# input.txtが存在しない場合は作成
if not os.path.exists('input.txt'):
    with open('input.txt', 'w', encoding='utf-8') as f:
        pass

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application")
        window_width = 800
        window_height = 600
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill='both', expand=True, padx=20, pady=20)

        self.output = customtkinter.CTkTextbox(self.frame, height=150)
        self.output.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky='ew')

        # ボタン一覧（修正済み）
        buttons = [
            ("JANコードチェック", lambda: print("JANコードチェック"), 1, 0, "#FF0000"),
            ("JANコード等のコピー", lambda: print("コピー機能実行"), 1, 1, "#FF0000"),
            ("input.txtを開く", lambda: self.open_file('input.txt'), 1, 2, "#FF0000"),
            ("クリップボードのデータ貼り付け", self.paste_clipboard_data, 2, 0, "#0078D4"),
            ("クリップボードのクリア", self.clear_clipboard, 2, 1, "#0078D4"),
            ("Type2x.bat実行とoutput.txt表示", lambda: self.execute_batch_and_open_file("Type2x.bat", "output.txt"), 2, 2, "#0078D4"),
            ("商品情報入力シートを開く", lambda: self.open_file('checksheet12.bat'), 3, 0, "#0078D4"),
            ("藤原産業を開く", lambda: self.open_file('fujiwarasanngyou.bat'), 3, 1, "#0078D4"),
            ("checkd01.txtを開く", lambda: self.open_file('checkd01.txt'), 3, 2, "#0078D4"),
        ]

        for text, command, row, col, color in buttons:
            button = customtkinter.CTkButton(self.frame, text=text, command=command, fg_color=color)
            button.grid(row=row, column=col, padx=5, pady=5, sticky='ew')

        # 説明ラベル追加（クリップボードのクリアの下）
        description_label = customtkinter.CTkLabel(
            self.frame,
            text="※名前記入欄に福永の名前が入っているか確認すること。",
            text_color="red"
        )
        description_label.grid(row=3, column=1, padx=5, pady=(0,10), sticky='ew')

    # 各種メソッド（仮実装）
    def paste_clipboard_data(self):
        data = pyperclip.paste()
        if data:
            self.output.insert(customtkinter.END, data + "\n")
        else:
            messagebox.showwarning("警告", "クリップボードにデータがありません。")

    def clear_clipboard(self):
        pyperclip.copy('')
        messagebox.showinfo("情報", "クリップボードをクリアしました。")

    @safe_file_operation
    def open_file(self, file_path):
        if os.path.exists(file_path):
            os.startfile(file_path)
        else:
            messagebox.showerror("エラー", f"{file_path}が見つかりません。")

    @safe_file_operation
    def execute_batch_and_open_file(self, batch_file, output_file):
        subprocess.run(batch_file, shell=True)
        time.sleep(1)
        self.open_file(output_file)

    def jancode_check(self):
        messagebox.showinfo("情報", "JANコードチェック処理を実装してください。")

    def jancode_copy(self):
        messagebox.showinfo("お知らせ", "JANコード等のコピー機能は準備中です。")

    def clear_clipboard(self):
        pyperclip.copy('')
        messagebox.showinfo("情報", "クリップボードをクリアしました。")

# アプリの起動
if __name__ == "__main__":
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    app = App()
    app.mainloop()
