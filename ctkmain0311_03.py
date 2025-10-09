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

# --- スプレッドシート認証やその他の関数の設定 ---
# ※ここに必要なスプレッドシート認証や他の関数の実装を追加してください。

# グローバル変数の定義（ファイル名など必要な情報を設定）
file_name = "CTKMain0311_02.py"

def open_file(file_path):
    """
    指定したファイルを既定のアプリケーションで開く関数
    """
    try:
        if os.path.exists(file_path):
            os.startfile(file_path)
        else:
            messagebox.showerror("Error", f"ファイルが見つかりません: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def check_and_count_jan_codes(output_widget):
    """
    JANコードの重複と項目抜けをチェックするダミー関数
    """
    # ここに実際のチェックロジックを実装してください
    output_widget.insert("end", "JANコードの重複と項目抜けをチェック中...\n")

def paste_and_execute():
    """
    テキスト貼り付けとコマンド実行のダミー関数
    """
    # ここに実際の処理内容を実装してください
    messagebox.showinfo("Info", "テキストに貼り付けて実行します。")

def execute_batch_and_skip_opening_file(batch_file):
    """
    バッチファイルを実行し、実行後にファイルを開かないダミー関数
    """
    try:
        subprocess.run(batch_file, shell=True)
        messagebox.showinfo("Info", f"{batch_file} を実行しました（ファイルは開かない）。")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def execute_batch_and_open_file(batch_file, output_file):
    """
    バッチファイルを実行した後、指定された出力ファイルを開くダミー関数
    """
    try:
        subprocess.run(batch_file, shell=True)
        open_file(output_file)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def clear_files():
    """
    クリップボードの内容をクリアするダミー関数
    """
    pyperclip.copy('')
    messagebox.showinfo("Info", "クリップボードをクリアしました。")

def show_under_construction():
    """
    機能が未実装である旨を通知するダミー関数
    """
    messagebox.showinfo("Info", "この機能は現在開発中です。")

def open_checksheet_with_prompt():
    """
    「商品情報入力シートに名前は入れましたか？」と尋ね、Yesならチェックシートを開く処理
    """
    answer = messagebox.askyesno("確認", "商品情報入力シートに名前は入れましたか？")
    if answer:
        open_file(r'C:\pysample_01\checksheet12.bat')

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        self.setup_ui()

    def setup_ui(self):
        # フレームの作成と配置
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)

        # テキストボックスの作成と配置
        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        # ボタンのリスト（テキスト, コマンド, 行, 列, 色）
        buttons = [
            ("重複と項目抜けのチェック", lambda: check_and_count_jan_codes(self.output), 1, 0, "#008000"),
            ("JANコード等のコピー", self.jancode_copy, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", paste_and_execute, 2, 1, "#008000"),
            ("データ抜けのチェック", lambda: execute_batch_and_skip_opening_file("Type1x.bat"), 2, 2, "#FF0000"),
            ("チェックシートを開く", open_checksheet_with_prompt, 3, 0, "#0078D4"),
            ("座標軸とコピー", show_under_construction, 2, 0, "#FF0000"),
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
        # ボタンの作成とグリッド配置
        for text, command, row, col, color in buttons:
            btn = customtkinter.CTkButton(self.frame, text=text, command=command, fg_color=color)
            btn.grid(row=row, column=col, padx=5, pady=2, sticky="ew")

        # グリッドのカラム幅を均等にする設定
        for i in range(4):
            self.frame.columnconfigure(i, weight=1)

        # 情報ラベルの追加
        self.info_label = customtkinter.CTkLabel(
            self.frame,
            text="※商品情報入力シートに必ず名前を明示すること",
            text_color="red",
            font=("メイリオ", 18, "bold")
        )
        self.info_label.grid(row=6, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

    def jancode_copy(self):
        """
        JANコードのコピー処理のダミー実装
        """
        messagebox.showinfo("Info", "JANコードをコピーしました。")

    def check_input_file(self):
        """
        input.txtのチェック処理のダミー実装
        """
        messagebox.showinfo("Info", "input.txt の内容をチェックします。")

    def open_subform(self):
        """
        サブフォームを開く処理のダミー実装
        """
        messagebox.showinfo("Info", "サブフォームを開きます。")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
