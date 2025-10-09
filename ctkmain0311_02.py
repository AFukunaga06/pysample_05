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

# --- 省略: スプレッドシート認証やその他の関数などはそのまま ---

def open_checksheet_with_prompt():
    """
    「商品情報入力シートに名前は入れましたか？」と尋ねて、
    Yes(はい)ならチェックシートを開く、No(いいえ)なら開かない。
    """
    answer = messagebox.askyesno("確認", "商品情報入力シートに名前は入れましたか？")
    if answer:
        open_file(r'C:\pysample_01\checksheet12.bat')

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        # ... 省略: ウィンドウの設定や UI セットアップ ...

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)

        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        buttons = [
            ("重複と項目抜けのチェック", lambda: check_and_count_jan_codes(self.output), 1, 0, "#008000"),
            ("JANコード等のコピー", self.jancode_copy, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", paste_and_execute, 2, 1, "#008000"),
            ("データ抜けのチェック", lambda: execute_batch_and_skip_opening_file("Type1x.bat"), 2, 2, "#FF0000"),

            # ★ ここを変更 ★
            # ("チェックシートを開く", lambda: open_file(r'C:\pysample_01\checksheet12.bat'), 3, 0, "#0078D4"),
            # 上記を以下のように置き換え、クリック時に先にメッセージを表示するようにする
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
        for text, command, row, col, color in buttons:
            btn = customtkinter.CTkButton(self.frame, text=text, command=command, fg_color=color)
            btn.grid(row=row, column=col, padx=5, pady=2, sticky="ew")

        # カラムの伸縮設定
        for i in range(4):
            self.frame.columnconfigure(i, weight=1)

        # --- ここから追加 ---
        self.info_label = customtkinter.CTkLabel(
            self.frame,
            text="※商品情報入力シートに必ず名前を明示すること",
            text_color="red",
            font=("メイリオ", 18, "bold")
        )
        self.info_label.grid(row=6, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
