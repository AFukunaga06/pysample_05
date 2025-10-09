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

class WindowPositionController:
    def __init__(self):
        self.main_window = None
        self.jancopy_window = None
        self.copycoord_window = None

    def set_main_window(self, window):
        self.main_window = window
        
    def update_window_positions(self):
        if not self.main_window:
            return
            
        # メインウィンドウの位置を取得
        main_x = self.main_window.winfo_x()
        main_y = self.main_window.winfo_y()
        
        # JANコピーウィンドウの位置設定
        if self.jancopy_window and self.jancopy_window.winfo_exists():
            jan_width = 220  # JANコピーウィンドウの幅
            self.jancopy_window.geometry(f"{jan_width}x150+{main_x - jan_width - 5}+{main_y}")
            
        # 座標コピーウィンドウの位置設定
        if self.copycoord_window and self.copycoord_window.winfo_exists():
            coord_width = 220  # 座標コピーウィンドウの幅
            coord_y = main_y + 155  # JANコピーウィンドウの高さ + 5ピクセル
            self.copycoord_window.geometry(f"{coord_width}x150+{main_x - coord_width - 5}+{coord_y}")
        
        # 1秒後に再度位置を更新
        if self.main_window:
            self.main_window.after(1000, self.update_window_positions)

    def register_jancopy_window(self, window):
        self.jancopy_window = window
        self.update_window_positions()

    def register_copycoord_window(self, window):
        self.copycoord_window = window
        self.update_window_positions()

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
    
    
    class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        self.geometry("800x600")
        
        # ウィンドウポジションコントローラーの初期化
        self.window_controller = WindowPositionController()
        self.window_controller.set_main_window(self)
        
        # 画面中央に配置
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 800) // 2
        y = (screen_height - 600) // 2
        self.geometry(f"800x600+{x}+{y}")
        
        self.setup_ui()
        
        # 位置の更新を開始
        self.window_controller.update_window_positions()

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)
        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        buttons = [
            ("重複と項目抜けのチェック", lambda: check_and_count_jan_codes(self.output), 1, 0, "#008000"),
            ("JANコードのコピー", self.jancode_copy, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", paste_and_execute, 2, 1, "#008000"),
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

    @safe_file_operation
    def jancode_copy(self):
        if not hasattr(self, 'jancopy_window') or not self.jancopy_window or not self.jancopy_window.winfo_exists():
            self.jancopy_window = customtkinter.CTkToplevel()
            self.window_controller.register_jancopy_window(self.jancopy_window)
            subprocess.Popen(['python', 'jancopy0906_01.py'])

    @safe_file_operation
    def open_coordinate_subform(self):
        if not hasattr(self, 'copycoord_window') or not self.copycoord_window or not self.copycoord_window.winfo_exists():
            self.copycoord_window = customtkinter.CTkToplevel()
            self.window_controller.register_copycoord_window(self.copycoord_window)
            subprocess.Popen(['python', 'Copycoord0902_01.py'])

    # [元のコードの他のメソッドはそのまま維持]

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()
    