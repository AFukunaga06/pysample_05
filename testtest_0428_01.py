import customtkinter as ctk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyperclip
import threading
import time
import pyautogui
import os

# Google Sheets API の認証情報を設定
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json', scope)
client = gspread.authorize(credentials)

# スプレッドシートを開く
spreadsheet_key = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
spreadsheet = client.open_by_key(spreadsheet_key)
worksheet = spreadsheet.get_worksheet(0)

class CombinedApp(ctk.CTk):
    # 1 cm あたりのピクセル数（モニタ DPI に応じて調整してください）
    PIXELS_PER_CM = 37.8  
    # 1 行分の高さ（ピクセル）。実際のフォントサイズに合わせて調整してください
    LINE_HEIGHT_PX = 20    

    def __init__(self):
        super().__init__()
        self.title("JANコード等のコピー")
        self.geometry("220x300")
        self.resizable(False, False)
        self.start_pos = None
        self.setup_ui()

    def setup_ui(self):
        coord_label = ctk.CTkLabel(self, text="座標取得＆ドラッグ", font=("Arial", 14, "bold"))
        coord_label.pack(pady=(20, 10))

        self.start_button = ctk.CTkButton(
            self, text="開始位置",
            command=self.capture_start,
            fg_color="#0078D4", text_color="white",
            hover_color="#0053A6", width=140
        )
        self.start_button.pack(pady=10)

    def capture_start(self):
        """開始位置を取得し、その位置から下7行・右1cm の位置までドラッグ＆コピー"""
        self.withdraw()
        time.sleep(1.0)  # ウィンドウを隠す時間
        x, y = pyautogui.position()
        self.deiconify()

        self.start_pos = (x, y)
        # 終了位置を計算
        end_x = x + CombinedApp.PIXELS_PER_CM
        end_y = y + 7 * CombinedApp.LINE_HEIGHT_PX

        # 別スレッドでドラッグ＆コピーを実行
        threading.Thread(target=self.drag_and_copy, args=((end_x, end_y),)).start()

    def drag_and_copy(self, end_pos):
        """計算した end_pos までドラッグ＆コピー"""
        try:
            # ドラッグ開始
            pyautogui.moveTo(self.start_pos[0], self.start_pos[1])
            pyautogui.mouseDown()
            # ドラッグ終了
            pyautogui.moveTo(end_pos[0], end_pos[1], duration=0.4)
            pyautogui.mouseUp()
            # コピー
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            copied = pyperclip.paste()
            messagebox.showinfo("コピー完了", f"選択範囲をコピーしました：\n{copied}")
        except Exception as e:
            messagebox.showerror("エラー", f"ドラッグ＆コピー中にエラーが発生しました：\n{e}")

if __name__ == '__main__':
    app = CombinedApp()
    app.mainloop()
