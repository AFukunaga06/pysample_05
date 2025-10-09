import customtkinter as ctk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyperclip
import threading
import time
import pyautogui

# Google Sheets APIの認証情報を設定
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/pysample_01/samplep20240906-5ae36c9a4acd.json', scope)
client = gspread.authorize(credentials)

# スプレッドシートを開く
spreadsheet_key = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
spreadsheet = client.open_by_key(spreadsheet_key)
worksheet = spreadsheet.get_worksheet(0)

# GUIの設定
window = ctk.CTk()
window.title("JANコードコピー")
window.geometry("220x150")  # ウィンドウサイズを固定

# ドラッグ可能にするための関数
def on_drag_start(event):
    window.startX = event.x
    window.startY = event.y

def on_drag_motion(event):
    x = window.winfo_x() - window.startX + event.x
    y = window.winfo_y() - window.startY + event.y
    window.geometry(f"+{x}+{y}")

# [残りのコードは元のまま]

# ウィジェットの設定
index_label = ctk.CTkLabel(window, text="指定した番目:")
index_label.pack(pady=10)

index_entry = ctk.CTkEntry(window, width=120)
index_entry.pack(pady=10)

copy_specified_button = ctk.CTkButton(window, text="JANコードをコピー", command=copy_specified_index, fg_color="#0078D7", text_color="white", hover_color="#0053A6")
copy_specified_button.pack(pady=10)

next_button = ctk.CTkButton(window, text="次へ進む", command=next_jan_code, fg_color="#0078D7", text_color="white", hover_color="#0053A6")
next_button.pack(pady=10)

previous_button = ctk.CTkButton(window, text="一つ戻る", command=previous_jan_code, fg_color="#0078D7", text_color="white", hover_color="#0053A6")
previous_button.pack(pady=10)

window.mainloop()