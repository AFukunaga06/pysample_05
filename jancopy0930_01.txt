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

# ドラッグ可能にするための関数
def on_drag_start(event):
    window.startX = event.x
    window.startY = event.y

def on_drag_motion(event):
    x = window.winfo_x() - window.startX + event.x
    y = window.winfo_y() - window.startY + event.y
    window.geometry(f"+{x}+{y}")

# 現在のインデックスを保持する変数
current_index = 1

def copy_jan_code():
    global current_index
    cell_address = f'E{current_index + 1}'
    jan_code = worksheet.acell(cell_address).value
    pyperclip.copy(jan_code)
    messagebox.showinfo("コピー完了", f"{current_index}番目のJANコード {jan_code} をコピーしました。")
    threading.Thread(target=wait_and_paste).start()

def wait_and_paste():
    time.sleep(3.0)  # 3.0秒待機に変更
    pyautogui.hotkey('ctrl', 'v')  # 貼り付け

def next_jan_code():
    global current_index
    current_index += 1
    copy_jan_code()

def previous_jan_code():
    global current_index
    if current_index > 1:
        current_index -= 1
        copy_jan_code()
    else:
        messagebox.showinfo("情報", "すでに最初のJANコードです。")

def copy_specified_index():
    global current_index
    try:
        specified_index = int(index_entry.get())
        if specified_index < 1:
            messagebox.showerror("エラー", "指定された番目は1以上の整数である必要があります。")
        else:
            current_index = specified_index
            copy_jan_code()
    except ValueError:
        messagebox.showerror("エラー", "有効な整数を入力してください。")

# ドラッグイベントをバインド
window.bind("<Button-1>", on_drag_start)
window.bind("<B1-Motion>", on_drag_motion)

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
