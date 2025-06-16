import customtkinter as ctk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyperclip
import threading
import time
import pyautogui
import os

# Google Sheets APIの認証情報を設定
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/pysample_01/samplep20240906-5ae36c9a4acd.json', scope)
client = gspread.authorize(credentials)

# スプレッドシートを開く
spreadsheet_key = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
spreadsheet = client.open_by_key(spreadsheet_key)
worksheet = spreadsheet.get_worksheet(0)

class CombinedApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("JANコード等のコピー")
        
        # ウィンドウサイズを設定
        window_width = 220
        window_height = 500
        
        # 位置の計算
        if 'WINDOW_POSITION' in os.environ:
            # 親ウィンドウからの相対位置を取得
            x_pos, y_pos = map(int, os.environ['WINDOW_POSITION'].split(','))
        else:
            # デフォルトの中央配置
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x_pos = (screen_width - window_width) // 2
            y_pos = (screen_height - window_height) // 2
        
        # ジオメトリを設定
        self.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
        
        # ウィンドウを固定
        self.resizable(False, False)
        
        # 変数の初期化
        self.current_index = 1
        self.start_pos = None
        self.end_pos = None
        self.total_start_time = None
        self.copied_content = None
        
        self.setup_ui()

    def setup_ui(self):
        # JANコードコピーセクション
        jan_label = ctk.CTkLabel(self, text="JANコードコピー", font=("Arial", 14, "bold"))
        jan_label.pack(pady=(20, 10))
        
        index_label = ctk.CTkLabel(self, text="指定した番目:")
        index_label.pack(pady=5)

        self.index_entry = ctk.CTkEntry(self, width=120)
        self.index_entry.pack(pady=5)

        copy_specified_button = ctk.CTkButton(
            self, 
            text="JANコードをコピー", 
            command=self.copy_specified_index, 
            fg_color="#0078D7",
            text_color="white",
            hover_color="#0053A6",
            width=140
        )
        copy_specified_button.pack(pady=5)

        next_button = ctk.CTkButton(
            self,
            text="次へ進む",
            command=self.next_jan_code,
            fg_color="#0078D7",
            text_color="white",
            hover_color="#0053A6",
            width=140
        )
        next_button.pack(pady=5)

        previous_button = ctk.CTkButton(
            self,
            text="一つ戻る",
            command=self.previous_jan_code,
            fg_color="#0078D7",
            text_color="white",
            hover_color="#0053A6",
            width=140
        )
        previous_button.pack(pady=5)

        # セパレータ
        separator = ctk.CTkFrame(self, height=2, width=200)
        separator.pack(pady=20)

        # 座標取得セクション
        coord_label = ctk.CTkLabel(self, text="座標取得", font=("Arial", 14, "bold"))
        coord_label.pack(pady=(10, 10))

        self.start_button = ctk.CTkButton(
            self,
            text="開始位置",
            command=self.capture_start,
            fg_color="#0078D4",
            text_color="white",
            hover_color="#0053A6",
            width=140
        )
        self.start_button.pack(pady=5)
        
        self.end_button = ctk.CTkButton(
            self,
            text="終了位置",
            command=self.capture_end,
            state="disabled",
            fg_color="#0078D4",
            text_color="white",
            hover_color="#0053A6",
            width=140
        )
        self.end_button.pack(pady=5)

    # JANコード関連の機能
    def copy_jan_code(self):
        cell_address = f'E{self.current_index + 1}'
        jan_code = worksheet.acell(cell_address).value
        pyperclip.copy(jan_code)
        messagebox.showinfo("コピー完了", f"{self.current_index}番目のJANコード {jan_code} をコピーしました。")
        threading.Thread(target=self.wait_and_paste).start()

    def wait_and_paste(self):
        time.sleep(3.0)
        pyautogui.hotkey('ctrl', 'v')

    def next_jan_code(self):
        self.current_index += 1
        self.copy_jan_code()

    def previous_jan_code(self):
        if self.current_index > 1:
            self.current_index -= 1
            self.copy_jan_code()
        else:
            messagebox.showinfo("情報", "すでに最初のJANコードです。")

    def copy_specified_index(self):
        try:
            specified_index = int(self.index_entry.get())
            if specified_index < 1:
                messagebox.showerror("エラー", "指定された番目は1以上の整数である必要があります。")
            else:
                self.current_index = specified_index
                self.copy_jan_code()
        except ValueError:
            messagebox.showerror("エラー", "有効な整数を入力してください。")

    # 座標取得関連の機能
    def capture_start(self):
        self.total_start_time = time.time()
        self.start_pos = self.capture_position("開始位置")
        if self.start_pos:
            self.end_button.configure(state="normal")

    def capture_end(self):
        self.end_pos = self.capture_position("終了位置")
        if self.end_pos:
            self.after(1000, self.perform_drag_and_copy)

    def capture_position(self, position_type):
        self.withdraw()
        time.sleep(1.5)
        try:
            x1, y1 = pyautogui.position()
            time.sleep(0.5)
            x2, y2 = pyautogui.position()
            avg_x = (x1 + x2) / 2
            avg_y = (y1 + y2) / 2
            return (avg_x, avg_y)
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")
            return None
        finally:
            self.deiconify()

    def perform_drag_and_copy(self):
        try:
            pyautogui.moveTo(self.start_pos[0], self.start_pos[1])
            pyautogui.mouseDown()
            pyautogui.moveTo(self.end_pos[0], self.end_pos[1], duration=0.5)
            pyautogui.mouseUp()
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.4)
            self.copied_content = pyperclip.paste()
            total_end_time = time.time()
            total_time = total_end_time - self.total_start_time
            print(f"全体の所要時間: {total_time}秒")
            self.save_coordinates_and_content()
        except Exception as e:
            print(f"ドラッグ＆コピー中にエラーが発生しました: {str(e)}")
        finally:
            self.end_button.configure(state="disabled")
            self.start_button.configure(state="normal")
            self.start_pos = None
            self.end_pos = None

    def save_coordinates_and_content(self):
        with open('coordinates_and_content.txt', 'w', encoding='utf-8') as f:
            f.write(f"開始位置: {self.start_pos}\n")
            f.write(f"終了位置: {self.end_pos}\n")
            f.write(f"コピーした内容:\n{self.copied_content}\n")

if __name__ == '__main__':
    app = CombinedApp()
    app.mainloop()