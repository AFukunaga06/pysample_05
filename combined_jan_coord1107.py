# combined_jan_coord1107.py

import customtkinter as ctk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyperclip
import threading
import time
import pyautogui
import os

class CombinedApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("JANコードコピーと座標軸")
        self.geometry("300x600")

        # Google Sheets API設定
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/pysample_01/samplep20240906-5ae36c9a4acd.json', scope)
        client = gspread.authorize(credentials)
        
        # スプレッドシート設定
        spreadsheet_key = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
        spreadsheet = client.open_by_key(spreadsheet_key)
        self.worksheet = spreadsheet.get_worksheet(0)
        
        # 変数の初期化
        self.current_index = 1
        self.start_pos = None
        self.end_pos = None
        self.total_start_time = None
        self.copied_content = None
        
        self.setup_ui()
        self.setup_drag()

    def setup_ui(self):
        # メインフレーム
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # タイトルラベル
        title_label = ctk.CTkLabel(main_frame, text="JANコードコピーと座標軸",
                                 font=("Arial", 18, "bold"))
        title_label.pack(pady=(0, 10))

        # JANコードセクション
        jan_frame = ctk.CTkFrame(main_frame)
        jan_frame.pack(fill="x", pady=10, padx=5)
        
        jan_label = ctk.CTkLabel(jan_frame, text="JANコードコピー",
                                font=("Arial", 14, "bold"))
        jan_label.pack(pady=5)
        
        # 入力フィールドフレーム
        input_frame = ctk.CTkFrame(jan_frame)
        input_frame.pack(fill="x", pady=5)
        
        self.index_label = ctk.CTkLabel(input_frame, text="指定した番目:")
        self.index_label.pack(pady=5)
        
        self.index_entry = ctk.CTkEntry(input_frame, width=120)
        self.index_entry.pack(pady=5)
        
        # JANコードボタン
        jan_buttons_frame = ctk.CTkFrame(jan_frame)
        jan_buttons_frame.pack(fill="x", pady=5)
        
        copy_button = ctk.CTkButton(jan_buttons_frame,
                                  text="JANコードをコピー",
                                  command=self.copy_specified_index,
                                  fg_color="#0078D7",
                                  hover_color="#0053A6")
        copy_button.pack(fill="x", pady=2)
        
        next_button = ctk.CTkButton(jan_buttons_frame,
                                  text="次へ進む",
                                  command=self.next_jan_code,
                                  fg_color="#0078D7",
                                  hover_color="#0053A6")
        next_button.pack(fill="x", pady=2)
        
        prev_button = ctk.CTkButton(jan_buttons_frame,
                                  text="一つ戻る",
                                  command=self.previous_jan_code,
                                  fg_color="#0078D7",
                                  hover_color="#0053A6")
        prev_button.pack(fill="x", pady=2)
        
        # 区切り線
        separator = ctk.CTkFrame(main_frame, height=2, fg_color="gray70")
        separator.pack(fill="x", pady=15)
        
        # 座標軸セクション
        coord_frame = ctk.CTkFrame(main_frame)
        coord_frame.pack(fill="x", pady=10, padx=5)
        
        coord_label = ctk.CTkLabel(coord_frame, text="座標軸コピー",
                                 font=("Arial", 14, "bold"))
        coord_label.pack(pady=5)
        
        self.start_button = ctk.CTkButton(coord_frame,
                                        text="開始位置",
                                        command=self.capture_start,
                                        fg_color="#0078D4",
                                        corner_radius=10,
                                        width=180,
                                        height=40)
        self.start_button.pack(pady=10)
        
        self.end_button = ctk.CTkButton(coord_frame,
                                      text="終了位置",
                                      command=self.capture_end,
                                      state="disabled",
                                      fg_color="#0078D4",
                                      corner_radius=10,
                                      width=180,
                                      height=40)
        self.end_button.pack(pady=10)

        # ステータス表示用ラベル
        self.status_label = ctk.CTkLabel(main_frame, text="")
        self.status_label.pack(pady=5)

    def setup_drag(self):
        self.bind("<Button-1>", self.on_drag_start)
        self.bind("<B1-Motion>", self.on_drag_motion)

    def on_drag_start(self, event):
        self.startX = event.x
        self.startY = event.y

    def on_drag_motion(self, event):
        x = self.winfo_x() - self.startX + event.x
        y = self.winfo_y() - self.startY + event.y
        self.geometry(f"+{x}+{y}")

    def update_status(self, message, clear_after=3000):
        self.status_label.configure(text=message)
        if clear_after:
            self.after(clear_after, lambda: self.status_label.configure(text=""))

    def copy_jan_code(self):
        try:
            cell_address = f'E{self.current_index + 1}'
            jan_code = self.worksheet.acell(cell_address).value
            if jan_code:
                pyperclip.copy(jan_code)
                self.update_status(f"{self.current_index}番目のJANコード {jan_code} をコピーしました")
                threading.Thread(target=self.wait_and_paste).start()
            else:
                self.update_status("JANコードが見つかりません")
        except Exception as e:
            self.update_status(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"JANコードのコピー中にエラーが発生しました：{str(e)}")

    def wait_and_paste(self):
        time.sleep(3.0)
        try:
            pyautogui.hotkey('ctrl', 'v')
        except Exception as e:
            self.update_status(f"貼り付けエラー: {str(e)}")

    def next_jan_code(self):
        self.current_index += 1
        self.copy_jan_code()

    def previous_jan_code(self):
        if self.current_index > 1:
            self.current_index -= 1
            self.copy_jan_code()
        else:
            self.update_status("これが最初のJANコードです")

    def copy_specified_index(self):
        try:
            specified_index = int(self.index_entry.get())
            if specified_index < 1:
                self.update_status("1以上の数値を入力してください")
                return
            self.current_index = specified_index
            self.copy_jan_code()
        except ValueError:
            self.update_status("有効な数値を入力してください")

    def capture_start(self):
        self.total_start_time = time.time()
        self.update_status("開始位置を取得中...")
        self.start_pos = self.capture_position("開始位置")
        if self.start_pos:
            self.end_button.configure(state="normal")
            self.update_status("開始位置を取得しました")

    def capture_end(self):
        self.update_status("終了位置を取得中...")
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
            self.update_status(f"エラー: {str(e)}")
            return None
        finally:
            self.deiconify()

    def perform_drag_and_copy(self):
        try:
            self.update_status("コピー操作を実行中...")
            pyautogui.moveTo(self.start_pos[0], self.start_pos[1])
            pyautogui.mouseDown()
            pyautogui.moveTo(self.end_pos[0], self.end_pos[1], duration=0.5)
            pyautogui.mouseUp()
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.4)
            self.copied_content = pyperclip.paste()
            total_end_time = time.time()
            total_time = total_end_time - self.total_start_time
            self.update_status(f"コピー完了 (所要時間: {total_time:.1f}秒)")
            self.save_coordinates_and_content()
        except Exception as e:
            self.update_status(f"エラー: {str(e)}")
        finally:
            self.end_button.configure(state="disabled")
            self.start_button.configure(state="normal")
            self.start_pos = None
            self.end_pos = None

    def save_coordinates_and_content(self):
        try:
            with open('coordinates_and_content.txt', 'w', encoding='utf-8') as f:
                f.write(f"開始位置: {self.start_pos}\n")
                f.write(f"終了位置: {self.end_pos}\n")
                f.write(f"コピーした内容:\n{self.copied_content}\n")
            self.update_status("座標とコンテンツを保存しました")
        except Exception as e:
            self.update_status(f"保存エラー: {str(e)}")

if __name__ == '__main__':
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    app = CombinedApp()
    app.mainloop()