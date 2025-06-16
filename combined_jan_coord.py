# combined_jan_coord.py

import customtkinter as ctk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyperclip
import threading
import time
import pyautogui
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

class CombinedApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("JANコードコピー")
        self.geometry("250x500")

        # デュアルモニター環境での初期位置設定
        try:
            total_width = pyautogui.size()[0]  # 全画面幅を取得
            single_monitor_width = 1920  # 23インチフルHDモニターの幅

            # 右側のモニターの中央よりやや左に配置
            window_x = single_monitor_width + (single_monitor_width - 300) // 2
            window_y = 100  # 上から100ピクセルの位置

            # ウィンドウ位置を設定
            self.geometry(f"+{window_x}+{window_y}")

            # ウィンドウを常に最前面に表示
            self.attributes('-topmost', True)
            
        except Exception as e:
            print(f"ウィンドウ位置の設定に失敗しました: {str(e)}")
            self.geometry("+100+100")  # デフォルト位置

        # Google Sheets API設定
        self.setup_google_sheets()
        
        # 変数の初期化
        self.current_index = 1
        self.start_pos = None
        self.end_pos = None
        self.total_start_time = None
        self.copied_content = None
        
        self.setup_ui()
        self.setup_drag()  # ドラッグ機能を有効化

    def setup_google_sheets(self):
        try:
            self.scope = ['https://spreadsheets.google.com/feeds',
                         'https://www.googleapis.com/auth/drive']
            self.credentials = ServiceAccountCredentials.from_json_keyfile_name(
                'C:/pysample_01/samplep20240906-5ae36c9a4acd.json', 
                self.scope
            )
            self.client = gspread.authorize(self.credentials)
            self.spreadsheet_key = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
            self.spreadsheet = self.client.open_by_key(self.spreadsheet_key)
            self.worksheet = self.spreadsheet.get_worksheet(0)
        except Exception as e:
            messagebox.showerror("初期化エラー", f"Google Sheets APIの初期化に失敗しました：{str(e)}")

    def setup_ui(self):
        # JANコードセクション
        jan_section = ctk.CTkFrame(self)
        jan_section.pack(fill="x", padx=10, pady=(10,5))
        
        # JANコードコピーのタイトル（グレー背景）
        title_frame = ctk.CTkFrame(jan_section, fg_color="gray80")
        title_frame.pack(fill="x", padx=0, pady=0)
        ctk.CTkLabel(title_frame, text="JANコードコピー").pack(pady=5)

        # 入力フィールド
        input_frame = ctk.CTkFrame(jan_section)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(input_frame, text="指定した番目:").pack(pady=2)
        self.index_entry = ctk.CTkEntry(input_frame)
        self.index_entry.pack(pady=2)
        self.index_entry.insert(0, "1")

        # JANコードボタン
        button_frame = ctk.CTkFrame(jan_section)
        button_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(button_frame,
                     text="JANコードをコピー",
                     command=self.safe_copy_jan_code,
                     fg_color="#0078D7").pack(fill="x", pady=2)
        
        ctk.CTkButton(button_frame,
                     text="次へ進む",
                     command=self.next_jan_code,
                     fg_color="#0078D7").pack(fill="x", pady=2)
        
        ctk.CTkButton(button_frame,
                     text="一つ戻る",
                     command=self.previous_jan_code,
                     fg_color="#0078D7").pack(fill="x", pady=2)

        # 座標軸セクション
        coord_section = ctk.CTkFrame(self)
        coord_section.pack(fill="x", padx=10, pady=5)
        
        # 座標軸コピーのタイトル（グレー背景）
        coord_title_frame = ctk.CTkFrame(coord_section, fg_color="gray80")
        coord_title_frame.pack(fill="x", padx=0, pady=0)
        ctk.CTkLabel(coord_title_frame, text="座標軸コピー").pack(pady=5)

        # 座標軸ボタン
        coord_button_frame = ctk.CTkFrame(coord_section)
        coord_button_frame.pack(fill="x", padx=10, pady=5)
        
        self.start_button = ctk.CTkButton(coord_button_frame,
                                        text="開始位置",
                                        command=self.capture_start,
                                        fg_color="#0078D7")
        self.start_button.pack(fill="x", pady=2)
        
        self.end_button = ctk.CTkButton(coord_button_frame,
                                      text="終了位置",
                                      command=self.capture_end,
                                      state="disabled",
                                      fg_color="#0078D7")
        self.end_button.pack(fill="x", pady=2)

    def setup_drag(self):
        # ドラッグ機能を有効化
        self.bind("<Button-1>", self.on_drag_start)
        self.bind("<B1-Motion>", self.on_drag_motion)

    def on_drag_start(self, event):
        self.startX = event.x
        self.startY = event.y

    def on_drag_motion(self, event):
        x = self.winfo_x() - self.startX + event.x
        y = self.winfo_y() - self.startY + event.y
        self.geometry(f"+{x}+{y}")

    def refresh_credentials(self):
        try:
            if self.credentials.access_token_expired:
                self.credentials.refresh(Request())
                self.client = gspread.authorize(self.credentials)
                self.spreadsheet = self.client.open_by_key(self.spreadsheet_key)
                self.worksheet = self.spreadsheet.get_worksheet(0)
                return True
        except Exception as e:
            messagebox.showerror("認証エラー", f"認証の更新に失敗しました：{str(e)}")
            return False
        return True

    def safe_copy_jan_code(self):
        try:
            if not self.refresh_credentials():
                return
            self.copy_jan_code()
        except Exception as e:
            messagebox.showerror("エラー", f"JANコードのコピー中にエラーが発生しました：{str(e)}")
            self.setup_google_sheets()

    def copy_jan_code(self):
        try:
            cell_address = f'E{self.current_index + 1}'
            jan_code = self.worksheet.acell(cell_address).value
            if jan_code:
                pyperclip.copy(jan_code)
                messagebox.showinfo("コピー完了", f"{self.current_index}番目のJANコード {jan_code} をコピーしました。")
                threading.Thread(target=self.wait_and_paste).start()
            else:
                messagebox.showwarning("警告", "JANコードが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"JANコードの取得中にエラーが発生しました：{str(e)}")

    def wait_and_paste(self):
        time.sleep(3.0)
        try:
            pyautogui.hotkey('ctrl', 'v')
        except Exception as e:
            print(f"貼り付けエラー: {str(e)}")

    def next_jan_code(self):
        self.current_index += 1
        self.safe_copy_jan_code()

    def previous_jan_code(self):
        if self.current_index > 1:
            self.current_index -= 1
            self.safe_copy_jan_code()
        else:
            messagebox.showinfo("情報", "これが最初のJANコードです。")

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
        """座標取得処理（制限なし）"""
        self.withdraw()
        time.sleep(1.5)
        try:
            x, y = pyautogui.position()
            return (x, y)
        except Exception as e:
            messagebox.showerror("エラー", f"座標の取得中にエラーが発生しました：{str(e)}")
            return None
        finally:
            self.deiconify()

    def perform_drag_and_copy(self):
        """ドラッグ＆コピー処理（制限なし）"""
        try:
            if not all([self.start_pos, self.end_pos]):
                raise Exception("開始位置または終了位置が設定されていません。")

            start_x, start_y = self.start_pos
            end_x, end_y = self.end_pos

            pyautogui.moveTo(start_x, start_y)
            pyautogui.mouseDown()
            pyautogui.moveTo(end_x, end_y, duration=0.5)
            pyautogui.mouseUp()
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.4)
            
            self.copied_content = pyperclip.paste()
            total_end_time = time.time()
            total_time = total_end_time - self.total_start_time
            print(f"全体の所要時間: {total_time:.1f}秒")
            self.save_coordinates_and_content()
            
        except Exception as e:
            messagebox.showerror("エラー", f"ドラッグ＆コピー中にエラーが発生しました：{str(e)}")
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
        except Exception as e:
            messagebox.showerror("エラー", f"データの保存中にエラーが発生しました：{str(e)}")

if __name__ == '__main__':
    try:
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        app = CombinedApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("起動エラー", f"アプリケーションの起動中にエラーが発生しました：{str(e)}")