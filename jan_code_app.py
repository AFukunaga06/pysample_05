import customtkinter as ctk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyperclip
import threading
import time
import pyautogui
import os
import re
from PIL import Image

# Google Sheets APIの認証情報を設定
try:
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials_path = os.getenv('GOOGLE_CREDENTIALS_PATH', 'C:/pysample_01/samplep20240906-5ae36c9a4acd.json')
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(credentials)
    spreadsheet_key = os.getenv('SPREADSHEET_KEY', '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8')
    spreadsheet = client.open_by_key(spreadsheet_key)
    worksheet = spreadsheet.get_worksheet(0)
    print("Google Sheets接続成功")
except Exception as e:
    print(f"Google Sheets接続失敗: {str(e)}")
    print("デモモードで起動します")
    worksheet = None

class CombinedApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("JANコード等のコピー")
        
        # ウィンドウサイズを設定
        window_width = 220
        window_height = 600
        
        # 位置の計算
        if 'WINDOW_POSITION' in os.environ:
            x_pos, y_pos = map(int, os.environ['WINDOW_POSITION'].split(','))
        else:
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x_pos = (screen_width - window_width) // 2
            y_pos = (screen_height - window_height) // 2
        
        self.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
        self.resizable(False, False)
        
        # 変数の初期化
        self.current_index = 1
        self.start_pos = None
        self.end_pos = None
        self.total_start_time = None
        self.copied_content = None
        self.extract_start_pos = None
        self.extract_end_pos = None
        
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
            self, text="JANコードをコピー", command=self.copy_specified_index, 
            fg_color="#0078D7", text_color="white", hover_color="#0053A6", width=140
        )
        copy_specified_button.pack(pady=5)

        next_button = ctk.CTkButton(
            self, text="次へ進む", command=self.next_jan_code,
            fg_color="#0078D7", text_color="white", hover_color="#0053A6", width=140
        )
        next_button.pack(pady=5)

        previous_button = ctk.CTkButton(
            self, text="一つ戻る", command=self.previous_jan_code,
            fg_color="#0078D7", text_color="white", hover_color="#0053A6", width=140
        )
        previous_button.pack(pady=5)

        # セパレータ
        separator = ctk.CTkFrame(self, height=2, width=200)
        separator.pack(pady=20)

        # 座標取得セクション
        coord_label = ctk.CTkLabel(self, text="座標取得", font=("Arial", 14, "bold"))
        coord_label.pack(pady=(10, 10))

        self.start_button = ctk.CTkButton(
            self, text="開始位置", command=self.capture_start,
            fg_color="#0078D4", text_color="white", hover_color="#0053A6", width=140
        )
        self.start_button.pack(pady=5)
        
        self.end_button = ctk.CTkButton(
            self, text="終了位置", command=self.capture_end, state="disabled",
            fg_color="#0078D4", text_color="white", hover_color="#0053A6", width=140
        )
        self.end_button.pack(pady=5)

        # セパレータ
        separator2 = ctk.CTkFrame(self, height=2, width=200)
        separator2.pack(pady=20)

        # テキスト抽出機能セクション
        extract_label = ctk.CTkLabel(self, text="テキスト抽出", font=("Arial", 14, "bold"))
        extract_label.pack(pady=(10, 10))

        self.extract_start_button = ctk.CTkButton(
            self, text="抽出開始位置", command=self.capture_extract_start,
            fg_color="#28A745", text_color="white", hover_color="#1E7E34", width=140
        )
        self.extract_start_button.pack(pady=5)
        
        self.extract_end_button = ctk.CTkButton(
            self, text="抽出終了位置", command=self.capture_extract_end, state="disabled",
            fg_color="#28A745", text_color="white", hover_color="#1E7E34", width=140
        )
        self.extract_end_button.pack(pady=5)

        self.extract_execute_button = ctk.CTkButton(
            self, text="JANコード～重量抽出", command=self.execute_text_extract, state="disabled",
            fg_color="#DC3545", text_color="white", hover_color="#C82333", width=140
        )
        self.extract_execute_button.pack(pady=5)

    # JANコード関連の機能
    def copy_jan_code(self):
        try:
            if worksheet is None:
                messagebox.showwarning("デモモード", "Google Sheetsに接続されていません。\nデモ用JANコードをコピーします。")
                demo_jan_code = f"123456789012{self.current_index % 10}"
                pyperclip.copy(demo_jan_code)
                messagebox.showinfo("コピー完了", f"{self.current_index}番目のJANコード {demo_jan_code} をコピーしました。")
                threading.Thread(target=self.wait_and_paste).start()
                return
                
            cell_address = f'E{self.current_index + 1}'
            jan_code = worksheet.acell(cell_address).value
            if jan_code:
                pyperclip.copy(jan_code)
                messagebox.showinfo("コピー完了", f"{self.current_index}番目のJANコード {jan_code} をコピーしました。")
                threading.Thread(target=self.wait_and_paste).start()
            else:
                messagebox.showwarning("警告", "指定された位置にJANコードが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"JANコードの取得に失敗しました: {str(e)}")

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
        try:
            with open('coordinates_and_content.txt', 'w', encoding='utf-8') as f:
                f.write(f"開始位置: {self.start_pos}\n")
                f.write(f"終了位置: {self.end_pos}\n")
                f.write(f"コピーした内容:\n{self.copied_content}\n")
        except Exception as e:
            print(f"ファイル保存中にエラーが発生しました: {str(e)}")

    # テキスト抽出関連の機能
    def capture_extract_start(self):
        self.extract_start_pos = self.capture_position("抽出開始位置")
        if self.extract_start_pos:
            self.extract_end_button.configure(state="normal")
            print(f"抽出開始位置: {self.extract_start_pos}")

    def capture_extract_end(self):
        self.extract_end_pos = self.capture_position("抽出終了位置")
        if self.extract_end_pos:
            self.extract_execute_button.configure(state="normal")
            print(f"抽出終了位置: {self.extract_end_pos}")
            messagebox.showinfo("準備完了", "抽出範囲が設定されました。「JANコード～重量抽出」ボタンを押してください。")

    def execute_text_extract(self):
        try:
            self.drag_and_extract_text()
        except Exception as e:
            messagebox.showerror("エラー", f"テキスト抽出中にエラーが発生しました: {str(e)}")
        finally:
            self.reset_extract_buttons()

    def drag_and_extract_text(self):
        try:
            x1 = min(self.extract_start_pos[0], self.extract_end_pos[0])
            y1 = min(self.extract_start_pos[1], self.extract_end_pos[1])
            x2 = max(self.extract_start_pos[0], self.extract_end_pos[0])
            y2 = max(self.extract_start_pos[1], self.extract_end_pos[1])

            pyautogui.moveTo(x1, y1)
            pyautogui.mouseDown()
            pyautogui.moveTo(x2, y2, duration=0.5)
            pyautogui.mouseUp()
            
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.3)
            
            copied_text = pyperclip.paste()
            print(f"取得したテキスト: {copied_text}")

            extracted_text = self.extract_jan_to_weight(copied_text)
            
            if extracted_text:
                pyperclip.copy(extracted_text)
                messagebox.showinfo("抽出完了", f"以下のテキストをコピーしました:\n\n{extracted_text}")
                self.save_extract_result(copied_text, extracted_text)
            else:
                messagebox.showwarning("抽出失敗", "JANコードまたは重量情報が見つかりませんでした。\n取得したテキストをクリップボードにコピーします。")
                pyperclip.copy(copied_text)

        except Exception as e:
            messagebox.showerror("エラー", f"テキスト抽出中にエラーが発生しました: {str(e)}")

    def extract_jan_to_weight(self, text):
        try:
            jan_pattern = r'\b\d{13}\b|\b\d{8}\b'
            weight_patterns = [
                r'\d+\.?\d*\s*[gG](?![a-zA-Z])',
                r'\d+\.?\d*\s*グラム',
                r'重量\s*[:：]?\s*\d+\.?\d*\s*[gG]?',
                r'\d+\.?\d*\s*ｇ',
            ]

            jan_matches = re.finditer(jan_pattern, text)
            jan_positions = [(match.start(), match.end(), match.group()) for match in jan_matches]

            if not jan_positions:
                print("JANコードが見つかりません")
                return None

            weight_positions = []
            for pattern in weight_patterns:
                weight_matches = re.finditer(pattern, text, re.IGNORECASE)
                weight_positions.extend([(match.start(), match.end(), match.group()) for match in weight_matches])

            if not weight_positions:
                print("重量情報が見つかりません")
                return None

            best_combination = None
            min_distance = float('inf')

            for jan_start, jan_end, jan_code in jan_positions:
                for weight_start, weight_end, weight_text in weight_positions:
                    if weight_start >= jan_end:
                        distance = weight_start - jan_end
                    else:
                        distance = jan_start - weight_end
                    
                    if distance >= 0 and distance < min_distance:
                        min_distance = distance
                        if weight_start >= jan_end:
                            best_combination = (jan_start, weight_end, jan_code, weight_text)
                        else:
                            best_combination = (weight_start, jan_end, jan_code, weight_text)

            if best_combination:
                start_pos, end_pos, jan_code, weight_text = best_combination
                extracted_text = text[start_pos:end_pos].strip()
                extracted_text = re.sub(r'\s+', ' ', extracted_text)
                print(f"抽出されたテキスト: {extracted_text}")
                return extracted_text
            else:
                print("適切なJANコードと重量の組み合わせが見つかりません")
                return None

        except Exception as e:
            print(f"テキスト抽出中にエラーが発生しました: {str(e)}")
            return None

    def save_extract_result(self, full_text, extracted_text):
        try:
            with open('extract_result.txt', 'w', encoding='utf-8') as f:
                f.write("=== テキスト抽出結果 ===\n")
                f.write(f"日時: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"抽出範囲: {self.extract_start_pos} ～ {self.extract_end_pos}\n\n")
                f.write("--- 取得したテキスト ---\n")
                f.write(f"{full_text}\n\n")
                f.write("--- 抽出されたテキスト ---\n")
                f.write(f"{extracted_text}\n\n")
        except Exception as e:
            print(f"抽出結果の保存に失敗しました: {str(e)}")

    def reset_extract_buttons(self):
        self.extract_end_button.configure(state="disabled")
        self.extract_execute_button.configure(state="disabled")
        self.extract_start_pos = None
        self.extract_end_pos = None

if __name__ == '__main__':
    app = CombinedApp()
    app.mainloop()