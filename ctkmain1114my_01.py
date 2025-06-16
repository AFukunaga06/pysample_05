#ctknain1114my_01.py
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
import threading
from oauth2client.service_account import ServiceAccountCredentials

# スプレッドシートの設定
SPREADSHEET_ID = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
KEY_FILE_PATH = r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json'

# Google認証の設定
def setup_google_auth():
    creds = Credentials.from_service_account_file(
        KEY_FILE_PATH, 
        scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).sheet1

# 座標取得アプリのクラス
class CopyCoordApp(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.title("座標取得アプリ")
        self.geometry("220x150")
        self.setup_ui()
        self.start_pos = None
        self.end_pos = None
        self.total_start_time = None
        self.copied_content = None

    def setup_ui(self):
        self.start_button = customtkinter.CTkButton(
            self, 
            text="開始位置", 
            command=self.capture_start, 
            fg_color="#0078D4", 
            corner_radius=10,
            width=180,
            height=40
        )
        self.start_button.pack(pady=10, padx=20)
        
        self.end_button = customtkinter.CTkButton(
            self, 
            text="終了位置", 
            command=self.capture_end, 
            state="disabled",
            fg_color="#0078D4", 
            corner_radius=10,
            width=180,
            height=40
        )
        self.end_button.pack(pady=10, padx=20)

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
            x1, y1 = p.position()
            time.sleep(0.5)
            x2, y2 = p.position()
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
            p.moveTo(self.start_pos[0], self.start_pos[1])
            p.mouseDown()
            p.moveTo(self.end_pos[0], self.end_pos[1], duration=0.5)
            p.mouseUp()
            p.hotkey('ctrl', 'c')
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

# JANコードコピーアプリのクラス
class JanCopyApp(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.title("JANコードコピー")
        self.setup_ui()
        self.current_index = 1
        self.worksheet = setup_google_auth()
        self.setup_drag()

    def setup_ui(self):
        index_label = customtkinter.CTkLabel(self, text="指定した番目:")
        index_label.pack(pady=10)

        self.index_entry = customtkinter.CTkEntry(self, width=120)
        self.index_entry.pack(pady=10)

        copy_specified_button = customtkinter.CTkButton(
            self,
            text="JANコードをコピー",
            command=self.copy_specified_index,
            fg_color="#0078D7",
            text_color="white",
            hover_color="#0053A6"
        )
        copy_specified_button.pack(pady=10)

        next_button = customtkinter.CTkButton(
            self,
            text="次へ進む",
            command=self.next_jan_code,
            fg_color="#0078D7",
            text_color="white",
            hover_color="#0053A6"
        )
        next_button.pack(pady=10)

        previous_button = customtkinter.CTkButton(
            self,
            text="一つ戻る",
            command=self.previous_jan_code,
            fg_color="#0078D7",
            text_color="white",
            hover_color="#0053A6"
        )
        previous_button.pack(pady=10)

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

    def copy_jan_code(self):
        cell_address = f'E{self.current_index + 1}'
        jan_code = self.worksheet.acell(cell_address).value
        pyperclip.copy(jan_code)
        messagebox.showinfo("コピー完了", f"{self.current_index}番目のJANコード {jan_code} をコピーしました。")
        threading.Thread(target=self.wait_and_paste).start()

    def wait_and_paste(self):
        time.sleep(3.0)
        p.hotkey('ctrl', 'v')

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

# メインアプリケーションのクラス
class MainApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + os.path.basename(sys.argv[0]))
        self.geometry("800x600")
        self.worksheet = setup_google_auth()
        self.setup_ui()

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)
        
        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        buttons = [
            ("重複と項目抜けのチェック", lambda: self.check_and_count_jan_codes(), 1, 0, "#008000"),
            ("JANコードのコピー", self.open_jan_copy, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", self.paste_and_execute, 2, 1, "#008000"),
            ("データ抜けのチェック", lambda: self.execute_batch_and_skip_opening_file("Type1x.bat"), 2, 2, "#FF0000"),
            ("チェックシートを開く", lambda: self.open_file(r'C:\pysample_01\checksheet12.bat'), 3, 0, "#0078D4"),
            ("座標軸とコピー", self.open_copy_coord, 2, 0, "#FF0000"),
            ("商品情報入力シートを開く", lambda: self.open_file('syouhin_n.bat'), 3, 1, "#0078D4"),
            ("藤原産業を開く", lambda: self.open_file('fujiwarasanngyou.bat'), 3, 2, "#0078D4"),
            ("input.txtのチェック", self.check_input_file, 4, 0, "#0078D4"),
            ("サブフォーム廃番処理", self.open_subform, 4, 1, "#0078D4"),
            ("Type2x.bat実行とoutput.txt表示", lambda: self.execute_batch_and_open_file("Type2x.bat", "output.txt"), 4, 2, "#0078D4"),
            ("checkd01.txtを開く", lambda: self.open_file('checkd01.txt'), 5, 1, "#0078D4"),
            ("checkd02.txtを開く", lambda: self.open_file('checkd02.txt'), 5, 2, "#0078D4"),
            ("クリップボードのクリア", self.clear_files, 5, 0, "#0078D4"),
            ("input.txtを開く", lambda: self.open_file('input.txt'), 1, 2, "#FF0000"),
        ]

        for text, command, row, col, color in buttons:
            btn = customtkinter.CTkButton(self.frame, text=text, command=command, fg_color=color)
            btn.grid(row=row, column=col, padx=5, pady=2, sticky="ew")

        for i in range(4):
            self.frame.columnconfigure(i, weight=1)

    def open_jan_copy(self):
        JanCopyApp()

    def open_copy_coord(self):
        CopyCoordApp()

    def open_subform(self):
        subprocess.Popen(['python', 'ckt0412sab01.py'])

    def check_and_count_jan_codes(self):
        try:
            with open('input.txt', 'r', encoding='utf-8') as f:
                data_str = f.read()
            jan_codes = re.findall(r'JANコード\t(\d+)', data_str)
            discontinued_codes = [
                code for code in jan_codes 
                if 'ブランド名\t廃番' in data_str.split(f'JANコード\t{code}')[1].split('JANコード')[0]
            ]
            duplicate_jan_codes = [
                code for code, count in Counter(jan_codes).items() 
                if count > 1
            ]
            
            self.output.delete("1.0", customtkinter.END)
            if duplicate_jan_codes:
                for duplicate in duplicate_jan_codes:
                    self.output.insert(customtkinter.END, f"JANコード {duplicate} が重複しています\n")
            else:
                self.output.insert(customtkinter.END, "重複はありません\n")
            
            for code in discontinued_codes:
                self.output.insert(customtkinter.END, f"JANコード {code} は廃番です\n")
            
            self.output.insert(customtkinter.END, f"JANコードは上から{len(jan_codes)}番目です\n")
            if jan_codes:
                self.output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました: {str(e)}")

    def paste_and_execute(self):
        clipboard_data = pyperclip.paste()
        if not clipboard_data:
            messagebox.showwarning("警告", "クリップボードにデータがありません。")
            return
        
        try:
            lines = clipboard_data.split('\n')
            output_data = "\n".join(line.strip() for line in lines if line.strip())
            
            with open('input.txt', 'a', encoding='utf-8') as file:
                file.write(output_data + '\n\n')
            
            self.output.insert(customtkinter.END, clipboard_data + "\n")
            print('新しいデータがinput.txtに追加されました。')
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました: {str(e)}")

    def execute_batch_and_skip_opening_file(self, batch_file):
        try:
            subprocess.run([batch_file])
            time.sleep(1)
            self.check_empty_checkd01()
            
	def check_empty_checkd01(self):
        try:
            with open('checkd01.txt', 'r', encoding='utf-8') as file:
                content = file.read().strip()
                if not content:
                    messagebox.showinfo("情報", "checkd01.txtは空です")
        except FileNotFoundError:
            messagebox.showwarning("ファイルが見つかりません", "checkd01.txt が見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

    def execute_batch_and_open_file(self, batch_file, output_file):
        try:
            subprocess.run([batch_file])
            time.sleep(1)
            self.open_file(output_file)
            self.check_data_differences()
        except Exception as e:
            messagebox.showerror("エラー", f"バッチ処理中にエラーが発生しました: {str(e)}")

    def open_file(self, file_path):
        try:
            os.startfile(file_path)
        except FileNotFoundError:
            messagebox.showwarning("ファイルが見つかりません", f"{file_path}が見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルを開く際にエラーが発生しました： {str(e)}")

    def clear_files(self):
        if messagebox.askyesno("確認", "本当にクリアして良いですか？"):
            try:
                files_to_clear = [
                    "checkd01.txt", 
                    "checkd02.txt", 
                    "output02.txt", 
                    "output.txt", 
                    "input.txt"
                ]
                for file in files_to_clear:
                    with open(file, 'w', encoding='utf-8') as f:
                        f.write('')
                messagebox.showinfo("クリア完了", "ファイルのデータをクリアしました。")
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルをクリアする際にエラーが発生しました： {str(e)}")

    def check_input_file(self):
        try:
            with open('input.txt', 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            invalid_an_code_lines = [
                i for i, line in enumerate(lines, 1) 
                if 'ANコード' in line and 'J' not in line
            ]
            
            thirteen_digit_lines = [
                i for i, line in enumerate(lines, 1) 
                if re.match(r'^\d{13}$', line.strip())
            ]
            
            message = (
                f"「ANコード」に「J」が含まれていない行: {', '.join(map(str, invalid_an_code_lines)) or 'なし'}\n"
                f"13桁の数値のみを含む行: {', '.join(map(str, thirteen_digit_lines)) or 'なし'}"
            )
            messagebox.showinfo("チェック結果", message)
        except FileNotFoundError:
            messagebox.showerror("エラー", "input.txtファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルのチェック中にエラーが発生しました： {str(e)}")

    def check_data_differences(self):
        try:
            with open('checkd01.txt', 'r', encoding='utf-8') as file1, \
                 open('checkd02.txt', 'r', encoding='utf-8') as file2:
                lines1 = file1.readlines()
                lines2 = file2.readlines()
                
                differences = []
                for i, (line1, line2) in enumerate(zip(lines1, lines2)):
                    jan_code1 = re.search(r'\d+', line1)
                    jan_code2 = re.search(r'\d+', line2)
                    if jan_code1 and jan_code2 and jan_code1.group() != jan_code2.group():
                        differences.append(f"{i + 1}番目が違います")
                
                self.output.delete("1.0", customtkinter.END)
                if differences:
                    self.output.insert(customtkinter.END, "\n".join(differences))
                else:
                    self.output.insert(customtkinter.END, "すべてのJANコードが一致しています")
        except FileNotFoundError:
            messagebox.showwarning(
                "ファイルが見つかりません", 
                "checkd01.txt または checkd02.txt が見つかりません。"
            )
        except Exception as e:
            messagebox.showerror(
                "エラー", 
                f"操作中にエラーが発生しました： {str(e)}"
            )

# メインエントリーポイント
def main():
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    app = MainApp()
    app.mainloop()

if __name__ == '__main__':
    main()





