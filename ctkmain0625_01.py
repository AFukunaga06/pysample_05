#sample
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
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException

spreadsheet_id = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'
KEY_FILE_PATH = r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json'

creds = Credentials.from_service_account_file(KEY_FILE_PATH, scopes=[
    'https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'
])
client = gspread.authorize(creds)
worksheet = client.open_by_key(spreadsheet_id).sheet1
file_name = os.path.basename(sys.argv[0])

# 自動検索用のグローバル変数
auto_search_driver = None
auto_search_thread = None
auto_search_running = False

if not os.path.exists('input.txt'):
    with open('input.txt', 'w', encoding='utf-8') as f:
        pass

def show_under_construction():
    messagebox.showinfo("お知らせ", "この機能は現在工事中です。\nしばらくお待ちください。")

def safe_file_operation(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{str(e)}ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")
    return wrapper

def setup_auto_search_driver():
    """自動検索用のWebDriverをセットアップ"""
    global auto_search_driver
    try:
        chrome_options = Options()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        auto_search_driver = webdriver.Chrome(options=chrome_options)
        auto_search_driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # 商品検索ページを開く（実際のURLに変更してください）
        # auto_search_driver.get("https://your-product-search-site.com")
        
        return True
    except Exception as e:
        messagebox.showerror("エラー", f"WebDriverの初期化に失敗しました： {str(e)}")
        return False

def monitor_search_box():
    """商品検索ボックスを監視し、JANコードが入力されたら自動検索"""
    global auto_search_driver, auto_search_running
    
    if not auto_search_driver:
        return
    
    last_value = ""
    
    while auto_search_running:
        try:
            # 商品検索ボックスの要素を取得（実際のセレクタに変更してください）
            search_box = auto_search_driver.find_element(By.CSS_SELECTOR, "input[type='text']")
            current_value = search_box.get_attribute("value")
            
            # 値が変更され、JANコードの形式（数字のみ、8-14桁）の場合
            if current_value != last_value and current_value.strip():
                if re.match(r'^\d{8,14}$', current_value.strip()):
                    print(f"JANコード検出: {current_value}")
                    
                    # 検索ボタンをクリック（実際のセレクタに変更してください）
                    try:
                        search_button = auto_search_driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit'], .search-button")
                        search_button.click()
                        print("自動検索実行")
                        
                        # ウィンドウのアウトプットに結果を表示
                        if hasattr(window, 'output'):
                            window.output.insert(customtkinter.END, f"自動検索実行: {current_value}\n")
                        
                    except NoSuchElementException:
                        print("検索ボタンが見つかりません")
                
                last_value = current_value
            
            time.sleep(0.5)  # 0.5秒間隔で監視
            
        except NoSuchElementException:
            # 検索ボックスが見つからない場合は少し待つ
            time.sleep(1)
        except Exception as e:
            print(f"監視中にエラー: {str(e)}")
            time.sleep(1)

def start_auto_search():
    """自動検索機能を開始"""
    global auto_search_thread, auto_search_running
    
    if auto_search_running:
        messagebox.showinfo("情報", "自動検索機能は既に開始されています。")
        return
    
    if setup_auto_search_driver():
        auto_search_running = True
        auto_search_thread = threading.Thread(target=monitor_search_box, daemon=True)
        auto_search_thread.start()
        messagebox.showinfo("開始", "自動検索機能を開始しました。")
        
        # ボタンの状態を更新
        if hasattr(window, 'auto_search_btn'):
            window.auto_search_btn.configure(text="自動検索停止", fg_color="#FF0000")

def stop_auto_search():
    """自動検索機能を停止"""
    global auto_search_driver, auto_search_running
    
    auto_search_running = False
    
    if auto_search_driver:
        try:
            auto_search_driver.quit()
        except:
            pass
        auto_search_driver = None
    
    messagebox.showinfo("停止", "自動検索機能を停止しました。")
    
    # ボタンの状態を更新
    if hasattr(window, 'auto_search_btn'):
        window.auto_search_btn.configure(text="自動検索開始", fg_color="#008000")

def toggle_auto_search():
    """自動検索機能の開始/停止を切り替え"""
    if auto_search_running:
        stop_auto_search()
    else:
        start_auto_search()

@safe_file_operation
def process_clipboard_data(jancode, clipboard_data):
    lines = clipboard_data.split('\n')
    output_data = "\n".join(line.strip() for line in lines if line.strip())
    if jancode:
        output_data = f"JANコード\t{jancode}\n{output_data}"
    with open('input.txt', 'a', encoding='utf-8') as file:
        file.write(output_data + '\n\n')
    print('新しいデータがinput.txtに追加されました。')

@safe_file_operation
def check_and_count_jan_codes(output):
    with open('input.txt', 'r', encoding='utf-8') as f:
        data_str = f.read()
    jan_codes = re.findall(r'JANコード\t(\d+)', data_str)
    discontinued_codes = [code for code in jan_codes if 'ブランド名\t廃番' in data_str.split(f'JANコード\t{code}')[1].split('JANコード')[0]]
    duplicate_jan_codes = [code for code, count in Counter(jan_codes).items() if count > 1]

    output.delete("1.0", customtkinter.END)

    if duplicate_jan_codes:
        for duplicate in duplicate_jan_codes:
            output.insert(customtkinter.END, f"JANコード {duplicate} が重複しています\n")
    else:
        output.insert(customtkinter.END, "重複はありません\n")

    for code in discontinued_codes:
        output.insert(customtkinter.END, f"JANコード {code} は廃番です\n")

    output.insert(customtkinter.END, f"JANコードは上から{len(jan_codes) + 1}番目です\n", "red_text")
    if jan_codes:
        output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")

    output.tag_config("red_text", foreground="red")

def paste_and_execute():
    clipboard_data = pyperclip.paste()
    if not clipboard_data:
        messagebox.showwarning("警告", "クリップボードにデータがありません。")
        return
    process_clipboard_data("", clipboard_data)
    window.output.insert(customtkinter.END, clipboard_data + "\n")

@safe_file_operation
def open_file(file_path):
    if not os.path.exists(file_path):
        messagebox.showerror("エラー", f"{file_path}が存在しません。")
        return
    os.startfile(file_path)

def execute_batch_and_open_file(batch_file, output_file):
    @safe_file_operation
    def execute():
        subprocess.run([batch_file])
        time.sleep(1)
        open_file(output_file)
        check_data_differences(window.output)
    execute()

def check_empty_checkd01():
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as file:
            content = file.read().strip()
            if not content:
                messagebox.showinfo("情報", "checkd01.txtは空です")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

def execute_batch_and_skip_opening_file(batch_file):
    @safe_file_operation
    def execute():
        subprocess.run([batch_file])
        time.sleep(1)
        check_empty_checkd01()
        check_data_differences(window.output)
    execute()

def clear_files():
    if messagebox.askyesno("確認", "本当にクリアして良いですか？"):
        try:
            for file in ["checkd01.txt", "checkd02.txt", "output02.txt", "output.txt", "input.txt"]:
                if os.path.exists(file):
                    with open(file, 'w', encoding='utf-8') as f:
                        f.write('')
            messagebox.showinfo("クリア完了", "ファイルのデータをクリアしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルをクリアする際にエラーが発生しました： {str(e)}")

@safe_file_operation
def check_data_differences(output):
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as file1, open('checkd02.txt', 'r', encoding='utf-8') as file2:
            lines1 = file1.readlines()
            lines2 = file2.readlines()
            differences = []
            for i, (line1, line2) in enumerate(zip(lines1, lines2)):
                jan_code1 = re.search(r'\d+', line1)
                jan_code2 = re.search(r'\d+', line2)
                if jan_code1 and jan_code2 and jan_code1.group() != jan_code2.group():
                    differences.append(f"{i + 1}番目が違います")
            output.delete("1.0", customtkinter.END)
            if differences:
                output.insert(customtkinter.END, "\n".join(differences))
            else:
                output.insert(customtkinter.END, "すべてのJANコードが一致しています")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt または checkd02.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

@safe_file_operation
def get_data_differences_result():
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as file1, open('checkd02.txt', 'r', encoding='utf-8') as file2:
            lines1 = file1.readlines()
            lines2 = file2.readlines()
            differences = []
            for i, (line1, line2) in enumerate(zip(lines1, lines2)):
                jan_code1 = re.search(r'\d+', line1)
                jan_code2 = re.search(r'\d+', line2)
                if jan_code1 and jan_code2 and jan_code1.group() != jan_code2.group():
                    differences.append(f"{i + 1}番目が違います")
            if differences:
                return "\n".join(differences)
            else:
                return "すべてのデータに問題はありません"
    except FileNotFoundError:
        return "checkd01.txt または checkd02.txt が見つかりません"
    except Exception as e:
        return f"データ抜けチェック中にエラーが発生しました： {str(e)}"

@safe_file_operation
def check_duplicate_and_data_gaps(output):
    with open('input.txt', 'r', encoding='utf-8') as f:
        data_str = f.read()
    jan_codes = re.findall(r'JANコード\t(\d+)', data_str)
    discontinued_codes = [code for code in jan_codes if 'ブランド名\t廃番' in data_str.split(f'JANコード\t{code}')[1].split('JANコード')[0]]
    duplicate_jan_codes = [code for code, count in Counter(jan_codes).items() if count > 1]

    output.delete("1.0", customtkinter.END)

    if duplicate_jan_codes:
        for duplicate in duplicate_jan_codes:
            output.insert(customtkinter.END, f"JANコード {duplicate} が重複しています\n")
    else:
        output.insert(customtkinter.END, "重複はありません\n")

    for code in discontinued_codes:
        output.insert(customtkinter.END, f"JANコード {code} は廃番です\n")

    output.insert(customtkinter.END, f"JANコードは上から{len(jan_codes) + 1}番目です\n", "red_text")
    if jan_codes:
        output.insert(customtkinter.END, f"現在のJANコードは{jan_codes[-1]}です\n")

    output.insert(customtkinter.END, "\nデータ抜けのチェック結果：\n")
    data_gap_result = get_data_differences_result()
    output.insert(customtkinter.END, data_gap_result)

    output.tag_config("red_text", foreground="red")

class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Application - " + file_name)
        window_width = 800
        window_height = 650  # 高さを少し増やす
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.resizable(False, False)
        self.setup_ui()
        
        # ウィンドウ終了時の処理
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """ウィンドウ終了時に自動検索機能を停止"""
        global auto_search_running, auto_search_driver
        auto_search_running = False
        if auto_search_driver:
            try:
                auto_search_driver.quit()
            except:
                pass
        self.destroy()

    def combined_check(self):
        execute_batch_and_skip_opening_file("Type1x.bat")
        check_duplicate_and_data_gaps(self.output)

    def setup_ui(self):
        self.frame = customtkinter.CTkFrame(self)
        self.frame.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)

        self.output = customtkinter.CTkTextbox(self.frame, width=760, height=200)
        self.output.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        buttons = [
            ("重複と項目抜けのチェック", self.combined_check, 1, 0, "#008000"),
            ("JANコード等のコピー", self.jancode_copy, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", paste_and_execute, 2, 1, "#008000"),
            ("データ抜けのチェック", lambda: execute_batch_and_skip_opening_file("Type1x.bat"), 2, 2, "#FF0000"),
            ("チェックシートを開く", lambda: open_file(r'C:\pysample_01\checksheet12.bat'), 3, 0, "#0078D4"),
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

        # 自動検索ボタンを追加
        self.auto_search_btn = customtkinter.CTkButton(
            self.frame, 
            text="自動検索開始", 
            command=toggle_auto_search, 
            fg_color="#008000"
        )
        self.auto_search_btn.grid(row=6, column=0, padx=5, pady=2, sticky="ew")

        for i in range(4):
            self.frame.columnconfigure(i, weight=1)

        self.info_label = customtkinter.CTkLabel(
            self.frame,
            text="※商品情報入力シートに必ず名前を明示すること",
            text_color="red",
            font=("メイリオ", 18, "bold")
        )
        self.info_label.grid(row=7, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        self.info_label2 = customtkinter.CTkLabel(
            self.frame,
            text="※チェックシートの内容をcheckd01.txtにコピーする",
            text_color="red",
            font=("メイリオ", 18, "bold")
        )
        self.info_label2.grid(row=8, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        # 自動検索機能の説明ラベル
        self.info_label3 = customtkinter.CTkLabel(
            self.frame,
            text="※自動検索：JANコード入力時に自動で検索ボタンをクリック",
            text_color="blue",
            font=("メイリオ", 14, "bold")
        )
        self.info_label3.grid(row=9, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

    @safe_file_operation
    def jancode_copy(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 220
        window_height = 500
        x_position = (screen_width - window_width) // 2 - 400 - 8
        y_position = (screen_height - window_height) // 2
        env = os.environ.copy()
        env['WINDOW_POSITION'] = f"{x_position},{y_position}"
        env['PARENT_WINDOW'] = f"{self.winfo_x()},{self.winfo_y()}"
        subprocess.Popen(['python', 'jancoordcopy1122_01.py'], env=env)

    @safe_file_operation
    def open_subform(self):
        subprocess.Popen(['python', 'ckt0412sab01.py'])

    def check_input_file(self):
        try:
            with open('input.txt', 'r', encoding='utf-8') as f:
                lines = f.readlines()
            invalid_an_code_lines = [i for i, line in enumerate(lines, 1) if 'ANコード' in line and 'J' not in line]
            thirteen_digit_lines = [i for i, line in enumerate(lines, 1) if re.match(r'^\d{13}$', line.strip())]
            message = (
                f"「ANコード」に「J」が含まれていない行: {', '.join(map(str, invalid_an_code_lines)) or 'なし'}\n"
                f"13桁の数値のみを含む行: {', '.join(map(str, thirteen_digit_lines)) or 'なし'}"
            )
            messagebox.showinfo("チェック結果", message)
        except FileNotFoundError:
            messagebox.showerror("エラー", "input.txtファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルのチェック中にエラーが発生しました： {str(e)}")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window = MainWindow()
    window.mainloop()