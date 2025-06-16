import customtkinter as ctk
import pyautogui
import time
import pyperclip

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class CopyCoordApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("座標取得アプリ")
        self.geometry("220x150")
        self.start_button = ctk.CTkButton(self, text="開始位置", command=self.capture_start, fg_color="#0078D4", corner_radius=10, width=180, height=40)
        self.start_button.pack(pady=10, padx=20)
        
        self.end_button = ctk.CTkButton(self, text="終了位置", command=self.capture_end, state="disabled", fg_color="#0078D4", corner_radius=10, width=180, height=40)
        self.end_button.pack(pady=10, padx=20)
        self.start_pos = None
        self.end_pos = None
        self.total_start_time = None
        self.copied_content = None

    def capture_start(self):
        self.total_start_time = time.time()
        self.start_pos = self.capture_position("開始位置")
        if self.start_pos:
            self.end_button.configure(state="normal")

    def capture_end(self):
        self.end_pos = self.capture_position("終了位置")
        if self.end_pos:
            self.after(1000, self.perform_drag_and_copy)  # 1.0秒に変更

    def capture_position(self, position_type):
        self.withdraw()
        time.sleep(1.5)  # 待ち時間を1.5秒に設定
        try:
            # 1回目の座標取得
            x1, y1 = pyautogui.position()
            time.sleep(0.5)  # 0.5秒後にもう一度取得
            # 2回目の座標取得
            x2, y2 = pyautogui.position()
            # 座標の平均値を計算
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
            time.sleep(0.4)  # コピーの完了を待つ
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

def get_coordinates_and_content():
    try:
        with open('coordinates_and_content.txt', 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return "座標とコピー内容が保存されていません。"

if __name__ == '__main__':
    app = CopyCoordApp()
    app.mainloop()
