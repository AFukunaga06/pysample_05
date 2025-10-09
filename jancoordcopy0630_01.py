import customtkinter as ctk
import pyautogui
import time
import pyperclip

# 左クリックで開始位置を取得し、終了位置までドラッグして選択したテキストをファイルに保存するアプリ

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class CopyCoordApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("テキスト選択コピーアプリ")
        self.geometry("220x150")  # ウィンドウサイズを固定

        # 開始位置ボタン
        self.start_button = ctk.CTkButton(
            self,
            text="開始位置をクリック",
            command=self.capture_start,
            fg_color="#0078D4",
            corner_radius=10,
            width=180,
            height=40
        )
        self.start_button.pack(pady=10, padx=20)

        # 終了位置ボタン（最初は無効）
        self.end_button = ctk.CTkButton(
            self,
            text="終了位置をクリック",
            command=self.capture_end,
            state="disabled",
            fg_color="#0078D4",
            corner_radius=10,
            width=180,
            height=40
        )
        self.end_button.pack(pady=10, padx=20)

        self.start_pos = None
        self.end_pos = None
        self.total_start_time = None
        self.copied_content = None

    def capture_start(self):
        # 開始位置を記録し、終了ボタンを有効化
        self.total_start_time = time.time()
        self.start_pos = self.capture_position("開始位置")
        if self.start_pos:
            self.end_button.configure(state="normal")
            self.start_button.configure(state="disabled")

    def capture_end(self):
        # 終了位置を記録し、ドラッグ＆コピーを実行
        self.end_pos = self.capture_position("終了位置")
        if self.end_pos:
            # 少し待ってからドラッグを実行
            self.after(1000, self.perform_drag_and_copy)

    def capture_position(self, position_type):
        # 一度ウィンドウを隠してカーソル位置をキャプチャ
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
            print(f"座標取得中にエラーが発生しました: {e}")
            return None
        finally:
            self.deiconify()

    def perform_drag_and_copy(self):
        try:
            # 開始位置から終了位置までドラッグ
            pyautogui.moveTo(self.start_pos[0], self.start_pos[1])
            pyautogui.mouseDown()
            pyautogui.moveTo(self.end_pos[0], self.end_pos[1], duration=0.5)
            pyautogui.mouseUp()
            # コピー実行
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.4)
            self.copied_content = pyperclip.paste()
            elapsed = time.time() - self.total_start_time
            print(f"選択とコピーに要した時間: {elapsed:.2f}秒")

            # テキストファイルに保存
            self.save_selected_content()
        except Exception as e:
            print(f"ドラッグ＆コピー中にエラーが発生しました: {e}")
        finally:
            # 状態をリセット
            self.end_button.configure(state="disabled")
            self.start_button.configure(state="normal")
            self.start_pos = None
            self.end_pos = None

    def save_selected_content(self):
        # 選択したテキストをファイルに書き込み
        try:
            with open('selected_text.txt', 'w', encoding='utf-8') as f:
                f.write(self.copied_content or '')
            print('選択したテキストをselected_text.txtに保存しました。')
        except Exception as e:
            print(f"ファイル保存中にエラーが発生しました: {e}")

if __name__ == '__main__':
    app = CopyCoordApp()
    app.mainloop()
