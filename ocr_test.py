# OCR実行機能
        ocr_frame = ctk.CTkFrame(self)
        ocr_frame.pack(pady=5, padx=10, fill="x")import customtkinter as ctk
from tkinter import messagebox, filedialog
import pyperclip
import time
import os
from PIL import Image, ImageTk
import tkinter as tk

# OCR関連ライブラリの動的インポート
try:
    import pytesseract
    import cv2
    import numpy as np
    OCR_AVAILABLE = True
    print("OCR機能が利用可能です")
    
    # Windows環境でTesseract-OCRのパスを設定
    if os.path.exists(r'C:\Program Files\Tesseract-OCR\tesseract.exe'):
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        print("Tesseract-OCRパスを設定しました")
    
except ImportError as e:
    OCR_AVAILABLE = False
    print(f"OCR機能が利用できません: {e}")
    print("必要なライブラリ: pip install pytesseract opencv-python")

# pyautoguiの動的インポート
try:
    import pyautogui
    SCREENSHOT_AVAILABLE = True
    print("スクリーンショット機能が利用可能です")
except ImportError:
    SCREENSHOT_AVAILABLE = False
    print("スクリーンショット機能が利用できません")
    print("必要なライブラリ: pip install pyautogui")

class SimpleOCRApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("OCR機能テスト - JANコード商品情報抽出")
        
        # ウィンドウサイズを設定
        window_width = 400
        window_height = 600
        
        # 中央配置
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_pos = (screen_width - window_width) // 2
        y_pos = (screen_height - window_height) // 2
        
        self.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
        self.resizable(False, False)
        
        # 変数の初期化
        self.screenshot = None
        self.selection_start = None
        self.selection_end = None
        self.left_monitor_region = None  # 左モニターの範囲を保存
        
        self.setup_ui()
        self.check_dependencies()

    def check_dependencies(self):
        """依存関係をチェックして状態を表示"""
        status_text = "システム状態:\n"
        
        if OCR_AVAILABLE:
            status_text += "✓ OCR機能: 利用可能\n"
        else:
            status_text += "✗ OCR機能: 利用不可\n"
            
        if SCREENSHOT_AVAILABLE:
            status_text += "✓ スクリーンショット: 利用可能\n"
        else:
            status_text += "✗ スクリーンショット: 利用不可\n"
            
        # Tesseract確認
        if OCR_AVAILABLE:
            try:
                # テスト用の小さな画像でTesseractをテスト
                test_img = Image.new('RGB', (100, 50), color='white')
                pytesseract.image_to_string(test_img)
                status_text += "✓ Tesseract-OCR: 正常動作\n"
                
                # 日本語サポート確認
                try:
                    langs = pytesseract.get_languages()
                    if 'jpn' in langs:
                        status_text += "✓ 日本語サポート: 利用可能\n"
                    else:
                        status_text += "⚠ 日本語サポート: 未確認\n"
                except:
                    status_text += "⚠ 日本語サポート: 未確認\n"
                    
            except Exception as e:
                status_text += f"✗ Tesseract-OCR: エラー\n({str(e)[:30]}...)\n"
        
        self.status_label.configure(text=status_text)

    def setup_ui(self):
        # タイトル
        title_label = ctk.CTkLabel(self, text="OCR機能テスト", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # システム状態表示
        self.status_label = ctk.CTkLabel(
            self, 
            text="システム状態を確認中...", 
            font=("Arial", 10),
            justify="left"
        )
        self.status_label.pack(pady=5)

        # 区切り線
        separator1 = ctk.CTkFrame(self, height=2)
        separator1.pack(fill="x", padx=20, pady=10)

        # 使用方法の説明
        usage_frame = ctk.CTkFrame(self)
        usage_frame.pack(pady=5, padx=10, fill="x")
        
        usage_label = ctk.CTkLabel(usage_frame, text="使用方法", font=("Arial", 12, "bold"))
        usage_label.pack(pady=5)
        
        usage_text = ctk.CTkLabel(
            usage_frame, 
            text="1. スクリーンショット撮影\n2. 範囲選択＆OCR実行\n3. JANコード右側から商品サイズまでドラッグ", 
            font=("Arial", 9),
            justify="left"
        )
        usage_text.pack(pady=5)

        # スクリーンショット機能
        screenshot_frame = ctk.CTkFrame(self)
        screenshot_frame.pack(pady=5, padx=10, fill="x")
        
        screenshot_label = ctk.CTkLabel(screenshot_frame, text="スクリーンショット機能", font=("Arial", 12, "bold"))
        screenshot_label.pack(pady=5)

        screenshot_button = ctk.CTkButton(
            screenshot_frame,
            text="スクリーンショット撮影",
            command=self.take_screenshot,
            width=250,
            height=35
        )
        screenshot_button.pack(pady=5)

        # 画像読み込み機能
        load_frame = ctk.CTkFrame(self)
        load_frame.pack(pady=5, padx=10, fill="x")
        
        load_label = ctk.CTkLabel(load_frame, text="画像読み込み機能", font=("Arial", 12, "bold"))
        load_label.pack(pady=5)

        load_button = ctk.CTkButton(
            load_frame,
            text="画像ファイルを読み込み",
            command=self.load_image_file,
            width=250,
            height=35
        )
        load_button.pack(pady=5)

        # 自動認識機能
        auto_frame = ctk.CTkFrame(self)
        auto_frame.pack(pady=5, padx=10, fill="x")
        
        auto_label = ctk.CTkLabel(auto_frame, text="自動認識機能", font=("Arial", 12, "bold"))
        auto_label.pack(pady=5)

        detect_button = ctk.CTkButton(
            auto_frame,
            text="左モニターのJANコード範囲を自動検出",
            command=self.auto_detect_left_monitor,
            width=250,
            height=35,
            fg_color="#9C27B0"
        )
        detect_button.pack(pady=2)

        manual_region_button = ctk.CTkButton(
            auto_frame,
            text="手動で範囲座標を設定",
            command=self.set_manual_region,
            width=250,
            height=35,
            fg_color="#607D8B"
        )
        manual_region_button.pack(pady=2)
        
        ocr_label = ctk.CTkLabel(ocr_frame, text="OCR実行機能", font=("Arial", 12, "bold"))
        ocr_label.pack(pady=5)

        ocr_button = ctk.CTkButton(
            ocr_frame,
            text="範囲選択＆OCR実行",
            command=self.select_area_and_ocr,
            width=250,
            height=35,
            fg_color="#FF6B35"
        )
        ocr_button.pack(pady=5)

        # 結果表示エリア
        result_frame = ctk.CTkFrame(self)
        result_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        result_label = ctk.CTkLabel(result_frame, text="OCR結果", font=("Arial", 12, "bold"))
        result_label.pack(pady=5)

        self.result_text = ctk.CTkTextbox(result_frame, height=120)
        self.result_text.pack(pady=5, padx=5, fill="both", expand=True)

        # ボタン群
        button_frame1 = ctk.CTkFrame(result_frame)
        button_frame1.pack(pady=5, fill="x")

        copy_button = ctk.CTkButton(
            button_frame1,
            text="クリップボードにコピー",
            command=self.copy_result,
            width=180
        )
        copy_button.pack(side="left", padx=2)

        save_button = ctk.CTkButton(
            button_frame1,
            text="ファイルに保存",
            command=self.save_to_file,
            width=180
        )
        save_button.pack(side="right", padx=2)

        # 第2ボタン行
        button_frame2 = ctk.CTkFrame(result_frame)
        button_frame2.pack(pady=2, fill="x")

        input_button = ctk.CTkButton(
            button_frame2,
            text="input.txtに追記",
            command=self.append_to_input,
            width=180,
            fg_color="#28A745"
        )
        input_button.pack(side="left", padx=2)

        clear_button = ctk.CTkButton(
            button_frame2,
            text="結果をクリア",
            command=self.clear_result,
            width=180,
            fg_color="#DC3545"
        )
        clear_button.pack(side="right", padx=2)

    def take_screenshot(self):
        """スクリーンショットを撮影（マルチディスプレイ対応）"""
        if not SCREENSHOT_AVAILABLE:
            messagebox.showerror("エラー", "pyautoguiがインストールされていません。\npip install pyautogui")
            return
            
        # モニター選択ダイアログを表示
        choice = messagebox.askyesnocancel(
            "モニター選択", 
            "撮影するモニターを選択してください:\n\n" +
            "「はい」: 全モニター撮影\n" +
            "「いいえ」: プライマリモニターのみ\n" +
            "「キャンセル」: 撮影中止"
        )
        
        if choice is None:  # キャンセル
            return
            
        try:
            # カウントダウンを表示
            for i in range(3, 0, -1):
                self.title(f"OCRテスト - {i}秒後に撮影...")
                self.update()
                time.sleep(1)
            
            # ウィンドウを最小化
            self.iconify()
            time.sleep(0.5)
            
            if choice:  # 全モニター撮影
                self.screenshot = pyautogui.screenshot()
            else:  # プライマリモニターのみ
                # プライマリモニターのサイズを取得
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()
                width = root.winfo_screenwidth()
                height = root.winfo_screenheight()
                root.destroy()
                
                # プライマリモニターのみ撮影
                self.screenshot = pyautogui.screenshot(region=(0, 0, width, height))
            
            # ウィンドウを復元
            self.deiconify()
            self.title("OCR機能テスト - JANコード商品情報抽出")
            
            monitor_type = "全モニター" if choice else "プライマリモニター"
            messagebox.showinfo("完了", f"{monitor_type}のスクリーンショットを撮影しました。\n「範囲選択＆OCR実行」を押してください。")
            
        except Exception as e:
            messagebox.showerror("エラー", f"スクリーンショット撮影エラー:\n{str(e)}")
            self.deiconify()
            self.title("OCR機能テスト - JANコード商品情報抽出")

    def load_image_file(self):
        """画像ファイルを読み込み"""
        try:
            file_path = filedialog.askopenfilename(
                title="画像ファイルを選択",
                filetypes=[
                    ("画像ファイル", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                    ("すべてのファイル", "*.*")
                ]
            )
            
            if file_path:
                self.screenshot = Image.open(file_path)
                messagebox.showinfo("完了", f"画像ファイルを読み込みました。\nファイル: {os.path.basename(file_path)}")
                
        except Exception as e:
            messagebox.showerror("エラー", f"画像読み込みエラー:\n{str(e)}")

    def select_area_and_ocr(self):
        """範囲選択とOCR実行"""
        if not OCR_AVAILABLE:
            messagebox.showerror("エラー", "OCR機能が利用できません。\n必要なライブラリをインストールしてください。")
            return
            
        if self.screenshot is None:
            messagebox.showwarning("警告", "画像が読み込まれていません。\nスクリーンショットを撮影するか、画像ファイルを読み込んでください。")
            return
            
        # 範囲選択ウィンドウを開く
        self.open_selection_window()

    def open_selection_window(self):
        """範囲選択用のウィンドウを開く"""
        self.selection_window = tk.Toplevel(self)
        self.selection_window.title("範囲選択 - JANコード右側から商品サイズまでドラッグしてください")
        self.selection_window.attributes('-topmost', True)
        
        # 画面サイズに合わせてリサイズ
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 画像をリサイズ（最大80%のサイズ）
        max_width = int(screen_width * 0.8)
        max_height = int(screen_height * 0.8)
        
        img_ratio = min(max_width / self.screenshot.width, max_height / self.screenshot.height)
        new_width = int(self.screenshot.width * img_ratio)
        new_height = int(self.screenshot.height * img_ratio)
        
        self.display_image = self.screenshot.resize((new_width, new_height), Image.Resampling.LANCZOS)
        self.photo = ImageTk.PhotoImage(self.display_image)
        
        # Canvasを作成
        self.canvas = tk.Canvas(self.selection_window, width=new_width, height=new_height)
        self.canvas.pack()
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
        
        # 説明ラベル
        instruction_label = tk.Label(
            self.selection_window, 
            text="JANコードの右側から商品サイズまでドラッグして選択してください",
            font=("Arial", 10, "bold"),
            bg="yellow"
        )
        instruction_label.pack(pady=5)
        
        # マウスイベントをバインド
        self.canvas.bind("<Button-1>", self.on_selection_start)
        self.canvas.bind("<B1-Motion>", self.on_selection_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_selection_end)
        
        # 縮小比率を保存
        self.scale_ratio = img_ratio
        
        # 選択範囲の初期化
        self.selection_start = None
        self.selection_end = None
        self.selection_rect = None

    def on_selection_start(self, event):
        """選択開始"""
        self.selection_start = (event.x, event.y)
        if self.selection_rect:
            self.canvas.delete(self.selection_rect)

    def on_selection_drag(self, event):
        """選択中のドラッグ"""
        if self.selection_start:
            if self.selection_rect:
                self.canvas.delete(self.selection_rect)
            self.selection_rect = self.canvas.create_rectangle(
                self.selection_start[0], self.selection_start[1],
                event.x, event.y,
                outline="red", width=3
            )

    def on_selection_end(self, event):
        """選択終了"""
        if self.selection_start:
            self.selection_end = (event.x, event.y)
            
            # 選択範囲をOCR処理
            self.perform_ocr()
            
            # 選択ウィンドウを閉じる
            self.selection_window.destroy()

    def perform_ocr(self):
        """OCR処理を実行"""
        try:
            if not self.selection_start or not self.selection_end:
                messagebox.showwarning("警告", "範囲が選択されていません。")
                return
            
            # 選択範囲を元の画像サイズに変換
            x1 = int(min(self.selection_start[0], self.selection_end[0]) / self.scale_ratio)
            y1 = int(min(self.selection_start[1], self.selection_end[1]) / self.scale_ratio)
            x2 = int(max(self.selection_start[0], self.selection_end[0]) / self.scale_ratio)
            y2 = int(max(self.selection_start[1], self.selection_end[1]) / self.scale_ratio)
            
            # 選択範囲を切り出し
            cropped_image = self.screenshot.crop((x1, y1, x2, y2))
            
            # OCR処理
            # PIL ImageをOpenCV形式に変換
            opencv_image = cv2.cvtColor(np.array(cropped_image), cv2.COLOR_RGB2BGR)
            
            # 画像の前処理（OCR精度向上のため）
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
            # ノイズ除去
            denoised = cv2.fastNlMeansDenoising(gray)
            
            # 二値化
            _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # OCR実行（日本語＋英語）
            try:
                # 日本語＋英語でOCR実行
                text = pytesseract.image_to_string(
                    binary, 
                    config=r'--oem 3 --psm 6 -l jpn+eng',
                    lang='jpn+eng'
                )
            except:
                try:
                    # 日本語のみで実行
                    text = pytesseract.image_to_string(
                        binary, 
                        config=r'--oem 3 --psm 6 -l jpn',
                        lang='jpn'
                    )
                except:
                    try:
                        # 英語のみで実行
                        text = pytesseract.image_to_string(binary, config=r'--oem 3 --psm 6')
                    except Exception as e:
                        messagebox.showerror("OCRエラー", f"OCR処理でエラーが発生しました:\n{str(e)}")
                        return
            
            # 結果を表示
            self.result_text.delete("1.0", "end")
            self.result_text.insert("1.0", text.strip())
            
            # 自動でクリップボードにコピー
            if text.strip():
                pyperclip.copy(text.strip())
                messagebox.showinfo("完了", f"OCR処理が完了しました。\n結果をクリップボードにコピーしました。\n\n文字数: {len(text.strip())}")
            else:
                messagebox.showwarning("警告", "テキストが検出されませんでした。\n画像の品質を確認してください。")
                
        except Exception as e:
            messagebox.showerror("エラー", f"OCR処理エラー:\n{str(e)}")

    def copy_result(self):
        """結果をクリップボードにコピー"""
        text = self.result_text.get("1.0", "end-1c")
        if text.strip():
            pyperclip.copy(text.strip())
            messagebox.showinfo("完了", "結果をクリップボードにコピーしました。")
        else:
            messagebox.showwarning("警告", "コピーするテキストがありません。")

    def save_to_file(self):
        """結果をテキストファイルに保存"""
        text = self.result_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning("警告", "保存するテキストがありません。")
            return
            
        try:
            file_path = filedialog.asksaveasfilename(
                title="テキストファイルを保存",
                defaultextension=".txt",
                filetypes=[
                    ("テキストファイル", "*.txt"),
                    ("すべてのファイル", "*.*")
                ]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(text.strip())
                messagebox.showinfo("完了", f"テキストファイルを保存しました:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル保存エラー:\n{str(e)}")

    def append_to_input(self):
        """結果をinput.txtに追記（元のアプリケーション形式）"""
        text = self.result_text.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showwarning("警告", "追記するテキストがありません。")
            return
            
        try:
            # OCR結果を元アプリケーションの形式に整形
            formatted_text = self.format_ocr_result(text.strip())
            
            with open('input.txt', 'a', encoding='utf-8') as f:
                f.write(formatted_text + '\n\n')
            
            messagebox.showinfo("完了", "input.txtに追記しました。")
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル追記エラー:\n{str(e)}")

    def format_ocr_result(self, text):
        """OCR結果を元アプリケーションの形式に整形"""
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        formatted_lines = []
        jan_code = None
        
        for line in lines:
            # JANコードの検出（13桁の数字）
            if line.isdigit() and len(line) == 13:
                jan_code = line
                formatted_lines.append(f"JANコード\t{jan_code}")
            # ブランド名の検出
            elif any(keyword in line for keyword in ['千吉', 'ブランド', '金貴', '株式会社', '藤原産業']):
                formatted_lines.append(f"ブランド名\t{line}")
            # 商品名の検出
            elif any(keyword in line for keyword in ['根力', '力キ', 'キ', '鋸', 'ノコギリ']):
                formatted_lines.append(f"商品名\t{line}")
            # 規格の検出
            elif any(keyword in line for keyword in ['本', 'セット', '個', 'ホン', 'コ']):
                formatted_lines.append(f"規格\t{line}")
            # サイズ・重量の検出
            elif any(keyword in line for keyword in ['mm', 'cm', 'g', 'kg', '幅', '高さ', '奥行', '重量', '長さ']):
                formatted_lines.append(f"商品サイズ\t{line}")
            else:
                # その他の情報
                formatted_lines.append(line)
        
        return '\n'.join(formatted_lines)

    def clear_result(self):
        """結果をクリア"""
        self.result_text.delete("1.0", "end")

    def auto_detect_left_monitor(self):
        """左モニターのJANコード範囲を自動検出してOCR実行"""
        if not SCREENSHOT_AVAILABLE or not OCR_AVAILABLE:
            messagebox.showerror("エラー", "必要な機能が利用できません。")
            return

        try:
            # 左モニターの範囲を取得（仮定：左モニターは負の座標から始まる）
            monitors = self.get_monitor_info()
            left_monitor = self.find_left_monitor(monitors)
            
            if not left_monitor:
                messagebox.showerror("エラー", "左モニターが検出できませんでした。\n手動で範囲を設定してください。")
                return

            # カウントダウン
            for i in range(3, 0, -1):
                self.title(f"OCRテスト - {i}秒後に左モニターを撮影...")
                self.update()
                time.sleep(1)

            # 左モニター全体を撮影
            self.screenshot = pyautogui.screenshot(region=left_monitor)
            self.title("OCR機能テスト - JANコード商品情報抽出")

            # JANコードと商品情報の自動検出
            self.auto_detect_product_info()

        except Exception as e:
            messagebox.showerror("エラー", f"自動検出エラー:\n{str(e)}")
            self.title("OCR機能テスト - JANコード商品情報抽出")

    def get_monitor_info(self):
        """モニター情報を取得"""
        try:
            import screeninfo
            monitors = screeninfo.get_monitors()
            return [(m.x, m.y, m.width, m.height) for m in monitors]
        except ImportError:
            # screeninfoがない場合の代替方法
            try:
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()
                
                # 全体の画面サイズを取得
                total_width = root.winfo_screenwidth()
                total_height = root.winfo_screenheight()
                
                # 仮定：デュアルモニターで左が負の座標
                # 実際の環境に合わせて調整が必要
                left_width = total_width // 2
                monitors = [
                    (-left_width, 0, left_width, total_height),  # 左モニター
                    (0, 0, left_width, total_height)  # 右モニター
                ]
                
                root.destroy()
                return monitors
            except:
                return None

    def find_left_monitor(self, monitors):
        """左端のモニターを見つける"""
        if not monitors:
            return None
        
        # X座標が最も小さいモニターを左モニターとする
        left_monitor = min(monitors, key=lambda m: m[0])
        return left_monitor

    def auto_detect_product_info(self):
        """商品情報を自動検出してOCR実行"""
        try:
            if not self.screenshot:
                return

            # PIL ImageをOpenCV形式に変換
            opencv_image = cv2.cvtColor(np.array(self.screenshot), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)

            # JANコードらしい13桁の数字を検出するための前処理
            # 二値化
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            # 全体に対してOCRを実行して、JANコードの位置を特定
            try:
                # データを取得（位置情報付き）
                data = pytesseract.image_to_data(
                    binary, 
                    config=r'--oem 3 --psm 6 -l jpn+eng',
                    output_type=pytesseract.Output.DICT
                )

                # JANコードの位置を特定
                jan_region = self.find_jan_code_region(data)
                
                if jan_region:
                    # JANコード周辺の商品情報領域を推定
                    product_region = self.estimate_product_info_region(jan_region, self.screenshot.size)
                    
                    # 推定された領域をOCR処理
                    cropped_image = self.screenshot.crop(product_region)
                    self.perform_ocr_on_image(cropped_image)
                else:
                    # JANコードが見つからない場合は全体をOCR
                    messagebox.showinfo("情報", "JANコードを自動検出できませんでした。\n画面全体をOCR処理します。")
                    self.perform_ocr_on_image(self.screenshot)

            except Exception as e:
                messagebox.showerror("エラー", f"自動検出中にエラー: {str(e)}")
                # エラーの場合は全体をOCR
                self.perform_ocr_on_image(self.screenshot)

        except Exception as e:
            messagebox.showerror("エラー", f"商品情報検出エラー:\n{str(e)}")

    def find_jan_code_region(self, ocr_data):
        """OCRデータからJANコードの位置を特定"""
        import re
        
        for i, text in enumerate(ocr_data['text']):
            # 13桁の数字をチェック
            if re.match(r'^\d{13}

if __name__ == '__main__':
    print("=== OCR機能テスト - JANコード商品情報抽出 ===")
    print("起動中...")
    
    # customtkinterの外観設定
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    
    app = SimpleOCRApp()
    app.mainloop(), text.strip()):
                # JANコードが見つかった場合の座標
                x = ocr_data['left'][i]
                y = ocr_data['top'][i]
                w = ocr_data['width'][i]
                h = ocr_data['height'][i]
                return (x, y, w, h)
        
        # 12桁や14桁の数字もチェック（JANコードの可能性）
        for i, text in enumerate(ocr_data['text']):
            if re.match(r'^\d{12,14}

if __name__ == '__main__':
    print("=== OCR機能テスト - JANコード商品情報抽出 ===")
    print("起動中...")
    
    # customtkinterの外観設定
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    
    app = SimpleOCRApp()
    app.mainloop(), text.strip()):
                x = ocr_data['left'][i]
                y = ocr_data['top'][i]
                w = ocr_data['width'][i]
                h = ocr_data['height'][i]
                return (x, y, w, h)
        
        return None

    def estimate_product_info_region(self, jan_region, image_size):
        """JANコードの位置から商品情報領域を推定"""
        jan_x, jan_y, jan_w, jan_h = jan_region
        img_width, img_height = image_size
        
        # JANコードの右側から商品サイズまでの領域を推定
        # 一般的な商品ページのレイアウトを想定
        
        start_x = max(0, jan_x - 50)  # JANコードの少し左から
        start_y = max(0, jan_y - 20)  # JANコードの少し上から
        
        # 商品情報は通常JANコードの下部に続くので、縦に長めに取る
        end_x = min(img_width, jan_x + jan_w + 400)  # JANコードから右に400px
        end_y = min(img_height, jan_y + 300)  # JANコードから下に300px
        
        return (start_x, start_y, end_x, end_y)

    def perform_ocr_on_image(self, image):
        """指定された画像に対してOCR処理を実行"""
        try:
            # PIL ImageをOpenCV形式に変換
            opencv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            
            # 画像の前処理
            gray = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            denoised = cv2.fastNlMeansDenoising(gray)
            _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # OCR実行
            try:
                text = pytesseract.image_to_string(
                    binary, 
                    config=r'--oem 3 --psm 6 -l jpn+eng',
                    lang='jpn+eng'
                )
            except:
                try:
                    text = pytesseract.image_to_string(
                        binary, 
                        config=r'--oem 3 --psm 6 -l jpn',
                        lang='jpn'
                    )
                except:
                    text = pytesseract.image_to_string(binary, config=r'--oem 3 --psm 6')
            
            # 結果を表示
            self.result_text.delete("1.0", "end")
            self.result_text.insert("1.0", text.strip())
            
            if text.strip():
                pyperclip.copy(text.strip())
                messagebox.showinfo("自動検出完了", f"左モニターから商品情報を自動検出しました。\n\n文字数: {len(text.strip())}")
            else:
                messagebox.showwarning("警告", "テキストが検出されませんでした。")
                
        except Exception as e:
            messagebox.showerror("エラー", f"OCR処理エラー:\n{str(e)}")

    def set_manual_region(self):
        """手動で左モニターの範囲を設定"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("左モニター範囲設定")
        dialog.geometry("400x300")
        dialog.transient(self)
        dialog.grab_set()
        
        # 現在の設定を表示
        current_label = ctk.CTkLabel(dialog, text="左モニターの座標を設定してください", font=("Arial", 12, "bold"))
        current_label.pack(pady=10)
        
        # 入力フィールド
        frame = ctk.CTkFrame(dialog)
        frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(frame, text="X座標 (左端):").pack(pady=2)
        x_entry = ctk.CTkEntry(frame, placeholder_text="-1920")
        x_entry.pack(pady=2)
        
        ctk.CTkLabel(frame, text="Y座標 (上端):").pack(pady=2)
        y_entry = ctk.CTkEntry(frame, placeholder_text="0")
        y_entry.pack(pady=2)
        
        ctk.CTkLabel(frame, text="幅:").pack(pady=2)
        w_entry = ctk.CTkEntry(frame, placeholder_text="1920")
        w_entry.pack(pady=2)
        
        ctk.CTkLabel(frame, text="高さ:").pack(pady=2)
        h_entry = ctk.CTkEntry(frame, placeholder_text="1080")
        h_entry.pack(pady=2)
        
        # 説明
        info_label = ctk.CTkLabel(
            dialog, 
            text="例: 左モニターが1920x1080の場合\nX: -1920, Y: 0, 幅: 1920, 高さ: 1080",
            font=("Arial", 10)
        )
        info_label.pack(pady=10)
        
        def save_settings():
            try:
                x = int(x_entry.get() or "-1920")
                y = int(y_entry.get() or "0")
                w = int(w_entry.get() or "1920")
                h = int(h_entry.get() or "1080")
                
                self.left_monitor_region = (x, y, w, h)
                messagebox.showinfo("保存完了", f"左モニター範囲を設定しました:\n({x}, {y}, {w}, {h})")
                dialog.destroy()
                
            except ValueError:
                messagebox.showerror("エラー", "数値を正しく入力してください。")
        
        def test_region():
            try:
                x = int(x_entry.get() or "-1920")
                y = int(y_entry.get() or "0")
                w = int(w_entry.get() or "1920")
                h = int(h_entry.get() or "1080")
                
                # テスト撮影
                test_screenshot = pyautogui.screenshot(region=(x, y, w, h))
                messagebox.showinfo("テスト完了", f"指定範囲の撮影に成功しました。\nサイズ: {test_screenshot.size}")
                
            except Exception as e:
                messagebox.showerror("テストエラー", f"指定範囲の撮影に失敗しました:\n{str(e)}")
        
        # ボタン
        button_frame = ctk.CTkFrame(dialog)
        button_frame.pack(pady=10)
        
        test_button = ctk.CTkButton(button_frame, text="テスト撮影", command=test_region)
        test_button.pack(side="left", padx=5)
        
        save_button = ctk.CTkButton(button_frame, text="保存", command=save_settings)
        save_button.pack(side="left", padx=5)
        
        cancel_button = ctk.CTkButton(button_frame, text="キャンセル", command=dialog.destroy)
        cancel_button.pack(side="left", padx=5)

if __name__ == '__main__':
    print("=== OCR機能テスト - JANコード商品情報抽出 ===")
    print("起動中...")
    
    # customtkinterの外観設定
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    
    app = SimpleOCRApp()
    app.mainloop()