import tkinter as tk
from tkinter import messagebox
from bs4 import BeautifulSoup
import pyperclip

ITEM_KEYS = ["JANコード", "ブランド名", "商品名", "規格", "商品サイズ", "重量"]

class JanExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("JANコード情報抽出ツール")
        self.geometry("400x300")

        # 説明ラベル
        self.info_label = tk.Label(self, text="クリップボードにコピーしたデータから\n必要な項目を抽出してinput.txtに保存します", 
                                  font=("メイリオ", 12), justify="center")
        self.info_label.pack(pady=20)

        # 抽出ボタン
        self.extract_btn = tk.Button(self, text="クリップボードから抽出", 
                                    command=self.extract_from_clipboard,
                                    font=("メイリオ", 14, "bold"),
                                    bg="#0078D4", fg="white",
                                    width=20, height=2)
        self.extract_btn.pack(pady=20)

        # 結果表示欄
        self.result_frame = tk.Frame(self)
        self.result_frame.pack(pady=10, fill="both", expand=True)
        
        self.result_labels = {}
        for key in ITEM_KEYS:
            lbl = tk.Label(self.result_frame, text=f"{key}: ", 
                          anchor="w", font=("メイリオ", 10),
                          relief="groove", bg="lightgray")
            lbl.pack(padx=10, pady=2, fill="x")
            self.result_labels[key] = lbl

    def extract_from_clipboard(self):
        try:
            # クリップボードからデータを取得
            clipboard_data = pyperclip.paste()
            if not clipboard_data.strip():
                messagebox.showwarning("警告", "クリップボードにデータがありません。")
                return

            # BeautifulSoupで解析
            html_text = '<root>' + ''.join(f'<item>{line}</item>' for line in clipboard_data.strip().split('\n')) + '</root>'
            soup = BeautifulSoup(html_text, 'html.parser')
            
            result = {}
            for item in soup.find_all('item'):
                if '\t' in item.text:
                    key, value = item.text.split('\t', 1)
                    result[key.strip()] = value.strip()
                else:
                    if item.text.startswith('重量'):
                        result['重量'] = item.text.replace('重量', '').strip()

            # 結果を表示
            for key in ITEM_KEYS:
                val = result.get(key, "（見つかりませんでした）")
                self.result_labels[key].config(text=f"{key}: {val}")

            # input.txtに保存
            with open("input.txt", "a", encoding="utf-8") as f:
                for key in ITEM_KEYS:
                    value = result.get(key, "")
                    f.write(f"{key}\t{value}\n")
                f.write("\n")

            messagebox.showinfo("完了", "クリップボードのデータを抽出してinput.txtに保存しました。")

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました: {e}")

if __name__ == '__main__':
    app = JanExtractorApp()
    app.mainloop() 

