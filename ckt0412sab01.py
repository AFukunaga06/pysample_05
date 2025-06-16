import customtkinter as ctk
import pyperclip

def add_to_file():
   jan_code = entry.get()
   brand_name = "廃番"
   with open("input.txt", "a", encoding="utf-8") as file:
       file.write(f"JANコード\t{jan_code}\n")
       file.write(f"ブランド名\t{brand_name}\n\n")
   pyperclip.copy(f"JANコード\t{jan_code}\nブランド名\t{brand_name}\n")
   entry.delete(0, ctk.END)

# CTkのセットアップ
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# メインウィンドウの作成
root = ctk.CTk()
root.title("JANコード入力")

# フレームの作成
frame = ctk.CTkFrame(master=root)
frame.pack(pady=20, padx=20, fill="both", expand=True)

# ラベルの作成
label = ctk.CTkLabel(master=frame, text="JANコードを入力してください")
label.pack(pady=12, padx=10)

# エントリーの作成
entry = ctk.CTkEntry(master=frame, placeholder_text="JANコードを入力")
entry.pack(pady=12, padx=10)

# ボタンの作成
button = ctk.CTkButton(master=frame, text="追加", command=add_to_file)
button.pack(pady=12, padx=10)

# メインループの実行
root.mainloop()