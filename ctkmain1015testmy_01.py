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

# スプレッドシートID
スプレッドシートID = '17Le1KA9nzMREt0Qp9_elM1OF1q8aSp-GDBZRPOntNI8'

# JSONキーファイルへのパス
キーファイルパス = r'C:\pysample_01\samplep20240906-5ae36c9a4acd.json'

# 認証情報を取得
認証情報 = Credentials.from_service_account_file(キーファイルパス, scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'])

# gspreadクライアントを作成
クライアント = gspread.authorize(認証情報)

# スプレッドシートを開く
ワークシート = クライアント.open_by_key(スプレッドシートID).sheet1

# スクリプトのファイル名を取得
ファイル名 = os.path.basename(sys.argv[0])

# 安全にファイル操作を行うためのデコレーター関数
def 安全なファイル操作(func):
    def ラッパー(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            messagebox.showwarning("ファイルが見つかりません", f"{str(e)}ファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")
    return ラッパー

# クリップボードのデータを処理し、JANコードとともにファイルに書き込む関数
@安全なファイル操作
def クリップボードデータ処理(janコード, クリップボードデータ):
    行 = クリップボードデータ.split('\n')
    出力データ = "\n".join(行.strip() for 行 in 行 if 行.strip())
    if janコード:
        出力データ = f"JANコード\t{janコード}\n{出力データ}"
    with open('input.txt', 'a', encoding='utf-8') as ファイル:
        ファイル.write(出力データ + '\n\n')
    print('新しいデータがinput.txtに追加されました。')

# JANコードの重複と廃番をチェックし、結果を出力する関数
@安全なファイル操作
def JANコードチェック(出力):
    with open('input.txt', 'r', encoding='utf-8') as f:
        データ文字列 = f.read()
    JANコード一覧 = re.findall(r'JANコード\t(\d+)', データ文字列)
    廃番コード = [コード for コード in JANコード一覧 if 'ブランド名\t廃番' in データ文字列.split(f'JANコード\t{コード}')[1].split('JANコード')[0]]
    重複JANコード = [コード for コード, 回数 in Counter(JANコード一覧).items() if 回数 > 1]
    出力.delete("1.0", customtkinter.END)
    if 重複JANコード:
        for 重複 in 重複JANコード:
            出力.insert(customtkinter.END, f"JANコード {重複} が重複しています\n")
    else:
        出力.insert(customtkinter.END, "重複はありません\n")
    for コード in 廃番コード:
        出力.insert(customtkinter.END, f"JANコード {コード} は廃番です\n")
    出力.insert(customtkinter.END, f"JANコードは上から{len(JANコード一覧)}番目です\n")
    if JANコード一覧:
        出力.insert(customtkinter.END, f"現在のJANコードは{JANコード一覧[-1]}です\n")

# クリップボードの内容を貼り付けて処理を実行する関数
def 貼り付けて実行():
    クリップボードデータ = pyperclip.paste()
    if not クリップボードデータ:
        messagebox.showwarning("警告", "クリップボードにデータがありません。")
        return
    クリップボードデータ処理("", クリップボードデータ)
    ウィンドウ.出力.insert(customtkinter.END, クリップボードデータ + "\n")

# ファイルを開く関数
@安全なファイル操作
def ファイルを開く(ファイルパス):
    os.startfile(ファイルパス)

# バッチファイルを実行し、出力ファイルを開く関数
def バッチ実行とファイル表示(バッチファイル, 出力ファイル):
    @安全なファイル操作
    def 実行():
        subprocess.run([バッチファイル])
        time.sleep(1)
        ファイルを開く(出力ファイル)  # バッチファイル実行後に出力ファイルを開く
        データ差分チェック(ウィンドウ.出力)  # バッチファイル実行後にデータ抜けチェックを自動実行
    実行()

# データ抜けのチェックを行い、checkd01.txtが空の場合にメッセージを表示する関数
def checkd01空チェック():
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as ファイル:
            内容 = ファイル.read().strip()
            if not 内容:
                messagebox.showinfo("情報", "checkd01.txtは空です")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

# バッチファイルを実行し、ファイルを開かずにデータをチェックする関数
def バッチ実行とデータチェック(バッチファイル):
    @安全なファイル操作
    def 実行():
        subprocess.run([バッチファイル])
        time.sleep(1)
        checkd01空チェック()  # checkd01.txtが空かどうかを確認
        データ差分チェック(ウィンドウ.出力)  # バッチファイル実行後にデータ抜けチェックを自動実行
    実行()

# 複数のファイルの内容をクリアする関数
def ファイルクリア():
    if messagebox.askyesno("確認", "本当にクリアして良いですか？"):
        try:
            for ファイル in ["checkd01.txt", "checkd02.txt", "output02.txt", "output.txt", "input.txt"]:
                with open(ファイル, 'w', encoding='utf-8') as f:
                    f.write('')
            messagebox.showinfo("クリア完了", "ファイルのデータをクリアしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルをクリアする際にエラーが発生しました： {str(e)}")

# 指定された範囲内で欠番をチェックする関数
def 欠番チェック(現在のインデックス, 指定インデックス):
    欠番 = []
    for i in range(現在のインデックス, 指定インデックス):
        セル値 = ワークシート.acell(f'A{i + 1}').value
        time.sleep(1)  # 1秒の待機時間を追加
        if セル値 is None:
            欠番.append(i + 1)
    return 欠番

# データの差分をチェックする関数
@安全なファイル操作
def データ差分チェック(出力):
    try:
        with open('checkd01.txt', 'r', encoding='utf-8') as ファイル1, open('checkd02.txt', 'r', encoding='utf-8') as ファイル2:
            行1 = ファイル1.readlines()
            行2 = ファイル2.readlines()
            差分 = []
            for i, (行1, 行2) in enumerate(zip(行1, 行2)):
                janコード1 = re.search(r'\d+', 行1)
                janコード2 = re.search(r'\d+', 行2)
                if janコード1 and janコード2 and janコード1.group() != janコード2.group():
                    差分.append(f"{i + 1}番目が違います")
            出力.delete("1.0", customtkinter.END)
            if 差分:
                出力.insert(customtkinter.END, "\n".join(差分))
            else:
                出力.insert(customtkinter.END, "すべてのJANコードが一致しています")
    except FileNotFoundError:
        messagebox.showwarning("ファイルが見つかりません", "checkd01.txt または checkd02.txt が見つかりません。")
    except Exception as e:
        messagebox.showerror("エラー", f"操作中にエラーが発生しました： {str(e)}")

# メインウィンドウのクラス
class メインウィンドウ(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("アプリケーション - " + ファイル名)
        self.geometry("800x600")
        self.UI設定()

    def UI設定(self):
        self.フレーム = customtkinter.CTkFrame(self)
        self.フレーム.pack(fill=customtkinter.BOTH, expand=True, padx=20, pady=20)
        self.出力 = customtkinter.CTkTextbox(self.フレーム, width=760, height=200)
        self.出力.grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        ボタン = [
            ("重複と廃番とデータ抜けのチェック", self.重複廃番データ抜けチェック, 1, 0, "#008000"),  # 濃い緑色
            ("JANコードのコピー", self.JANコードコピー, 1, 1, "#FF0000"),
            ("テキストに貼り付けと実行", 貼り付けて実行, 2, 1, "#008000"),  # 濃い緑色
            ("データ抜けのチェック", lambda: バッチ実行とデータチェック("Type1x.bat"), 2, 2, "#FF0000"),
            ("チェックシートを開く", lambda: ファイルを開く(r'C:\pysample_01\checksheet12.bat'), 3, 0, "#0078D4"),
            ("座標軸とコピー", self.座標サブフォーム表示, 2, 0, "#FF0000"),
            ("商品情報入力シートを開く", lambda: ファイルを開く('syouhin_n.bat'), 3, 1, "#0078D4"),
            ("藤原産業を開く", lambda: ファイルを開く('fujiwarasanngyou.bat'), 3, 2, "#0078D4"),
            ("input.txtのチェック", self.inputファイルチェック, 4, 0, "#0078D4"),
            ("サブフォーム廃番処理", self.サブフォーム表示, 4, 1, "#0078D4"),
            ("Type2x.bat実行とoutput.txt表示", lambda: バッチ実行とファイル表示("Type2x.bat", "output.txt"), 4, 2, "#0078D4"),
            ("checkd01.txtを開く", lambda: ファイルを開く('checkd01.txt'), 5, 1, "#0078D4"),
            ("checkd02.txtを開く", lambda: ファイルを開く('checkd02.txt'), 5, 2, "#0078D4"),
            ("クリップボードのクリア", ファイルクリア, 5, 0, "#0078D4"),
            ("input.txtを開く", lambda: ファイルを開く('input.txt'), 1, 2, "#FF0000"),
        ]
        for テキスト, コマンド, 行, 列, 色 in ボタン:
            ボタン = customtkinter.CTkButton(self.フレーム, text=テキスト, command=コマンド, fg_color=色)
            ボタン.grid(row=行, column=列, padx=5, pady=2, sticky="ew")
        for i in range(4):
            self.フレーム.columnconfigure(i, weight=1)

    def 重複廃番データ抜けチェック(self):
        JANコードチェック(self.出力)
        バッチ実行とデータチェック("Type1x.bat")
        データ差分チェック(self.出力)

    # JANコードのコピーを実行する関数
    @安全なファイル操作
    def JANコードコピー(self):
        subprocess.Popen(['python', 'jancopy0906_01.py'])

	# サブフォームを開く関数
    @安全なファイル操作
    def サブフォーム表示(self):
        subprocess.Popen(['python', 'ckt0412sab01.py'])

    # 座標サブフォームを開く関数
    @安全なファイル操作
    def 座標サブフォーム表示(self):
        subprocess.Popen(['python', 'Copycoord0902_01.py'])

    # input.txtファイルをチェックする関数
    def inputファイルチェック(self):
        try:
            with open('input.txt', 'r', encoding='utf-8') as f:
                行 = f.readlines()
            無効ANコード行 = [i for i, 行 in enumerate(行, 1) if 'ANコード' in 行 and 'J' not in 行]
            十三桁数値行 = [i for i, 行 in enumerate(行, 1) if re.match(r'^\d{13}$', 行.strip())]
            メッセージ = (
                f"「ANコード」に「J」が含まれていない行: {', '.join(map(str, 無効ANコード行)) or 'なし'}\n"
                f"13桁の数値のみを含む行: {', '.join(map(str, 十三桁数値行)) or 'なし'}"
            )
            messagebox.showinfo("チェック結果", メッセージ)
        except FileNotFoundError:
            messagebox.showerror("エラー", "input.txtファイルが見つかりません。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルのチェック中にエラーが発生しました： {str(e)}")

if __name__ == '__main__':
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    ウィンドウ = メインウィンドウ()
    ウィンドウ.mainloop()