import customtkinter as ctk
import subprocess
import sys
import os

class RunTestForm(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Test 実行フォーム")
        self.geometry("300x150")
        self.resizable(False, False)

        # ラベル
        label = ctk.CTkLabel(self, text="testtest_0428_01.py を実行", font=("Arial", 14))
        label.pack(pady=(20, 10))

        # 実行ボタン
        run_button = ctk.CTkButton(
            self,
            text="実行",
            command=self.run_test_script,
            fg_color="#0078D7",
            text_color="white",
            hover_color="#0053A6",
            width=120
        )
        run_button.pack(pady=10)

    def run_test_script(self):
        """外部スクリプトをサブプロセスで実行"""
        # 同一フォルダにある testtest_0428_01.py を実行
        script_path = os.path.join(os.path.dirname(__file__), "testtest_0428_01.py")
        if not os.path.isfile(script_path):
            ctk.CTkMessagebox(title="エラー", message=f"{script_path} が見つかりません。")
            return

        try:
            # python コマンドは環境に合わせて調整してください
            completed = subprocess.run(
                [sys.executable, script_path],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            # 実行結果をダイアログで表示
            ctk.CTkMessagebox(
                title="実行完了",
                message=f"標準出力:\n{completed.stdout}"
            )
        except subprocess.CalledProcessError as e:
            ctk.CTkMessagebox(
                title="実行エラー",
                message=f"コード: {e.returncode}\n標準エラー:\n{e.stderr}"
            )

if __name__ == "__main__":
    app = RunTestForm()
    app.mainloop()
