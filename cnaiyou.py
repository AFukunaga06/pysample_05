import os

# Cドライブの内容を調べる
def list_c_drive_contents():
    c_drive_path = 'C:\\'
    try:
        # Cドライブの内容を取得
        contents = os.listdir(c_drive_path)
        print("Cドライブの内容:")
        for item in contents:
            print(item)
    except Exception as e:
        print(f"エラーが発生しました: {e}")

# 関数を実行
list_c_drive_contents()
