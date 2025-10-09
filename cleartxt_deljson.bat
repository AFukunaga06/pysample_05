@echo off
rem フォルダ内のすべてのテキストファイルの中身を削除する
for %%f in (C:\pysample_01\*.txt) do (
    echo. > "%%f"
)

rem 特定のJSONファイル(samplep20240906-5ae36c9a4acd.json)を削除
del /q "C:\pysample_01\samplep20240906-5ae36c9a4acd.json"
