@echo off
rem フォルダ内のすべてのテキストファイルの中身を削除するが、memo20241002.txtは除外
for %%f in (C:\pysample_01\*.txt) do (
    if not "%%~nxf" == "memo20241002.txt" (
        echo. > "%%f"
    )
)

rem 特定のJSONファイル(samplep20240906-5ae36c9a4acd.json)を削除
del /q "C:\pysample_01\samplep20240906-5ae36c9a4acd.json"
