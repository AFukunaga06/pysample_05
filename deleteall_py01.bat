@echo off
rem フォルダ内の全ファイルを削除
del /f /s /q "C:\pysample_01\*.*"

rem 空になったフォルダを削除
rmdir /s /q "C:\pysample_01"
