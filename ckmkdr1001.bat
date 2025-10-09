@echo off
chcp 65001 >nul

:: ディレクトリの存在をチェック
if exist C:\pysample_01 (
    echo C:\pysample_01 は既に存在します。
) else (
    echo C:\pysample_01 は存在しません。作成します。
    mkdir C:\pysample_01
)

pause
