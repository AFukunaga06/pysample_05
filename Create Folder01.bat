@echo off
chcp 65001
echo 現在のドライブ： %~d0
xcopy %~d0\source\*.* C:\pysample_01 /s /e /y
echo コピーが完了しました。
pause
