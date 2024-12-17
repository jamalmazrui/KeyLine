@echo off
cls
tasklist /nh /fi "imagename eq WinWord.exe" | find /i "WinWord.exe" >nul && (echo Terminating WinWord.exe & taskkill /f /im WinWord.exe)