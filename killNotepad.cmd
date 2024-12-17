@echo off
cls
tasklist /nh /fi "imagename eq notepad.exe" | find /i "notepad.exe" >nul && (echo Terminating notepad.exe & taskkill /f /im notepad.exe)
S