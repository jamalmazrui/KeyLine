@echo off
cls
tasklist /nh /fi "imagename eq Outlook.exe" | find /i "Outlook.exe" >nul && (echo Terminating Outlook.exe & taskkill /f /im Outlook.exe)