@echo off
cls
tasklist /nh /fi "imagename eq FileDir.exe" | find /i "FileDir.exe" >nul && (echo Terminating FileDir.exe & taskkill /f /im FileDir.exe)