@echo off
cls
tasklist /nh /fi "imagename eq Excel.exe" | find /i "Excel.exe" >nul && (echo Terminating Excel.exe & taskkill /f /im Excel.exe)