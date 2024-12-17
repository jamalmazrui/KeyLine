@echo off
cls
tasklist /nh /fi "imagename eq NVDA.exe" | find /i "NVDA.exe" >nul && (echo Terminating NVDA.exe & taskkill /f /im NVDA.exe)