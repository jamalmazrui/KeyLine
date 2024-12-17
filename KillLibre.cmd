@echo off
cls
tasklist /nh /fi "imagename eq soffice.exe" | find /i "soffice.exe" >nul && (echo Terminating soffice.exe & taskkill /f /im soffice.exe)