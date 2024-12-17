@echo off
cls
tasklist /nh /fi "imagename eq ebook-convert.exe" | find /i "ebook-convert.exe" >nul && (echo Terminating ebook-convert.exe & taskkill /f /im ebook-convert.exe)