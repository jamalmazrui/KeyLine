@echo off
cls
tasklist /nh /fi "imagename eq firefox.exe" | find /i "firefox.exe" >nul && (echo Terminating firefox.exe & taskkill /f /im firefox.exe)