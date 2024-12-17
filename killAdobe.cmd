@echo off
cls
tasklist /nh /fi "imagename eq acrord32.exe" | find /i "acrord32.exe" >nul && (echo Terminating acrord32.exe & taskkill /f /im acrord32.exe)