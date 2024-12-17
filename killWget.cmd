@echo off
cls
tasklist /nh /fi "imagename eq Wget.exe" | find /i "Wget.exe" >nul && (echo Terminating Wget.exe & taskkill /f /im Wget.exe)