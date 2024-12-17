@echo off
cls
tasklist /nh /fi "imagename eq PowerPnt.exe" | find /i "PowerPnt.exe" >nul && (echo Terminating PowerPnt.exe & taskkill /f /im PowerPnt.exe)