@echo off
cls
tasklist /nh /fi "imagename eq jfw.exe" | find /i "jfw.exe" >nul && (echo Terminating jfw.exe & taskkill /f /im jfw.exe)