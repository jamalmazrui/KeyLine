@echo off
cls
tasklist /nh /fi "imagename eq pandoc.exe" | find /i "pandoc.exe" >nul && (echo Terminating pandoc.exe & taskkill /f /im pandoc.exe)