@echo off
cls
tasklist /nh /fi "imagename eq IEExplore.exe" | find /i "IEExplore.exe" >nul && (echo Terminating IEExplore.exe & taskkill /f /im IEExplore.exe)