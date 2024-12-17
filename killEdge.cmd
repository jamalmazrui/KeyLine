@echo off
cls
tasklist /nh /fi "imagename eq msedge.exe" | find /i "msedge.exe" >nul && (echo Terminating msedge.exe & taskkill /f /im msedge.exe)