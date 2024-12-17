@echo off
cls
tasklist /nh /fi "imagename eq chrome.exe" | find /i "chrome.exe" >nul && (echo Terminating chrome.exe & taskkill /f /im chrome.exe)