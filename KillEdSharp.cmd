@echo off
cls
tasklist /nh /fi "imagename eq EdSharp.exe" | find /i "EdSharp.exe" >nul && (echo Terminating EdSharp.exe & taskkill /f /im EdSharp.exe)