@echo off
cls
tasklist /nh /fi "imagename eq CScript.exe" | find /i "CScript.exe" >nul && (echo Terminating CScript.exe & taskkill /f /im CScript.exe)