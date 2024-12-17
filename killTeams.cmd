@echo off
cls
tasklist /nh /fi "imagename eq ms-teams.exe" | find /i "Teams.exe" >nul && (echo Terminating Teams.exe & taskkill /f /im Teams.exe)