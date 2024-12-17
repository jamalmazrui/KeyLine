@echo off
cls
tasklist /nh /fi "imagename eq slack.exe" | find /i "Slack.exe" >nul && (echo Terminating Slack.exe & taskkill /f /im Slack.exe)