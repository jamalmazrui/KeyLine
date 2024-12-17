@echo off
cls
set app=%1
if "%app%"=="" echo No match & goTo :eof

tasklist /nh /fi "imagename eq %app%.exe" | find /i "%app%.exe" >nul && (echo Terminating %app%.exe & taskkill /f /im %app%.exe)