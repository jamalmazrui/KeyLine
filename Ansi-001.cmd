@echo off
setlocal
cls

for %%f in ("%1") do echo %%~nxf & "%~dp0bin\ansi64.exe" "%%~f" >nul
