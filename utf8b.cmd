@echo off
setlocal
cls

for %%f in ("%1") do echo %%~nxf & "%~dp0bin\utf8b64.exe" "%%~f" >nul
