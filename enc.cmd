@echo off
setlocal
cls

if not exist "%~1" echo No match & goto :eof
echo here
for %%f in ("%~1") do echo >nul & echo %%~nxf & "%~dp0bin\enc.exe" "%%~f"
