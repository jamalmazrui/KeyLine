@echo off
setlocal
cls

for %%f in ("%1") do echo %%~nxf & "%~dp0bin\utf1664.exe" "%%~f" >nul
