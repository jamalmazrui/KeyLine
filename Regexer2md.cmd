@echo off
setlocal
cls

set spec=%~1
if "%spec%"=="" set spec=*.mdx

if not exist "%spec%" echo No match & goto :eof
for %%f in ("%spec%") do if not exist "%%~nf.md" echo %%~nxf & "%~dp0Regexer.cmd" "%%~f" "%~dp0ini\mdx-Regexer.settings" "%%~nf.md" >nul
