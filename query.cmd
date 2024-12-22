@echo off
setLocal enableDelayedExpansion
cls

set spec=%~1
if not exist "%spec%" echo No match & goto :eof
for %%f in ("%spec%") do "%~dp0code\q.exe" "%%~f" %2 %3 %4 %5 %6 %7 %8 %9
