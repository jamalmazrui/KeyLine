@echo off
setLocal enableDelayedExpansion
cls

set spec=%~1
if not exist "%spec%" echo No match & goto :eof
for %%f in ("%spec%") do "%~dp0bin\regexer.exe" "%%~f" %2 %3
