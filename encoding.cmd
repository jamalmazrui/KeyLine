@echo off
setlocal
cls

if not exist "%~1" echo No match & goto :eof
if "%~2"=="" for %%f in ("%~1") do "%~dp0bin\encoding.exe" show "%%~f"
goto :eof

for %%f in ("%~1") do "%~dp0bin\encoding.exe" convert "%%~f" %2 %3
