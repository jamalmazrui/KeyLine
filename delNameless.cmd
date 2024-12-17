@echo off
SetLocal EnableDelayedExpansion

cls
set spec=%~1
if "%spec%"=="" set spec=*.*
if not exist "%spec%" echo No Match & goTo :eof
cscript.exe /nologo "%~dp0bin\delNameless.vbs" "%spec%"
