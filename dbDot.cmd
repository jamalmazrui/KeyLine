@echo off
SetLocal EnableDelayedExpansion

cls
set db=%~1
set table=%~2
cscript.exe /nologo "%~dp0bin\dbDot.vbs" "%db%" "%table%"
