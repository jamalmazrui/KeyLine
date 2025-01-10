@echo off
SetLocal EnableDelayedExpansion

cls
set db=%~1
set table=%~2
rem cscript.exe /nologo "%~dp0code\dbDot.vbs" "%db%" "%table%"
cscript.exe /nologo c:\KeyLine\code\dbDot.vbs c:\KeyLine\settings\Pax.db rules
