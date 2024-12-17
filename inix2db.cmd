@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code

set inix=%~f1
rem set settings=%~dpn1.settings
set settings=%temp%\%~n1.settings
set inix=%~f1
set db=%~f2

rem call regexer.cmd "%inix%" inix2ini "%settings%" >nul
"C:\anaconda3\python.exe" "%code%\inix2db.py" "%inix%" "%db%"
 if exist "%settings%" del "%settings%"
