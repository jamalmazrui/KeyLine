@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
set sqlite=%kl%code\sqlite3.exe

set db=%~f1
set sql=%~2
set inix=%~f3
if "%inix%"=="" set inix=output.inix

"%sqlite%" "%db%" -line -newline ~[]~ ".once %inix%" "%sql%"
call regexer.cmd "%inix%" inix2ini
