@echo off
setlocal
cls

echo Moving StartUp
set temp=%appdata%\Microsoft\Word\STARTUP\*.*
rem echo "%temp%" .
move "%temp%" .
