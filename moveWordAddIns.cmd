@echo off
setlocal
cls

echo Moving AddIns
set temp=%appdata%\Microsoft\AddIns\*.*
rem echo "%temp%" .
move "%temp%" .
