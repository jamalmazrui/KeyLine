@echo off
setlocal
cls

echo Moving Normal.dotm to current directory
set source=%appdata%\Microsoft\Templates\Normal.dotm
set target=%cd%\Normal.dotm
move /y "%source%" "%target%" >nul
if not exist "%target%" echo Error
