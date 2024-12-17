@echo off
setLocal enableDelayedExpansion
cls

set nirCmd=%~dp0bin\nircmd.exe

echo ScreenOff
"!nirCmd!" monitor off
