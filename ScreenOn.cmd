@echo off
setLocal enableDelayedExpansion
cls

set nirCmd=%~dp0bin\nircmd.exe

echo ScreenOn
"!nirCmd!" monitor on
