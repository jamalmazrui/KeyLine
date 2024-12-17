@echo off
setLocal enableDelayedExpansion
cls

set nirCmd=%~dp0bin\nircmd.exe

"!nirCmd!" speak text ~$clipboard$
