@echo off
setLocal enableDelayedExpansion
cls

set nirCmd=%~dp0bin\nircmd.exe

echo CloseDrive %1
"!nirCmd!" cdrom close %1:
