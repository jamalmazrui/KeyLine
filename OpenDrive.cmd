@echo off
setLocal enableDelayedExpansion
cls

set nirCmd=%~dp0bin\nircmd.exe

echo OpenDrive %1
"!nirCmd!" cdrom open %1:
