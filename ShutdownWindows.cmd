@echo off
cls

set seconds=%1
if "%seconds%"=="" set seconds=5
shutdown.exe -s -f -t %seconds%
