@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set durl=%code%\durl.exe

if not exist "%~1" goto url

for %%f in ("%~1") do "%durl%" "%%~f" -t "image" %4 %5 %6 %7 %8 %9
goto :eof

:url 
rem "%durl%" "%~1"
"%durl%" "%~1" -t "image" %4 %5 %6 %7 %8 %9

