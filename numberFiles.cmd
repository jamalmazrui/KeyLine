@echo off
setLocal enableDelayedExpansion
cls

set input=%~f1
set counter=0
rem for /F "tokens=*" %%l in ("%input%") do call :work "%%l"
for /F "tokens=*" %%l in (%1) do call :work "%%l"
goTo :eof

:work
set source=%~f1
if "!source!"=="" exit /b

set targetBase=%~nx1
echo !targetBase!
set /a counter+=1
set prefix=!counter!
if !counter! leq 9 set prefix=0!prefix!
if !counter! leq 99 set prefix=0!prefix!
set target=!prefix!-!targetBase!
echo = !target!
echo=
ren "!source!" "!target!"
exit /b
