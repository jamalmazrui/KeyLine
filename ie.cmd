@echo off
setLocal enableDelayedExpansion

set spec=%~1
if "%spec%"=="" set spec=*.*
if not exist "%spec%" echo No match & goTo :eof
for %%f in ("%spec%") do start "iexplore", "%%~f" %2 %3 %4 %5 %6 %7 %8 %9 & exit /b
