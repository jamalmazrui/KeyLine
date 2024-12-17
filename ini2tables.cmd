@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set spec=%~1
if "%spec%"=="" set spec=*.settings
for %%f in ("%spec%") do cscript.exe /nologo "%code%\ini2tables.vbs" "%%~f" %2 %3 %4 %5 %6 %7 %8 %9
