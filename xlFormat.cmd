@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set spec=%~1
if "%spec%"=="" set spec=*.xlsx

for %%f in ("%spec%") do cscript.exe /nologo "%code%\xlFormat.vbs" "%%~f"
