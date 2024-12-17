@echo off
SetLocal EnableDelayedExpansion
cls

echo Uninstalling Pax from search path

set kl=%~dp0
set code=%kl%code
set pkDir="%kl%"
set pkDir=!pkDir:~1,-2!
set pathed=%code%\pathed.exe
"%pathed%" /user /remove %pkDir% >nul

set msg=Restart Windows now to complete the uninstall? (y/n)
set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
