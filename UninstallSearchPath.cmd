@echo off
SetLocal EnableDelayedExpansion
cls

echo Uninstalling KeyLine from search path

set kl=%~dp0
set code=%kl%code
set klDir="%kl%"
set klDir=!klDir:~1,-2!
set pathed=%code%\pathed.exe
"%pathed%" /user /remove %klDir% >nul

set msg=Restart Windows now to complete the uninstall? (y/n)
set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
