@echo off
SetLocal EnableDelayedExpansion
rem cls

echo Installing KeyLine to search path

set kl=%~dp0
set code=%kl%code
set klDir="%kl%"
set klDir=!klDir:~1,-2!
set pathed=%code%\pathed.exe
"%pathed%" /user /remove %klDir% >nul
"%pathed%" /user /add "%klDir%" >nul

if "%1"=="n" goTo :eof

set msg=Restart Windows now to complete installation? (y/n)
rem set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
