@echo off
cls
echo Uninstalling Pax from desktop shortcut
set source=%~dp0ini\Pax.lnk
set target=%UserProfile%\Desktop\Pax.lnk
if exist "%target%" del "%target%"
if exist "%target%" echo Error & goto :eof

set msg=Restart Windows now to complete the uninstall? (y/n)
set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
