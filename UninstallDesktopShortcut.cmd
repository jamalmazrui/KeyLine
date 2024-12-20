@echo off
cls
echo Uninstalling KeyLine from desktop shortcut
set source=%~dp0settings\KeyLine.lnk
set target=%UserProfile%\Desktop\KeyLine.lnk
if exist "%target%" del "%target%"
if exist "%target%" echo Error & goto :eof

set msg=Restart Windows now to complete the uninstall? (y/n)
set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
