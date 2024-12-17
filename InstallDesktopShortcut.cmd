@echo off
rem cls
echo Installing KeyLine to desktop shortcut
rem copy "%~dp0settings\KeyLine.lnk" "%UserProfile%\Desktop" >nul
set source=%~dp0settings\KeyLine.lnk
set target=%UserProfile%\Desktop\KeyLine.lnk
rem if exist "%source%" del "%source%"
copy "%source%" "%target%" >nul
if not exist "%target%" echo Error & goto :eof
if "%1"=="n" goTo :eof

set msg=Restart Windows now to complete installation? (y/n)
rem set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
