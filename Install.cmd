@echo off
setLocal enableDelayedExpansion
cls

set kl=%1
if "!kl!"=="" set kl=%~dp0
cd "!kl!"
set workDir=!kl!\work
if not exist "%workDir%" md "%workDir%"

rem if "%errorlevel%"=="0" goTo :start
set adminMode=0
openfiles >nul 2>&1
if "%errorlevel%"=="0" set adminMode=1
rem set adminMode=1
if "%adminMode%"=="1" goTo :start

echo This install command was not run with administrative rights.
echo It will be more automated when installing the Calibre, Pandoc, and LibreOffice software packages if you open the cmd environment as administrator.
echo For example, press WindowsKey, type "cmd" and then press Control+Shift+Enter.
echo Remember to run install.cmd from the KeyLine directory.
echo Press Control+C to Cancel, or for more manual installation steps,
pause

:start
if "%adminMode%"=="1" ( setx /m dircmd /b ) else setx dircmd /b
if "%adminMode%"=="1" assoc .md=txtfile
if "%adminMode%"=="1" assoc .mdx=txtfile
if "%adminMode%"=="1" assoc .ini=txtfile
if "%adminMode%"=="1" assoc .inix=txtfile

call "!kl!\InstallDesktopShortcut.cmd"
echo After the next restart of Windows, you can activate a command prompt with KeyLine active 
echo by using the keyboard shortcut Alt+Control+K
pause

call "!kl!\InstallSearchPath.cmd"
echo After the next restart of Windows, KeyLine functionality will be available from a command prompt in any directory
pause

set adminmode=1
if "%adminMode%"=="1" ( call "!kl!\CheckSoftwareAdmin.cmd" ) else call "!kl!\checkSoftware.cmd"

pause
set msg=Restart Windows now to complete installation? (y/n)
set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
