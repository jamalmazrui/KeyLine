@echo off
setLocal enableDelayedExpansion
cls

echo Installing KeyLine support packages
set kl=%1
if "!kl!"=="" set kl=%~dp0
cd "!kl!"
set node=C:\Program Files\nodejs\node.exe
set npm=C:\Program Files\nodejs\npm.cmd

rem if "%errorlevel%"=="0" goTo :start
set adminMode=0
openfiles >nul 2>&1
if "%errorlevel%"=="0" set adminMode=1
set adminMode=1
if "%adminMode%"=="1" goTo :start

echo This install command was not run with administrative rights.
echo It will be more automated when installing the Calibre, Pandoc, and LibreOffice software packages if you open the cmd environment as administrator.
echo For example, press WindowsKey, type "cmd" and then press Control+Shift+Enter.
echo Remember to run install.cmd from the KeyLine directory.
echo Press Control+C to Cancel, or for more manual installation steps,
rem pause

:start
rem if "%adminMode%"=="1" ( setx /m dircmd /b ) else setx dircmd /b
if "%adminMode%"=="1" ( setx /m dircmd /b >nul) else setx dircmd /b >nul
if "%adminMode%"=="1" assoc .md=txtfile >nul
if "%adminMode%"=="1" assoc .mdx=txtfile >nul
if "%adminMode%"=="1" assoc .ini=txtfile >nul
if "%adminMode%"=="1" assoc .inix=txtfile >nul

echo Set Drive W as directory !kl!\work
set startup=%appdata%\Microsoft\Windows\Start Menu\Programs\Startup
copy "!kl!\settings\setDriveW.lnk" "%startup%" >nul

echo Copy Hotkey shortcuts to desktop
set userdesktop=%userProfile%\desktop
copy "!kl!\settings\restartWindows.lnk" "%userdesktop%" >nul
copy "!kl!\settings\openChrome.lnk" "%userdesktop%" >nul
copy "!kl!\settings\openEdge.lnk" "%userdesktop%" >nul
copy "!kl!\settings\openFirefox.lnk" "%userdesktop%" >nul
copy "!kl!\settings\closeChrome.lnk" "%userdesktop%" >nul
copy "!kl!\settings\closeEdge.lnk" "%userdesktop%" >nul
copy "!kl!\settings\closeFirefox.lnk" "%userdesktop%" >nul

copy "!kl!\settings\openJAWS.lnk" "%userdesktop%" >nul
copy "!kl!\settings\openNVDA.lnk" "%userdesktop%" >nul
copy "!kl!\settings\closeJAWS.lnk" "%userdesktop%" >nul
copy "!kl!\settings\closeNVDA.lnk" "%userdesktop%" >nul

call "!kl!\InstallDesktopShortcut.cmd"
echo After the next restart of Windows, you can activate a command prompt with KeyLine active 
echo by using the keyboard shortcut Alt+Control+K
rem pause

call "!kl!\InstallSearchPath.cmd"
echo After the next restart of Windows, KeyLine functionality will be available from a command prompt in any directory
rem pause

set adminMode=1
if "%adminMode%"=="1" ( call "!kl!\CheckSoftwareAdmin.cmd" ) else call "!kl!\checkSoftware.cmd"

echo Installing TestURL support
cd "!kl!"\code\TestURL
call npm install

echo Installing TestPage support
cd "!kl!"\code\TestPage
call npm install

cd "!kl!"

rem pause
set msg=Restart Windows now to complete installation? (y/n)
rem set /p reply=%msg%
if "%reply%"=="y" shutdown.exe -r -f -t 1
