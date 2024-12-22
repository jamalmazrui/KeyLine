@echo off
setLocal enableDelayedExpansion
cls

REM Ensure the script is run with administrative privileges
whoami /groups | find "S-1-16-12288" >nul
if errorlevel 1 (
echo This script requires administrative privileges.
echo Please right-click the batch file and select "Run as Administrator."
exit /b
)

REM Set DIRCMD to "/b"
echo Setting DIRCMD environment variable to "/b"
setx DIRCMD /b /M >nul 2>&1
if %errorlevel% equ 0 (
echo Successfully set DIRCMD to "/b". The change will take effect in new Command Prompt windows.
) else (
echo Failed to set DIRCMD. Please ensure you have administrative privileges.
)

REM Configure older cmd.exe as the default terminal application
echo Configuring the older cmd.exe environment as the default terminal application
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Terminal\Settings" /v DefaultTerminalApplication /t REG_SZ /d "{0caa0dad-35be-5f56-a8ff-afceeeaa6101}" /f >nul 2>&1
if %errorlevel% equ 0 (
echo Successfully set the old cmd.exe as the default terminal application.

REM Verify the registry key value
rem echo Verifying the registry key value...
for /f "tokens=2*" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Terminal\Settings" /v DefaultTerminalApplication 2^>nul') do (
if "%%B"=="{0caa0dad-35be-5f56-a8ff-afceeeaa6101}" (
echo Registry key verified successfully.
) else (
echo Registry key value is incorrect. Expected "{0caa0dad-35be-5f56-a8ff-afceeeaa6101}" but found "%%B".
exit /b
)
)
) else (
echo Failed to set the old cmd.exe as the default terminal application. Please ensure you have administrative privileges.
)

echo Installing KeyLine support packages
set kl=%1
if "!kl!"=="" set kl=%~dp0

set node=C:\Program Files\nodejs\node.exe
set npm=C:\Program Files\nodejs\npm.cmd

cd "!kl!"
echo Set Drive W as convenient work directory
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

call "!kl!\InstallSearchPath.cmd"
echo After the next restart of Windows, KeyLine functionality will be available from a command prompt in any directory

echo The package manager call Chocolatey will be installed, and then be used to install GitHub, Node.js, Calibre, Pandoc, and LibreOffice.
echo In general, you do not need to follow the log messages that indicate activity.
rem suppress verbose output
rem @"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " [System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\code" >nul 2>&1
@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " [System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1')) | Out-Null" > nul 2>&1

echo Installing the latest GitHub software
call choco upgrade git.install -y

echo Installing the latest Node.js software
call choco upgrade nodejs -y

echo Installing the latest Calibre software
call choco upgrade Calibre -y

echo Installing the latest Pandoc software
call choco upgrade pandoc -y --ia=ALLUSERS=1

echo Installing the latest LibreOffice software (which may take several minutes)
call choco upgrade LibreOffice -y

echo Installing TestURL support
cd "!kl!\code\TestURL"
call npm install

echo Installing TestPage support
cd "!kl!\code\TestPage"
call npm install

cd "!kl!"
refreshenv
