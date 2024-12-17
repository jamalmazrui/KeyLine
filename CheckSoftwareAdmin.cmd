@echo off
setLocal enableDelayedExpansion
cls

openfiles >nul 2>&1
rem if not "%errorlevel%"=="0" echo Before running this command, please run cmd as administrator & echo For example, press WindowsKey, type "cmd" and then press Control+Shift+Enter & echo Remember to run this command from the KeyLine directory & goto :eof

echo The software package manager call Chocolatey will be installed, which will then be used to install Calibre, Pandoc, and LibreOffice.

rem echo a pause for confirmation occurs before each, where you can press Enter to continue or Control+C to cancel.
echo In general, you do not need to follow the installation log messages.
echo If you notice a serious error, however, try running cmd as administrator again and then rerunning this %~nx0 command.
echo About to install the latest Chocolatey software
rem pause
rem powershell.exe Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " [System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\code"

rem echo About to install the latest GitHub software
rem pause
rem call choco upgrade git.install -y

echo About to install the latest Node.js software
rem pause
call choco upgrade nodejs -y

echo About to install the latest Calibre software
rem pause
call choco upgrade Calibre -y

echo About to install the latest Pandoc software
rem pause
rem call choco upgrade Pandoc -y
choco upgrade pandoc -y --ia=ALLUSERS=1

echo About to install the latest LibreOffice software (which may take several minutes)
rem pause
call choco upgrade LibreOffice -y
refreshenv
