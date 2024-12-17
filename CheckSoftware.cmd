@echo off
SetLocal EnableDelayedExpansion
rem cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set calibre=C:\Program Files\calibre2\ebook-convert.exe
set libre=c:\program files\libreoffice\program\soffice.exe
set pandoc=C:\Program Files\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=%localAppData%\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=C:\Program Files (x86)\Pandoc\pandoc.exe
set utf8b=%code%\utf8b64.exe
set utf8n=%code%\utf8n64.exe

set ppVert=%code%\ppVert64.exe
set wdVert=%code%\wdVert64.exe
set xlVert=%code%\xlVert64.exe

set spec=%~1

echo Checking for open source software that supports file conversions by KeyLine
call :InstallCalibre
call :InstallPandoc
call :InstallLibre

echo Run any installers that you downloaded, accepting all default choices.
rem echo An exception is that, for Pandoc, check the box to install for all users, not just the current user.
goto :eof

:InstallCalibre
set msg=Calibre not found.  Download it? (y/n)
if exist "%calibre%" set msg=Calibre found. Download latest version? (y/n)
set /p reply=%msg%
if "%reply%"=="y" "%kl%openWait.cmd" "https://calibre-ebook.com/dist/win64"
exit /b

:InstallPandoc
set msg=Pandoc not found.  Download it? (y/n)
if exist "%pandoc%" set msg=Pandoc found. Download latest version? (y/n)
set /p reply=%msg%
if "%reply%"=="y" echo Activate the link called Download the latest installer for Windows (64-bit) & "%kl%openWait.cmd" "https://pandoc.org/installing.html"
exit /b

:InstallLibre
set msg=LibreOffice not found.  Download it? (y/n)
if exist "%libre%" set msg=LibreOffice found. Download latest version? (y/n)
set /p reply=%msg%
if "%reply%"=="y" echo Activate the Download link that occurs after the text about choosing your operating system & "%kl%openWait.cmd" "https://www.libreoffice.org/download/download/"
exit /b
