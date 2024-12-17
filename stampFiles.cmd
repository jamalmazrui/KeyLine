@echo off
setLocal enableDelayedExpansion
cls

set spec=%~1
if "%spec%"=="" set spec=*.*

set startDir=%cd%
set kl=%~dp0
set code=%~dp0bin
set xlFormat=%kl%xlFormat.cmd
set md2htm=%kl%md2htm.cmd
set pandoc=C:\Program Files\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=%localAppData%\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=C:\Program Files (x86)\Pandoc\pandoc.exe
set utf8b=%code%\utf8b64.exe
set utf8n=%code%\utf8n64.exe


for /r %%f in ("%spec%") do call :work "%%f" "%spec%"
cd "%startDir%"
goTo :eof

:work
set sourceExt=%~x1
set sourceRoot="%~n1"
set sourceRoot=!sourceRoot:~1,-1!
set sourceBase=!sourceRoot!!sourceExt!
set sourceDir="%~dp1"
set sourceDir=!sourceDir:~1,-2!
set source=!sourceDir!\!sourceBase!

set midExt=.epub
set midRoot=!sourceRoot!
set midBase=!midRoot!!midExt!
set midDir=%temp%
set mid=!midDir!\!midBase!

set targetRoot=!sourceRoot!
set targetBase=!targetRoot!!targetExt!
set targetDir=!sourceDir!
set target=!targetDir!\!targetBase!

echo !sourceBase!
rem echo !target!

cd "!targetDir!"
copy /b "!source!" +,, >nul


cd "!targetDir!"
if "!sourceExt!"==".md" call "!md2htm!" "!source!" >nul
if not "%errorLevel%"=="0" echo Error

exit /b
