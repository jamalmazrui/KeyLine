@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings
set calibre=C:\Program Files\calibre2\ebook-convert.exe
set pandoc=C:\Program Files\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=%localAppData%\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=C:\Program Files (x86)\Pandoc\pandoc.exe
set ppVert=%code%\ppVert64.exe
set xlVert=%code%\xlVert64.exe
set wdVert=%code%\wdVert64.exe
set utf8b=%code%\utf8b.exe

if not exist "%~1" echo No match & goto :eof

for %%f in ("%~1") do call :work "%%~f"
goto :eof

:work
set sourceExt=%~x1
set sourceRoot="%~n1"
set sourceRoot=!sourceRoot:~1,-1!
set sourceBase=!sourceRoot!!sourceExt!
set sourceDir="%~dp1"
set sourceDir=!sourceDir:~1,-2!
set source=!sourceDir!\!sourceBase!

set midExt=.docx
set midRoot=!sourceRoot!
set midBase=!midRoot!!midExt!
set midDir=%temp%
set mid=!midDir!\!midBase!

set targetExt=.inix
set targetRoot=!sourceRoot!
set targetBase=!targetRoot!!targetExt!
set targetDir=%cd%
set target=!targetDir!\!targetBase!

if exist "!target!" exit /b
echo !sourceBase!

if "!sourceExt!"==".htm" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".html" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".md" "!utf8b!" "!source!" >nul

pushd "%temp%"
if not "%errorLevel%"=="0" echo Error & exit /b
call "!kl!pandoc2docx" "!source!" " >nul
popd
if not exist "!mid!" echo Error & exit /b

call "!kl!wd2inix.cmd" "!mid!" >nul
if exist "!mid!" del "!mid!"

set targetSpec=!targetDir!\!targetRoot!*!targetExt!
if not exist "!targetSpec!" echo Error

set imageDir=!targetDir!\!targetRoot!_files
if exist "!imageDir!" rd /s /q "!imageDir!"

exit /b
