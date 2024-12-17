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
set Regexer=%code%\regexer.exe
set ppMeta=%code%\ppMeta.vbs
set ppVert=%code%\ppVert64.exe
set xlVert=%code%\xlVert64.exe
set wdVert=%code%\wdVert64.exe
set utf8b=%code%\utf8b64.exe

set spec=%~1
if "%spec%"=="" set spec=*.vtt
if not exist "%spec%" echo No match & goto :eof

for %%f in ("%spec%") do call :work "%%~f"
goto :eof

:work
set sourceExt=%~x1
set sourceRoot="%~n1"
set sourceRoot=!sourceRoot:~1,-1!
set sourceBase=!sourceRoot!!sourceExt!
set sourceDir="%~dp1"
set sourceDir=!sourceDir:~1,-2!
set source=!sourceDir!\!sourceBase!

set midExt=.md
set midRoot=meta-!sourceRoot!
set midBase=!midRoot!!midExt!
rem set midDir=%temp%
set midDir=%cd%
set mid=!midDir!\!midBase!

set targetExt=.txt
set targetRoot=!SourceRoot!
set targetBase=!targetRoot!!targetExt!
set targetDir=%cd%
set target=!targetDir!\!targetBase!

if exist "!target!" exit /b
echo !sourceBase!

if "!sourceExt!"==".htm" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".html" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".md" "!utf8b!" "!source!" >nul

"!Regexer!" "!Source!" vtt2txt "!Target!"
if not exist "!target!" echo Error
exit /b
