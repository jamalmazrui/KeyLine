@echo off
SetLocal EnableDelayedExpansion
cls

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
set 7z=%code%\7z.exe

set ppVert=%code%\ppVert64.exe
set wdVert=%code%\wdVert64.exe
set xlVert=%code%\xlVert64.exe

set spec=%~1
if "%spec%"=="" set spec=*.zip
set spec2=%~2
if "%spec2%"=="" set spec2=*

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

set midExt=.epub
set midRoot=!sourceRoot!
set midBase=!midRoot!!midExt!
set midDir=%temp%
set mid=!midDir!\!midBase!

set targetExt=.htm
set targetRoot=!sourceRoot!
set targetBase=!targetRoot!!targetExt!
set targetDir=%cd%
set target=!targetDir!\!targetBase!

echo !sourceBase!

if "!sourceExt!"==".htm" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".html" "!utf8b!" "!source!" >nul 
if "!sourceExt!"==".md" "!utf8b!" "!source!" >nul

rem "!7z!" -y x "!source!" >nul
"!7z!" -y x "!source!" "!spec2!" >nul
if errorLevel 1 echo Error
exit /b
