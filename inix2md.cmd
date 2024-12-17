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

set ppVert=%code%\ppVert64.exe
set wdVert=%code%\wdVert64.exe
set xlVert=%code%\xlVert64.exe

set spec=%~1
if "%spec%"=="" set spec=*.inix

if not exist "%spec%" echo No match & goto :eof
for %%f in ("%spec%") do call :work "%%~f" %2 %3 %4 %5 %6 %7 %8 %9
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
rem set midDir=%cd%
set midDir=%temp%
set mid=!midDir!\!midBase!

set targetExt=.md
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
If %errorlevel% NEQ 0 echo Error & exit /b
rem cscript.exe /nologo "%code%\ini2tables.vbs" "!source!" "" "" htm >nul
call "!kl!inix2docx.cmd" "!source!" >nul
popd
if not exist "!mid!" echo Error
call "!kl!docx2md.cmd" "!mid!" >nul
if exist "!mid!" del "!mid!"
if not exist "!target!" echo Error & exit /b

"!utf8b!" "!target!" >nul
exit /b
