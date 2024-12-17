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
set pp2htm=%code%\pp2htm.vbs
set ppVert=%code%\ppVert64.exe
set xlVert=%code%\xlVert64.exe
set wdVert=%code%\wdVert64.exe
set utf8b=%code%\utf8b64.exe

set spec=%~1
if "%spec%"=="" set spec=*.pptx
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
set midRoot=!sourceRoot!
set midBase=!midRoot!!midExt!
rem set midDir=%temp%
set midDir=%cd%
set mid=!midDir!\!midBase!

set targetExt=.html
set targetRoot=!MidRoot!
set targetBase=!targetRoot!!targetExt!
set targetDir=%cd%
set target=!targetDir!\!targetBase!

if exist "!Mid!" exit /b
if exist "!target!" exit /b
echo !sourceBase!

if "!sourceExt!"==".htm" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".html" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".md" "!utf8b!" "!source!" >nul

set sourceDir=%cd%
rem pushd "%temp%"
rem If %errorlevel% NEQ 0 echo Error & goTo :eof
cscript.exe /nologo "%code%\pp2htm.vbs" "!source!" >nul
rem cscript.exe /nologo "%code%\pp2htm.vbs" "!source!"
if not exist "!mid!" echo Error & exit /b
rem popd
"!pandoc!" -s --quiet "!mid!" -o "!target!" >nul
rem del "!mid!"
if not exist "!target!" echo Error
exit /b
