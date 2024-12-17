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
if "%spec%"=="" set spec=*.*

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

if exist "!target!" exit /b
echo !sourceBase!

if "!sourceExt!"==".htm" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".html" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".md" "!utf8b!" "!source!" >nul

rem for some reason, pdf2htm.cmd does not handle file name with & character
rem if "!sourceExt!"==".pdf" echo "!kl!pdf2htm.cmd" "!source!" >nul
if "!sourceExt!"==".pdf" call :pdf2htm "!source!"
if exist "!target!" "!utf8b!" "!target!" >nul & exit /b

set convert=!sourceExt:~1!2htm
set util=!kl!!convert!.cmd
rem echo util=!util!
if exist "!util!" call "!util!" "!source!" >nul
if exist "!target!" "!utf8b!" "!target!" >nul & exit /b

if exist "!wdVert!" call "!kl!wd2htm.cmd" "!source!" >nul
if exist "!target!" "!utf8b!" "!target!" >nul & exit /b

if exist "!calibre!" (if exist "!pandoc!" call "!kl!calibre_pandoc2htm.cmd" "!source!" >nul)
if exist "!target!" "!utf8b!" "!target!" >nul & exit /b

if exist "!pandoc!" call "!kl!pandoc2htm.cmd" "!source!" >nul
if exist "!target!" "!utf8b!" "!target!" >nul & exit /b

if exist "!libre!" call "!kl!libre2htm.cmd" "!source!" >nul
if exist "!target!" "!utf8b!" "!target!" >nul

if not exist "!target!" echo Error
exit /b

:pdf2htm
rem Copied from pdf2htm.cmd
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

if exist "!target!" exit /b
rem echo !sourceBase!

if "!sourceExt!"==".htm" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".html" "!utf8b!" "!source!" >nul
if "!sourceExt!"==".md" "!utf8b!" "!source!" >nul

if exist "!wdVert!" call "!kl!wd2htm.cmd" "!source!" >nul
if exist "!target!" exit /b

if exist "!calibre!" (if exist "!pandoc!" call "!kl!calibre_pandoc2htm.cmd" "!source!" >nul)
if exist "!target!" exit /b

if exist "!libre!" call "!kl!libre2htm.cmd" "!source!" >nul

if not exist "!target!" echo Error
exit /b
