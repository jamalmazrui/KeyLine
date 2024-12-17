@echo off
setLocal enableDelayedExpansion
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

set regexer=%kl%regexer.cmd
set sqlite=%code%\sqlite3.exe
set sql=select * from tmp order by numrow;

set spec=%~1
if "%spec%"=="" set spec=*.csv

if not exist "%spec%" echo No match & goto :eof
set db=:memory:
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

set midExt=.csv
set midRoot=tmp
set midBase=!midRoot!!midExt!
set midDir=%temp%
set mid=!midDir!\!midBase!
set midSlash=!midDir!/!midBase!

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

call "%regexer%" "%source%" numrow "%midBase%" >nul
"!sqlite!" "%db%" -csv ".import %midBase% tmp" ".save tmp.db"
if exist "%midBase%" del "%midBase%"
"!sqlite!" "tmp.db" -line -newline "~[]~" ".once %targetBase%" "%sql%"
if exist tmp.db del tmp.db
call regexer.cmd "%targetBase%" inix2ini >nul
if not exist "!target!" echo Error
exit /b


:work

"%sqlite%" "%db%" -line -newline ~[]~ ".once %inix%" "%sql%"
call regexer.cmd "%inix%" inix2ini >nul
