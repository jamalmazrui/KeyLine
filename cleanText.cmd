@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings
set utf8b=%code%\utf8b64.exe
set utf8n=%code%\utf8n64.exe
set ftfy=%code%\ftfy.exe

set spec=%~1
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

set midExt=!sourceExt!
set midRoot=!sourceRoot!
set midBase=!midRoot!!midExt!
set midDir=%temp%
set mid=!midDir!\!midBase!

set targetExt=!sourceExt!
set targetRoot=!sourceRoot!
set targetBase=!targetRoot!!targetExt!
set targetDir=%cd%
set target=!targetDir!\!targetBase!

echo !sourceBase!

move "!source!" "!mid!" >nul
rem "!utf8n!" "!mid!" >nul
"!ftfy!" -g -o "!target!" "!mid!"

if not exist "!target!" echo Error & move "!mid!" "!target!" >nul & exit /b

del "!mid!" >nul
"!utf8b!" "!target!" >nul
exit /b
