@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set testPageDir=%code%\TestWeb
set testPage=%testPageDir%\testPage.js
set node=C:\Program Files\nodejs\node.exe
set settings=%kl%settings

set spec=%~1
rem echo %spec%
if not exist "%spec%" (
rem echo Not a file, so assume url
rem echo "!node!" "!testPage!" "%spec%"
"!node!" "!testPage!" "%spec%"
exit /b 1
)

rem echo a File
set urlList=%spec%
for /f "tokens=* delims=" %%A in (%urlList%) do (
set trimLine=%%A
rem set trimLine=!line:~0,-1!
for /f "tokens=* delims= " %%B in ("!trimLine!") do set trimLine=%%B
if not "!trimLine!"=="" (
echo !trimLine!
"!node!" "!testPage!" "!trimLine!"
)
)


