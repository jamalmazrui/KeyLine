@echo off
SetLocal EnableDelayedExpansion
cls

set testPageDir=%~dp0
set testPage=%testPageDir%\testPage.js
set node=C:\Program Files\nodejs\node.exe

set spec=%~1
if not exist "%spec%" (
"!node!" "!testPage!" "%spec%"
exit /b 1
)

set urlList=%spec%
for /f "tokens=* delims=" %%A in (%urlList%) do (
set line=%%A
for /f "tokens=* delims= " %%B in ("!line!") do set line=%%B
if not "!line!"=="" (
echo Testing !line!
"!node!" "!testPage!" "!line!"
)
)


