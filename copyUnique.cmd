@echo off
setLocal enableDelayedExpansion
cls

set spec=%~1
if "%spec%"=="" set spec=*.*

set targetDir=%~2
if "%targetDir%"=="" echo The first parameter should be a file specification, and the second should be a target directory & goTo :eof

if not exist "%targetDir%" call :makeDir "%targetDir%"
if not exist "%targetDir%" goTo :eof

for /r %%f in ("%spec%") do call :work "%%f" "%targetDir%"
goTo :eof

:work
set source=%~1
set sourceBase=%~nx1
echo %sourceBase%
set targetDir=%~2
set targetRoot=%~n1
set targetRootRandom=%targetRoot%-%random%
set targetExt=%~x1
set target=%targetDir%\%targetRoot%%targetExt%
set targetRandom=%targetDir%\%targetRootRandom%%targetExt%
if exist "%target%" set target=!targetRandom!
rem echo source=%source%
rem echo target=%target%
copy "!source!" "!target!" >nul
exit /b

:makeDir
set targetDir=%~1
set msg=Directory %targetDir% not found.  Create it? (y/n)
set /p reply=%msg%
if "%reply%"=="y" md "%targetDir%"
exit /b

