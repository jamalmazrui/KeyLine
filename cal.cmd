@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set arg1=%~1
set arg2=%~2
rem cscript.exe /nologo "%code%\cal.vbs" "%arg1%" "%arg2%"
cscript.exe /nologo "%code%\cal.vbs" %arg1% %arg2%
