@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set spec=%~1
rem cscript.exe /nologo "%code%\cal.vbs" "%spec%"
cscript.exe /nologo "%code%\cal.vbs" %spec%
