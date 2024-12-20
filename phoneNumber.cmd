@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set spec=%~1
cscript.exe /nologo "%code%\phoneNumber.vbs" %spec%
