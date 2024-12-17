@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

cscript.exe /nologo "%code%\xlImport.vbs" %*

