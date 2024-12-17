@echo off
setLocal enableDelayedExpansion
cls

cscript.exe /nologo "%~dp0bin\GetAttachments.vbs" %*
