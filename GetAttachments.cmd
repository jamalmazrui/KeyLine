@echo off
setLocal enableDelayedExpansion
cls

cscript.exe /nologo "%~dp0code\GetAttachments.vbs" %*
