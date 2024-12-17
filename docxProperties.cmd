@echo off
setlocal
cls

cscript.exe /nologo "%~dp0bin\DocxProperties.vbs" %1 %2
