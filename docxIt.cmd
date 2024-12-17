@echo off
setlocal
cls

cscript.exe /nologo "%~dp0bin\DocxIt.vbs" %1 %2
