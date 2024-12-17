@echo off
setlocal
cls

cscript.exe /nologo "%~dp0bin\DocxStyles.vbs" %1 %2
