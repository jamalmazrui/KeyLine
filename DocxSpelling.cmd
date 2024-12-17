@echo off
setlocal
cls

cscript.exe /nologo "%~dp0bin\DocxSpelling.vbs" %1 %2
