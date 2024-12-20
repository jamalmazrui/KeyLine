@echo off
setlocal
cls

cscript.exe /nologo "%~dp0code\DocxSpelling.vbs" %1 %2
