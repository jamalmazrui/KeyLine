@echo off
setLocal enableDelayedExpansion
cls

set spec=%~1
if "%spec%"=="" set spec=*.docx

"%~dp0wd2pdf.cmd" "%spec%"
