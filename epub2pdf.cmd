@echo off
setLocal enableDelayedExpansion
cls

set spec=%~1
if "%spec%"=="" set spec=*.epub

"%~dp0calibre2pdf.cmd" "%spec%"
