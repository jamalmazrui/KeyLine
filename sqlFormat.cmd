@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set spec=%~1
if "%spec%"=="" set spec=*.sql

for %%f in ("%spec%") do echo %%~f & npx sql-formatter "%%~f" -l postgresql -t -u --lines-between-queries 2 -o "%%~f"  
