@echo off
cls
setLocal enableDelayedExpansion
cls

"%~dp0bin\sqlite3.exe" %*
