@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set pskill=%code%\pskill.exe
set settings=%kl%settings

"!pskill!" jfw
