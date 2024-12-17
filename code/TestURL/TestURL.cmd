@echo off
SetLocal EnableDelayedExpansion
cls

set testUrlDir=%~dp0
set testUrl=%testUrlDir%\testUrl.js
set node=C:\Program Files\nodejs\node.exe

set spec=%~1
"!node!" "!testUrl!" "!spec!"
