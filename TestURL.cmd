@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set testUrlDir=%kl%\code\TestURL
set testUrl=%testUrlDir%\testUrl.js
set node=C:\Program Files\nodejs\node.exe

set spec=%~1
"!node!" "!testUrl!" "!spec!"
