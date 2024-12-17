@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
set metabook=%kl%code\metabook.exe
"!metabook!" -t %1
