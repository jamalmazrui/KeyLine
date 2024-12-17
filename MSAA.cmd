@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
set msaa=%kl%code\msaa.exe
set utf8b=%kl%code\utf8b.exe
set target=MSAA_Objects.txt

"!msaa!" %*
if not exist "!target!" echo Error & goTo :eof
