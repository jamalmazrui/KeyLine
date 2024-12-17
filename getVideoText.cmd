@echo off
setLocal enableDelayedExpansion
cls

set youtube=%~dp0bin\youtube-dl.exe
set url=%~1
if "%url%"=="" echo No URL & goTo :eof
"!youtube!" --write-auto-sub --write-description --write-sub --skip-download "%url%"
