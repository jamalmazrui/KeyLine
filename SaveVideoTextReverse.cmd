@echo off
setLocal enableDelayedExpansion
cls

set youtube=%~dp0bin\youtube-dl.exe
set url=%~1
if "%url%"=="" echo No URL & goTo :eof
rem "!youtube!" --write-auto-sub --write-description --write-sub --skip-download "%url%"
"!youtube!" --geo-bypass --no-check-certificate --playlist-reverse --restrict-filenames --sleep-interval 1 --yes-playlist --ignore-errors --write-auto-sub --write-description --write-sub --skip-download "%url%"
