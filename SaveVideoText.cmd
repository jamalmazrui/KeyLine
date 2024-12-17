@echo off
setLocal enableDelayedExpansion
cls

set youtube=%~dp0bin\youtube-dl.exe
set url=%~1
if "%url%"=="" echo No URL & goTo :eof
rem "!youtube!" --write-auto-sub --write-description --write-sub --skip-download "%url%"
rem "!youtube!" --geo-bypass --no-check-certificate --restrict-filenames --sleep-interval 1 --yes-playlist --ignore-errors --write-auto-sub --write-description --write-sub --skip-download "%url%"
"!youtube!" --convert-subs vtt --geo-bypass --no-check-certificate --restrict-filenames --sleep-interval 1 --yes-playlist --ignore-errors --write-auto-sub --write-description --write-sub --skip-download "%url%"
