@echo off
setLocal enableDelayedExpansion
cls

set youtube=%~dp0bin\youtube-dl.exe
set urllist=%~1
if "%urllist%"=="" echo No URL & goTo :eof
rem "!youtube!" --write-auto-sub --write-description --write-sub --skip-download --batch-file "%urllist%"
"!youtube!" --geo-bypass --no-check-certificate --restrict-filenames --sleep-interval 1 --yes-playlist --ignore-errors --write-auto-sub --write-description --write-sub --skip-download --batch-file "%urllist%"
