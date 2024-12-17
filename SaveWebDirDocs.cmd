@echo off
setLocal enableDelayedExpansion
cls

set httrack=c:\program files\WinHTTrack\httrack.exe
set wget=%~dp0bin\wget.exe

rem "!httrack!" "%~1" -o0 -s0 %2 %3 %4 %5 %6 %7 %8 %9
rem "!httrack!" "%~1" -W %2 %3 %4 %5 %6 %7 %8 %9
rem "!httrack!" "%~1" -w -D -s0 %2 %3 %4 %5 %6 %7 %8 %9

rem ‘--config=FILE’
rem -A | --accept
rem --content-disposition
rem --convert-links
rem -E | --adjust-extension
rem -e |--execute robots=off  
rem --follow-ftp
rem -l | -- level
rem --metalink-over-http
rem -nc | --no-clobber
rem --no-check-certificate
rem -np | --no-parent
rem -nv | --no-verbose
rem -r | --recursive
rem -R | --reject
rem --trust-server-names
rem -w | --wait

set url=%~1
set accept=%~2
set reject=%~3

rem Try to download only documents and archives
set accept=*
set reject=css,exe,gif,htm,html,ico,icon,ini,jpeg,jpg,js,md,m4v,mp3,mp4,mpeg,mpg,msi,png,svg,wav,wmf,wmv,woff,woff2,xhtml,xml

if "%url%"=="" echo No URL & goTo :eof
if "%accept%"=="" goTo :url
if "%reject%"=="" goTo :accept
"!wget!" -A "%accept%" -R "%reject%" --content-disposition -E -e robots=off --ignore-case -l -nc --no-check-certificate -np -nv -r --trust-server-names -w2 "%url%" & goTo :eof

:url
"!wget!" --content-disposition -E -e robots=off --ignore-case -l -nc --no-check-certificate -np -nv -r --trust-server-names -w2 "%url%" & goTo :eof
goTo :eof

:accept
"!wget!" -A "%accept%" --content-disposition -E -e robots=off --ignore-case -l -nc --no-check-certificate -np -nv -r --trust-server-names -w2 "%url%" & goTo :eof
goTo :eof
