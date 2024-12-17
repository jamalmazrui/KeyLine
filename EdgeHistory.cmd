@echo off
setLocal enableDelayedExpansion
cls

echo Getting Edge history
set kl=%~dp0
set code=%kl%code
set view=%code%\BrowsingHistoryView.exe
set xl2csv=%kl%xl2csv.cmd
set sqlite=%kl%sqlite.cmd
set regexer=%kl%regexer.cmd
set pandoc2htm=%kl%pandoc2htm.cmd

set pandoc=C:\Program Files\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=%localAppData%\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=C:\Program Files (x86)\Pandoc\pandoc.exe
set utf8b=%code%\utf8b64.exe
set utf8n=%code%\utf8n64.exe
set csvcut=%code%\csvcut.exe

set days=%1
if exist EdgeHistory.htm del EdgeHistory.htm
if exist EdgeHistory.csv del EdgeHistory.csv
if "%days%"=="" ( "%view%" /scomma EdgeHistory.csv /loadChrome 0 /loadEdge 1 /loadFirefox 0 /loadIE 0 /loadSafari 0 /sort "~Visit Time" ) else "%view%" /scomma EdgeHistory.csv /loadChrome 0 /loadEdge 1 /loadFirefox 0 /loadIE 0 /loadSafari 0 /sort "~Visit Time" /VisitTimeFilter 3 /VisitTimeFilterValue %days%
rem pause
call "!regexer!" EdgeHistory.csv csv >nul
if exist EdgeHistory.md del EdgeHistory.md
call "!sqlite!" ":memory:" ".headers on" ".mode csv" ".import EdgeHistory.csv EdgeHistory" ".mode markdown" ".once EdgeHistory.md" "select Title, URL from EdgeHistory where [Web Browser] == 'Edge' order by cast(Sequence as integer);" >nul
if exist EdgeHistory.csv del EdgeHistory.csv
call :setSize EdgeHistory.md
if "%size%"=="0" echo No history & goTo :eof
call "!regexer!" EdgeHistory.md url >nul
if exist EdgeHistory.htm del EdgeHistory.htm
call "!pandoc2htm!" EdgeHistory.md >nul
if exist EdgeHistory.md del EdgeHistory.md
start "", EdgeHistory.htm
goTo :eof

:setSize
set size=%~z1
exit /b

