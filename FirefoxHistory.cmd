@echo off
setLocal enableDelayedExpansion
cls

echo Getting Firefox history
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
if exist FirefoxHistory.htm del FirefoxHistory.htm
if exist FirefoxHistory.csv del FirefoxHistory.csv
if "%days%"=="" ( "%view%" /scomma FirefoxHistory.csv /loadChrome 0 /loadEdge 0 /loadFirefox 1 /loadIE 0 /loadSafari 0 /sort "~Visit Time" ) else "%view%" /scomma FirefoxHistory.csv /loadChrome 0 /loadEdge 0 /loadFirefox 1 /loadIE 0 /loadSafari 0 /sort "~Visit Time" /VisitTimeFilter 3 /VisitTimeFilterValue %days%
rem pause
call "!regexer!" FirefoxHistory.csv csv >nul
if exist FirefoxHistory.md del FirefoxHistory.md
call "!sqlite!" ":memory:" ".headers on" ".mode csv" ".import FirefoxHistory.csv FirefoxHistory" ".mode markdown" ".once FirefoxHistory.md" "select Title, URL from FirefoxHistory where [Web Browser] == 'Firefox' order by cast(Sequence as integer);" >nul
if exist FirefoxHistory.csv del FirefoxHistory.csv
call :setSize FirefoxHistory.md
if "%size%"=="0" echo No history & goTo :eof
call "!regexer!" FirefoxHistory.md url >nul
if exist FirefoxHistory.htm del FirefoxHistory.htm
call "!pandoc2htm!" FirefoxHistory.md >nul
if exist FirefoxHistory.md del FirefoxHistory.md
start "", FirefoxHistory.htm
goTo :eof

:setSize
set size=%~z1
exit /b

