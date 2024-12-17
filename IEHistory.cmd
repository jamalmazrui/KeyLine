@echo off
setLocal enableDelayedExpansion
cls

echo Getting IE history
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
if exist IEHistory.htm del IEHistory.htm
if exist IEHistory.csv del IEHistory.csv
if "%days%"=="" ( "%view%" /scomma IEHistory.csv /loadChrome 0 /loadEdge 0 /loadFirefox 0 /loadIE 1 /loadSafari 0 /sort "~Visit Time" ) else "%view%" /scomma IEHistory.csv /loadChrome 0 /loadEdge 0 /loadFirefox 0 /loadIE 1 /loadSafari 0 /sort "~Visit Time" /VisitTimeFilter 3 /VisitTimeFilterValue %days%
rem pause
call "!regexer!" IEHistory.csv csv >nul
if exist IEHistory.md del IEHistory.md
call "!sqlite!" ":memory:" ".headers on" ".mode csv" ".import IEHistory.csv IEHistory" ".mode markdown" ".once IEHistory.md" "select Title, URL from IEHistory where [Web Browser] == 'IE' order by cast(Sequence as integer);" >nul
if exist IEHistory.csv del IEHistory.csv
call :setSize IEHistory.md
if "%size%"=="0" echo No history & goTo :eof
call "!regexer!" IEHistory.md url >nul
if exist IEHistory.htm del IEHistory.htm
call "!pandoc2htm!" IEHistory.md >nul
if exist IEHistory.md del IEHistory.md
start "", IEHistory.htm
goTo :eof

:setSize
set size=%~z1
exit /b

