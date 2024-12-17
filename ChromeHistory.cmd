@echo off
setLocal enableDelayedExpansion
cls

echo Getting Chrome history
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
if exist ChromeHistory.htm del ChromeHistory.htm
if exist ChromeHistory.csv del ChromeHistory.csv
if "%days%"=="" ( "%view%" /scomma ChromeHistory.csv /loadChrome 1 /loadEdge 0 /loadFirefox 0 /loadIE 0 /loadSafari 0 /sort "~Visit Time" ) else "%view%" /scomma ChromeHistory.csv /loadChrome 1 /loadEdge 0 /loadFirefox 0 /loadIE 0 /loadSafari 0 /sort "~Visit Time" /VisitTimeFilter 3 /VisitTimeFilterValue %days%
rem pause
call "!regexer!" ChromeHistory.csv csv >nul
if exist ChromeHistory.md del ChromeHistory.md
call "!sqlite!" ":memory:" ".headers on" ".mode csv" ".import ChromeHistory.csv ChromeHistory" ".mode markdown" ".once ChromeHistory.md" "select Title, URL from ChromeHistory where [Web Browser] == 'Chrome' order by cast(Sequence as integer);" >nul
if exist ChromeHistory.csv del ChromeHistory.csv
call :setSize ChromeHistory.md
if "%size%"=="0" echo No history & goTo :eof
call "!regexer!" ChromeHistory.md url >nul
if exist ChromeHistory.htm del ChromeHistory.htm
call "!pandoc2htm!" ChromeHistory.md >nul
if exist ChromeHistory.md del ChromeHistory.md
start "", ChromeHistory.htm
goTo :eof

:setSize
set size=%~z1
exit /b

