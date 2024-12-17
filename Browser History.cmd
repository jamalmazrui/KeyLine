@echo off
setLocal enableDelayedExpansion
cls

echo Getting browser history
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
if exist BrowserHistory.htm del BrowserHistory.htm
if exist BrowserHistory.csv del BrowserHistory.csv
if "%days%"=="" ( "%view%" /scomma BrowserHistory.csv /sort "~Visit Time" ) else "%view%" /scomma BrowserHistory.csv /sort "~Visit Time" /VisitTimeFilter 3 /VisitTimeFilterValue %days%
rem pause
call "!regexer!" BrowserHistory.csv csv >nul
if exist BrowserHistory.md del BrowserHistory.md
call "!sqlite!" ":memory:" ".headers on" ".mode csv" ".import BrowserHistory.csv BrowserHistory" ".mode markdown" ".once BrowserHistory.md" "select Title, URL, [Web Browser] as Browser from BrowserHistory order by cast(Sequence as integer);" >nul
if exist BrowserHistory.csv del BrowserHistory.csv
call :setSize BrowserHistory.md
if "%size%"=="0" echo No history & goTo :eof
call "!regexer!" BrowserHistory.md url >nul
if exist BrowserHistory.htm del BrowserHistory.htm
call "!pandoc2htm!" BrowserHistory.md >nul
if exist BrowserHistory.md del BrowserHistory.md
start "", BrowserHistory.htm
goTo :eof

:setSize
set size=%~z1
exit /b

