@echo off
setLocal enableDelayedExpansion
cls

echo Getting password history
set kl=%~dp0
set code=%kl%code
set view=%code%\WebBrowserPassView.exe
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
if exist PasswordHistory.htm del PasswordHistory.htm
if exist PasswordHistory.csv del PasswordHistory.csv
"%view%" /scomma PasswordHistory.csv /sort "~Modified Time"
rem pause
call "!regexer!" PasswordHistory.csv csv >nul
if exist PasswordHistory.md del PasswordHistory.md
call "!sqlite!" ":memory:" ".headers on" ".mode csv" ".import PasswordHistory.csv PasswordHistory" ".mode markdown" ".once PasswordHistory.md" "select URL, [User Name] as User, Password, [Web Browser] as Browser from PasswordHistory order by cast(Sequence as integer);" >nul
if exist PasswordHistory.csv del PasswordHistory.csv
call :setSize PasswordHistory.md
if "%size%"=="0" echo No history & goTo :eof
call "!regexer!" PasswordHistory.md url >nul
if exist PasswordHistory.htm del PasswordHistory.htm
call "!pandoc2htm!" PasswordHistory.md >nul
if exist PasswordHistory.md del PasswordHistory.md
start "", PasswordHistory.htm
goTo :eof

:setSize
set size=%~z1
exit /b

