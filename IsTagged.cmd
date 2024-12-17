@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
set isTagged=%kl%code\IsTagged.exe
set utf8b=%kl%code\utf8b64.exe

set spec=%~1
if "%spec%"=="" set spec=*.pdf
if not exist "%spec%" echo No match & goTo :eof

"!isTagged!" "%spec%" >TAGGED.txt
type TAGGED.txt
"!utf8b!" TAGGED.txt >nul

