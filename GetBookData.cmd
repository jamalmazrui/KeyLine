@echo off
setLocal enableDelayedExpansion
cls

set authors=%~1
set title=%~2
if "%authors%"=="" echo Pass book authors as first parameter and title as optional second parameter & goTo :eof
if "%title%"=="" goto :authors
"c:\program files\Calibre2\fetch-ebook-metadata.exe" -a "%authors%" -t "%title%"
goto :eof

:authors
"c:\program files\Calibre2\fetch-ebook-metadata.exe" -a "%authors%"


