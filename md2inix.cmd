@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings
set calibre=C:\Program Files\calibre2\ebook-convert.exe
set pandoc=C:\Program Files\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=%localAppData%\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=C:\Program Files (x86)\Pandoc\pandoc.exe
set ppVert=%code%\ppVert64.exe
set xlVert=%code%\xlVert64.exe
set wdVert=%code%\wdVert64.exe
set utf8b=%code%\utf8b.exe

set spec=%~1
if "%spec%"=="" set spec=*.md
if not exist "%spec%" echo No match & goto :eof

call "!kl!pandoc_wd2inix" "%spec%"
