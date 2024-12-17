@echo off
SetLocal EnableDelayedExpansion
cls

set kl=%~dp0
set code=%kl%code
set settings=%kl%settings

set calibre=C:\Program Files\calibre2\ebook-convert.exe
set libre=c:\program files\libreoffice\program\soffice.exe
set pandoc=C:\Program Files\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=%localAppData%\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=C:\Program Files (x86)\Pandoc\pandoc.exe
set utf8b=%code%\utf8b64.exe
set utf8n=%code%\utf8n64.exe

set ppVert=%code%\ppVert64.exe
set ppInsert=%code%\ppInsert.vbs
set wdVert=%code%\wdVert64.exe
set xlVert=%code%\xlVert64.exe

rem if not exist "%~1" echo No match & goto :eof
"%pandoc%" -o C:\AccAuthor\AccAudit_Body.pptx C:\AccAuthor\AccAudit_Body.md
cscript /nologo "%ppInsert%"
