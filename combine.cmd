@echo off
setLocal enableDelayedExpansion
cls

set targetRoot=%~2
if "%targetRoot%"=="" echo Pass a file spec as the first parameter and a target root name as the second parameter & goTo :eof
set spec=%~1
if "%spec%"=="" set spec=*.htm

if not exist "%spec%" echo No match & goTo :eof

echo Combining
rem echo 10
set pandoc=C:\Program Files\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=%localAppData%\Pandoc\pandoc.exe
if not exist "%pandoc%" set pandoc=C:\Program Files (x86)\Pandoc\pandoc.exe

set kl=%~dp0
set settings=%kl%settings
set code=%kl%code
set utf8b=%code%\utf8b64.exe
set regexer=%code%\regexer.exe

set sourceDir=%cd%
pushd "%temp%"
If %errorlevel% NEQ 0 echo Error & goTo :eof

rem if exist "Pax\*.md" del /q "Pax\*.md"
if exist "%temp%\Pax\*" del /q "%temp%\Pax\*"
if not exist "Pax" md Pax
cd Pax
rem call "!kl!htm2md" "%sourceDir%\%spec%" >nul
rem call c:\code\pandoc_htm2md.bat "%sourceDir%\%spec%" >nul
for %%f in ("%sourceDir%\%spec%") do "!utf8b!" "%%f" >nul
for %%f in ("%sourceDir%\%spec%") do "!regexer!" "%%f" "%settings%\no_images-Regexer.settings" >nul
for %%f in ("%sourceDir%\%spec%") do "!regexer!" "%%f" "%settings%\h1-Regexer.settings" >nul
for %%f in ("%sourceDir%\%spec%") do "!pandoc!" --quiet --markdown-headings=atx -f html -t markdown_strict --wrap=none -o "%%~nf.md" "%%f" >nul
rem echo 20
rem copy *.md "%targetRoot%.md" >nul
copy *.md /a "%targetRoot%.md" >nul
rem echo 30
rem call c:\code\fixHeadings.bat "%targetRoot%.md" >nul
"!regexer!" "%targetRoot%.md" "%settings%\fixHeadings-Regexer.settings" >nul
rem echo 40
rem call c:\code\pandoc_md2htm_toc.bat "%targetRoot%.md" >nul
"!pandoc!" --quiet -s --toc --toc-depth=2 -f markdown -t html -o "%targetRoot%.htm" "%targetRoot%.md" > nul
rem echo 50
move "%targetRoot%.htm" "%sourceDir%" >nul
rem echo 60
del /q *.md 
cd ..
rd Pax
popd
if exist "%sourceDir%\%targetRoot%.htm" (echo Done & start "", "%targetRoot%.htm") else echo Error
