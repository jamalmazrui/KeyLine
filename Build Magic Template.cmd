@echo off
setlocal
call C:\Pax\SetPath.cmd
call KillWinWord.cmd
cls

if exist MagicTemplate.docx del MagicTemplate.docx
echo Copying Pandoc template
copy "%~dp0ini\PandocTemplate.docx" MagicTemplate.docx
rem pause

echo Copying Python styles
copy "%~dp0ini\Python-docxTemplate.docx" PythonTemplate.docx
call DocxIt.cmd PythonTemplate.docx CopySourceStyles-DocxIt.settings
rem pause

echo Creating Magic styles
call DocxStyles.cmd MagicTemplate.docx C:\Pax\settings\MagicTemplate-DocxStyles.settings
rem call "%~dp0SetPythonPath.cmd"
rem python.exe "%~dp0bin\MagicTemplate.py"
rem cscript.exe /nologo "%~dp0bin\MagicTemplate.vbs"
