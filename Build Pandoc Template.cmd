@echo off
setlocal
cls

"C:\Program Files\Pandoc\pandoc.exe" -o PandocTemplate.docx --print-default-data-file reference.docx
