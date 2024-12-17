@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
set fileDir=c:\program files (x86)\FileDir\FileDir.exe
set fileDir=c:\FileDir\FileDir.exe
call "!kl!SetPath.cmd"
C:
cd \Pax\work
if exist "!fileDir!" ("!fileDir!" C:\Pax\work) else start "", C:\Pax\work
rem if exist "!fileDir!" (start "!fileDir!" C:\Pax\work) else start "", C:\Pax\work
cd \Pax\work
