@echo off
setLocal enableDelayedExpansion
cls

set package=%~1
if "%package%"=="" echo Pass a package name & goTo :eof

start "", "http://media.readthedocs.org/pdf/%package%/latest/%package%.pdf"

