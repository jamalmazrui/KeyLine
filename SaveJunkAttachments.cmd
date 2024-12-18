@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
call "%kl%GetAttachments.cmd" save Junk %1

