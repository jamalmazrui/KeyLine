@echo off
setLocal enableDelayedExpansion
cls

set kl=%~dp0
call "%kl%GetAttachments.cmd" list deleted %1

