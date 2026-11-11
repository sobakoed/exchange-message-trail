@echo off
rem Starts the classic window of Exchange Message Trail (Windows Forms). Use EMT.cmd for the usual window.
rem The classic window stays for one version as a fallback.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0ExchangeMessageTrail.Classic.ps1" %*
