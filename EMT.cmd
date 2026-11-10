@echo off
rem Starts Exchange Message Trail in STA mode with a minimized console window.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Minimized -File "%~dp0ExchangeMessageTrail.ps1" %*
