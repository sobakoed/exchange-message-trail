@echo off
rem Starts Exchange Message Trail in STA mode. The console window stays hidden.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0ExchangeMessageTrail.ps1" %*
