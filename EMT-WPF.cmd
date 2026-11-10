@echo off
rem Starts the WPF prototype of Exchange Message Trail. The PowerShell console stays hidden.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0ExchangeMessageTrail.Wpf.ps1" %*
