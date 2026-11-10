@echo off
rem Starts MailCleaner in STA mode with a minimized console window.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Minimized -File "%~dp0MailCleaner.ps1" %*
