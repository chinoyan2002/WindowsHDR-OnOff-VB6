@echo off
powershell.exe -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0InstallDrv.ps1\"' -Verb RunAs" & REM 自動提權並以絕對路徑呼叫同目錄的腳本