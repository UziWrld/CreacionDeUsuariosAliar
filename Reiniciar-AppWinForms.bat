@echo off
chcp 65001 >nul
title Reiniciar - Aliar S.A. Active Directory
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-Process powershell -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne $PID -and ($_.MainWindowTitle -like '*Active Directory*' -or $_.MainWindowTitle -like '*Aliar*') } | Stop-Process -Force -ErrorAction SilentlyContinue" 2>nul
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0App-CreacionUsuariosAliar.ps1"
exit
