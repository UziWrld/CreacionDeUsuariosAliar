@echo off
chcp 65001 >nul
title Aliar S.A. - Gestión Active Directory
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0App-CreacionUsuariosAliar.ps1"
exit
