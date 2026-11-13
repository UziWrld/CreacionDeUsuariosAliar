@echo off
chcp 65001 >nul
echo Creando acceso directo en el Escritorio (Version 8.2)...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell; " ^
  "$desktop = [Environment]::GetFolderPath('Desktop'); " ^
  "$s = $ws.CreateShortcut(\"$desktop\Active Directory Aliar.lnk\"); " ^
  "$s.TargetPath = 'wscript.exe'; " ^
  "$s.Arguments = \"`\"%~dp0Iniciar-AppWinForms.vbs`\"\"; " ^
  "$s.WorkingDirectory = '%~dp0'; " ^
  "$s.IconLocation = 'C:\Windows\System32\dsuiext.dll,0'; " ^
  "$s.Description = 'Aliar S.A. - Active Directory Provisioning Hub v8.2'; " ^
  "$s.Save(); " ^
  "Write-Host 'Acceso directo actualizado exitosamente en tu Escritorio.' -ForegroundColor Green"
pause
