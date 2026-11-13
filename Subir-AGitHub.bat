@echo off
title Subir Cambios a GitHub - Aliar S.A.
cd /d "%~dp0"
echo ======================================================================
echo  Subiendo repositorio a https://github.com/UziWrld/CreacionDeUsuariosAliar.git
echo ======================================================================
echo.
git push -u origin main
echo.
if %errorlevel% equ 0 (
    echo ======================================================================
    echo  [EXITO] Cambios subidos exitosamente a GitHub!
    echo ======================================================================
) else (
    echo  [ERROR] Ocurrio un error al subir los cambios a GitHub.
)
echo.
pause
