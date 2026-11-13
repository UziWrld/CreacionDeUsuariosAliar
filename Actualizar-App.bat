@echo off
title Actualizar Herramienta Active Directory - Aliar S.A.
cd /d "%~dp0"
echo ======================================================================
echo  Consultando ultimas actualizaciones en GitHub...
echo ======================================================================
echo.
git pull origin main
echo.
if %errorlevel% equ 0 (
    echo ======================================================================
    echo  [EXITO] La herramienta esta completamente actualizada.
    echo ======================================================================
) else (
    echo  [AVISO] No se pudo completar la actualizacion. Revisa tu conexion.
)
echo.
pause
