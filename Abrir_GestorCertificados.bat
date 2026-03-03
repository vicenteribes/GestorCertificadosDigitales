@echo off
:: Lanzador para GestorCertificados.ps1
:: Coloca este .bat en la misma carpeta que el .ps1

cd /d "%~dp0"

:: Intentar establecer política de ejecución (por si no se ha hecho)
powershell -Command "Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force" >nul 2>&1

:: Ejecutar el script
powershell -ExecutionPolicy Bypass -File "%~dp0GestorCertificados.ps1"

:: Si hubo error, mostrar mensaje y no cerrar
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ============================================
    echo  ERROR al ejecutar el script.
    echo  Codigo de error: %ERRORLEVEL%
    echo ============================================
    echo.
    pause
)
