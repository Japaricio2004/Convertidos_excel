@echo off
:: ===========================
:: Script: guardar.bat
:: Autor: Japaricio2004
:: Descripción: Agrega, commitea y sube cambios automáticamente a GitHub
:: ===========================

echo.
echo 💾 Iniciando guardado automático...
echo.

:: Moverse al directorio del script
cd /d "%~dp0"

:: Verificar si Git está inicializado
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo ❌ No se encontró un repositorio Git en esta carpeta.
    pause
    exit /b
)

:: Detectar si hay cambios
for /f "delims=" %%i in ('git status --porcelain') do set CHANGES=1

if not defined CHANGES (
    echo ⚠️ No hay cambios para guardar.
    pause
    exit /b
)

:: Agregar todos los archivos
git add .

:: Crear commit con fecha y hora
set FECHA=%date:~6,4%-%date:~3,2%-%date:~0,2%
set HORA=%time:~0,2%:%time:~3,2%
set HORA=%HORA: =0%
git commit -m "Actualización automática %FECHA% %HORA%"

:: Subir cambios a GitHub
git push -u origin main

echo.
echo ✅ Cambios subidos correctamente a GitHub.
echo.
pause
