@echo off
setlocal EnableExtensions
title ACTUALIZAR Y PUBLICAR DASHBOARD - SINCRONIZACION GITHUB

cd /d "%~dp0"

echo ============================================================
echo   ACTUALIZAR Y PUBLICAR DASHBOARD DE GRUAS
echo   SINCRONIZACION CON GITHUB
echo ============================================================
echo.

if not exist ".git" (
    echo ERROR: No se encontro la carpeta .git.
    pause
    exit /b 1
)

if not exist "Actualizar_dashboard.py" (
    echo ERROR: No se encontro Actualizar_dashboard.py.
    pause
    exit /b 1
)

if not exist "DATA GRUAS.xlsx" (
    echo ERROR: No se encontro DATA GRUAS.xlsx.
    pause
    exit /b 1
)

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON=py"
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON=python"
    ) else (
        echo ERROR: No se encontro Python.
        pause
        exit /b 1
    )
)

echo [1/6] Descargando cambios actuales de GitHub...
git fetch origin
if errorlevel 1 (
    echo ERROR: No se pudo hacer git fetch.
    pause
    exit /b 1
)

echo.
echo [2/6] Sincronizando la rama local con GitHub...
git pull --rebase --autostash origin main
if errorlevel 1 (
    echo.
    echo ============================================================
    echo ERROR: Git encontro un conflicto durante la sincronizacion.
    echo NO se continuara para evitar perder informacion.
    echo ============================================================
    echo.
    echo Ejecuta:
    echo     git status
    echo.
    echo y envia una captura para resolver el conflicto.
    pause
    exit /b 1
)

echo.
echo [3/6] Actualizando dashboard desde DATA GRUAS.xlsx...
%PYTHON% Actualizar_dashboard.py
if errorlevel 1 (
    echo ERROR: Actualizar_dashboard.py fallo.
    pause
    exit /b 1
)

echo.
echo [4/6] Preparando cambios...
git add index.html Actualizar_dashboard.py

git diff --cached --quiet
if %errorlevel%==0 (
    echo No hay cambios nuevos.
    echo El dashboard local ya coincide con la version publicada.
    echo.
    echo URL:
    echo https://ydrogo-lozano-ux.github.io/dashboard-gruas/
    pause
    exit /b 0
)

git commit -m "Actualizar dashboard"

if errorlevel 1 (
    echo ERROR: No se pudo crear el commit.
    pause
    exit /b 1
)

echo.
echo [5/6] Verificando nuevamente GitHub antes del push...
git fetch origin
if errorlevel 1 (
    echo ERROR: No se pudo consultar GitHub.
    pause
    exit /b 1
)

git pull --rebase --autostash origin main
if errorlevel 1 (
    echo.
    echo ERROR: Se detecto un conflicto antes de publicar.
    echo No se sobrescribira GitHub automaticamente.
    pause
    exit /b 1
)

echo.
echo [6/6] Publicando en GitHub...
git push origin main

if errorlevel 1 (
    echo.
    echo ============================================================
    echo ERROR: git push fallo.
    echo ============================================================
    echo.
    echo Ejecuta este BAT nuevamente.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   PUBLICACION COMPLETADA CORRECTAMENTE
echo ============================================================
echo.
echo Commit publicado:
git rev-parse --short HEAD
echo.
echo URL:
echo https://ydrogo-lozano-ux.github.io/dashboard-gruas/
echo.
echo Espera 30-90 segundos y luego usa CTRL+F5.
echo ============================================================
pause
