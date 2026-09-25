@echo off
setlocal
cd /d "%~dp0"

echo ==============================================
echo   DASHBOARD DE GRUAS - ACTUALIZAR Y PUBLICAR
echo ==============================================
echo.

set PY=
where py >nul 2>nul
if %errorlevel%==0 set PY=py
if not defined PY (
  where python >nul 2>nul
  if %errorlevel%==0 set PY=python
)
if not defined PY (
  echo ERROR: Python no esta instalado o no esta en PATH.
  pause
  exit /b 1
)

echo [1/3] Actualizando dashboard desde DATA GRUAS.xlsx...
%PY% Actualizar_dashboard.py
if errorlevel 1 (
  echo.
  echo ERROR: No se pudo actualizar el dashboard.
  pause
  exit /b 1
)

echo.
echo [2/3] Guardando cambios en Git...
git add index.html

git diff --cached --quiet
if %errorlevel%==0 (
  echo No hay cambios nuevos para publicar.
) else (
  git commit -m "Actualizar dashboard - %date% %time%"
  if errorlevel 1 (
    echo ERROR al crear el commit.
    pause
    exit /b 1
  )
  echo.
  echo [3/3] Publicando en GitHub...
  git push origin main
  if errorlevel 1 (
    echo ERROR al hacer push. Verifica GitHub y tus credenciales.
    pause
    exit /b 1
  )
  echo.
  echo PUBLICACION ENVIADA CORRECTAMENTE.
  echo GitHub Actions tardara normalmente unos segundos/minutos en actualizar la web.
)

echo.
echo Dashboard local: %~dp0index.html
echo.
pause
