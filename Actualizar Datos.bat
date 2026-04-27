@echo off
REM ====================================================
REM   FINANZAS FAMILIA - Actualizar todo
REM ====================================================
REM
REM Lo que hace:
REM   1. Procesa los archivos que esten en 'NUEVOS del banco/'
REM   2. Auto-clasifica los movimientos nuevos
REM   3. Mueve los archivos procesados a 'PROCESADOS/'
REM   4. Regenera el JSON del dashboard
REM   5. Sube todo a GitHub Pages
REM   6. Te abre el Excel si hay sin clasificar para revisar
REM ====================================================

cd /d "%~dp0"
echo.
echo ============================================
echo   FINANZAS FAMILIA - Actualizando
echo ============================================
echo.

python scripts\actualizar.py
set ERR=%errorlevel%

if %ERR% geq 2 (
    echo.
    echo ERROR en el procesamiento.
    pause
    exit /b 1
)

REM Push a GitHub
echo.
echo --- Publicando en GitHub Pages ---
git add -A
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Actualizacion %date% %time%" >nul
    git push
    if errorlevel 1 (
        echo  !! ERROR al hacer push. Revisa la conexion.
    ) else (
        echo  OK Dashboard online actualizado
        echo  Link: https://juanmopeirano.github.io/finanzas-familia/
    )
) else (
    echo  Sin cambios para publicar.
)

echo.
echo ============================================

REM Si hay sin clasificar (exit 1), abrir el Excel
if %ERR% equ 1 (
    echo.
    echo  ATENCION: hay movimientos sin clasificar
    echo  Te abro el Excel para que los revises.
    echo  Despues de corregir, volve a ejecutar este .bat
    timeout /t 3 >nul
    start "" "C:\Users\jmpei\OneDrive\Finanzas Pili JM\Planilla Familia.xlsx"
)

echo.
pause
