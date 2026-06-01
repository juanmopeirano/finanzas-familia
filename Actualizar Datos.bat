@echo off
REM ====================================================
REM   FINANZAS FAMILIA - Actualizar datos (local)
REM ====================================================
REM
REM Lo que hace:
REM   1. Procesa los archivos que esten en 'NUEVOS del banco/'
REM   2. Auto-clasifica los movimientos nuevos
REM   3. Mueve los archivos procesados a 'PROCESADOS/'
REM   4. Regenera el JSON del dashboard (LOCAL)
REM   5. Te abre el Excel si hay sin clasificar para revisar
REM
REM   NO sube nada a la web. Para publicar usa "Publicar Todo.bat".
REM ====================================================

cd /d "%~dp0"
echo.
echo ============================================
echo   FINANZAS FAMILIA - Actualizando datos
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

echo.
echo ============================================
echo   Datos actualizados localmente.
echo   Para publicar a la web usa "Publicar Todo.bat".
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
