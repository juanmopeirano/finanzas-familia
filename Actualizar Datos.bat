@echo off
REM Pipeline completo: importa archivo del banco (si arrastras uno encima),
REM clasifica con ML y regenera el JSON del dashboard.
REM
REM USO 1: doble click - solo regenera el JSON con lo que ya esta en el Excel.
REM USO 2: arrastrar el archivo del banco encima de este .bat para importarlo.

cd /d "%~dp0"
echo.
echo ========================================
echo   Actualizando datos
echo ========================================
echo.

if "%~1"=="" (
    echo Sin archivo del banco - solo regenero JSON.
    python scripts\actualizar.py --solo-export
) else (
    echo Importando: %~1
    python scripts\actualizar.py "%~1"
)

echo.
echo ========================================
echo   Listo. Cerra esta ventana cuando quieras.
echo ========================================
pause
