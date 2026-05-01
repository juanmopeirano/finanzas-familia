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
    REM Auto-responder "n" a los prompts de cleanup (OneDrive locks)
    (for /L %%i in (1,1,300) do @echo n) | git push 2>&1 | findstr /V "Deletion of directory"
    if errorlevel 1 (
        echo  !! Verifica con git status si subio bien.
    ) else (
        echo  OK Dashboard online actualizado
        echo  Link: https://finanzas-familia.pages.dev/
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
