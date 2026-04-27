@echo off
REM ====================================================
REM   FINANZAS FAMILIA - Publicar TODO
REM ====================================================
REM
REM Sube a GitHub cualquier cosa modificada:
REM   - Cambios manuales del Excel (data/finanzas.json)
REM   - Cambios del codigo del dashboard (HTML/CSS/JS)
REM   - Cambios de scripts de Python
REM
REM Despues de 1-2 min Cloudflare redespliega solo.
REM ====================================================

cd /d "%~dp0"
echo.
echo ============================================
echo   FINANZAS - Publicando todo
echo ============================================
echo.

echo [1/2] Regenerando JSON desde el Excel...
python scripts\actualizar.py --solo-export
if errorlevel 1 (
    echo ERROR al regenerar JSON.
    pause
    exit /b 1
)

echo.
echo [2/2] Subiendo a GitHub...
git add -A
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Actualizacion %date% %time%" >nul
    git push
    if errorlevel 1 (
        echo  !! ERROR al hacer push.
    ) else (
        echo  OK Subido. Cloudflare redespliega en 1-2 min.
        echo  Link: https://finanzas-familia.pages.dev/
    )
) else (
    echo  Sin cambios para publicar.
)

echo.
echo ============================================
pause
