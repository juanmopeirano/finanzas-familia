@echo off
REM Levanta un servidor local y abre el dashboard en el navegador.
REM Cerrar esta ventana negra apaga el servidor.

cd /d "%~dp0"
echo.
echo ========================================
echo   Finanzas JM ^& Pili - Dashboard local
echo ========================================
echo.
echo Abriendo http://localhost:8765 en el navegador...
echo (Mantene esta ventana abierta mientras uses el dashboard)
echo.

start "" "http://localhost:8765"
python -m http.server 8765
pause
