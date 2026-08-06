@echo off
REM ============================================================
REM  ACTUALIZAR DASHBOARD SOHO (Winback + Cross-selling + Metas)
REM  Doble clic: re-ejecuta el pipeline y abre el dashboard.
REM  Requisito de la reunion: "que se actualice solita"
REM ============================================================
cd /d "%~dp0"
echo.
echo  Actualizando dashboard SOHO con los datos mas recientes...
echo  (esto puede tardar 1-3 minutos)
echo.
python dash_prod\extraer_datos.py
if errorlevel 1 goto :error
echo.
echo  ✔ Dashboard actualizado.
start "" "dash_prod\index.html"
exit /b 0

:error
echo.
echo  [ERROR] No se pudo actualizar. Revisa el mensaje de arriba.
echo  Causas comunes: Python no instalado, o el archivo de reporte
echo  (Reporte_soho_junio_25 (1).xlsb) no esta en Amir/soho_prod/.
pause
exit /b 1
