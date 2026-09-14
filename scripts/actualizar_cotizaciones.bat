@echo off
rem Extraccion diaria de la casilla cotizaciones@ (programada 16:30).
rem Requiere: sesion de Windows abierta y Outlook corriendo (lectura por COM).
cd /d "%~dp0.."
set PY=python
echo. >> data\cotizaciones_auto.log
echo ===== %date% %time% ===== >> data\cotizaciones_auto.log
"%PY%" scripts\dump_cotizaciones.py >> data\cotizaciones_auto.log 2>&1
"%PY%" scripts\extractor_cotizaciones.py >> data\cotizaciones_auto.log 2>&1
