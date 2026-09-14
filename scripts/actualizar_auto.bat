@echo off
rem Actualizacion automatica de la casilla oc@ (programada 8:00/10:00/12:00/16:00)
rem Requiere: sesion de Windows abierta y Outlook corriendo (lectura por COM).
cd /d "%~dp0.."
set PY=python
echo. >> data\actualizacion_auto.log
echo ===== %date% %time% ===== >> data\actualizacion_auto.log
"%PY%" scripts\dump_oc.py >> data\actualizacion_auto.log 2>&1
"%PY%" scripts\extractor.py >> data\actualizacion_auto.log 2>&1
"%PY%" scripts\informe_clientes.py >> data\actualizacion_auto.log 2>&1
