@echo off
title El Cerebro - Panel local
echo Iniciando El Cerebro...
echo Abri en el navegador:  http://127.0.0.1:5000
echo (Para cerrar el panel, cerra esta ventana)
start "" http://127.0.0.1:5000
"python" "%~dp0app.py"
pause
