@echo off
rem Envia el informe comercial por vendedor (HTML) por Outlook. Requiere Outlook abierto.
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0enviar_comercial.py" >> "%~dp0..\data\envio_comercial.log" 2>&1
