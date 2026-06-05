@echo off
title Cargando Dashboard Gestión Creativa II
echo Iniciando el sistema de gestion...

:: %~dp0 extrae automáticamente la ruta de la carpeta donde está este archivo .bat
cd /d "%~dp0"

:: Ejecuta streamlit sobre el archivo app.py que está en esta misma carpeta
streamlit run app.py

pause