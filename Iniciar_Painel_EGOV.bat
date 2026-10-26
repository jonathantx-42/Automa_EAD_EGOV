@echo off
title Central de Automacoes EAD - EGOV
cd /d "%~dp0"
python app.py
if errorlevel 1 (
    echo.
    echo Ocorreu um erro ao iniciar a aplicacao.
    pause
)
