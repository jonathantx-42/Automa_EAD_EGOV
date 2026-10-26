@echo off
title Instalador de Requisitos - EGOV
cd /d "%~dp0"
echo ============================================================
echo   INSTALANDO REQUISITOS DA CENTRAL DE AUTOMACAO EGOV
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] O Python nao foi encontrado no sistema!
    echo Certifique-se de instalar o Python e marcar "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo 1. Instalando bibliotecas Python (pandas, playwright, customtkinter, etc)...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao instalar algumas bibliotecas.
    pause
    exit /b 1
)

echo.
echo 2. Instalando o navegador do Playwright (Chromium)...
python -m playwright install chromium
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao instalar o navegador do Playwright.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   TUDO PRONTO! O computador esta configurado com sucesso.
echo   Basta executar o arquivo 'Iniciar_Painel_EGOV.bat'.
echo ============================================================
echo.
pause
