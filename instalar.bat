@echo off
chcp 65001 >nul
title WhatsApp Manager - Instalacao

echo.
echo ==========================================
echo   WHATSAPP MANAGER - INSTALACAO
echo ==========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado!
    echo Baixe em: https://www.python.org/downloads/
    echo Marque a opcao "Add Python to PATH" na instalacao.
    pause
    exit /b 1
)

echo [OK] Python encontrado.
echo.
echo Instalando dependencias...
echo.

pip install playwright openpyxl rich --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)

echo.
echo Instalando navegador Chromium...
python -m playwright install chromium
if errorlevel 1 (
    echo [ERRO] Falha ao instalar Chromium.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo   Execute: iniciar.bat
echo ==========================================
echo.
pause
