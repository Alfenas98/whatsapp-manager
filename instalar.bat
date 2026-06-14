@echo off
chcp 65001 >nul
title WhatsApp Manager - Instalação

echo.
echo ╔══════════════════════════════════════════╗
echo ║     WHATSAPP MANAGER - INSTALAÇÃO        ║
echo ╚══════════════════════════════════════════╝
echo.

:: Verifica Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python não encontrado!
    echo Baixe em: https://www.python.org/downloads/
    echo Marque a opção "Add Python to PATH" na instalação.
    pause
    exit /b 1
)

echo [OK] Python encontrado.
echo.
echo Instalando dependências...
echo.

pip install playwright openpyxl rich --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependências.
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
echo ╔══════════════════════════════════════════╗
echo ║     INSTALAÇÃO CONCLUÍDA COM SUCESSO!    ║
echo ║     Execute: iniciar.bat                 ║
echo ╚══════════════════════════════════════════╝
echo.
pause
