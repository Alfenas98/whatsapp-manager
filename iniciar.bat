@echo off
chcp 65001 >nul
title WhatsApp Manager v1.0
cd /d "%~dp0"

:: Tenta usar o Python salvo na instalação
set PYTHON_CMD=
if exist python_path.txt (
    set /p PYTHON_CMD=<python_path.txt
)

:: Se não encontrou, tenta os padrões
if "%PYTHON_CMD%"=="" (
    python --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=python
)

if "%PYTHON_CMD%"=="" (
    py --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=py
)

if "%PYTHON_CMD%"=="" (
    echo [ERRO] Python nao encontrado. Execute instalar.bat primeiro.
    pause
    exit /b 1
)

%PYTHON_CMD% main.py
pause
