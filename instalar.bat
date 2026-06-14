@echo off
chcp 65001 >nul
title WhatsApp Manager - Instalacao

echo.
echo ==========================================
echo   WHATSAPP MANAGER - INSTALACAO
echo ==========================================
echo.

:: Tenta encontrar o Python em locais comuns
set PYTHON_CMD=

python --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON_CMD=python
    goto :found
)

py --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON_CMD=py
    goto :found
)

:: Procura nas pastas padrão do Windows
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
) do (
    if exist %%P (
        set PYTHON_CMD=%%P
        goto :found
    )
)

echo [ERRO] Python nao encontrado!
echo Baixe em: https://www.python.org/downloads/
echo Marque a opcao "Add Python to PATH" na instalacao.
pause
exit /b 1

:found
echo [OK] Python encontrado: %PYTHON_CMD%
echo.
echo Instalando dependencias...
echo.

%PYTHON_CMD% -m pip install playwright openpyxl rich --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)

echo.
echo Instalando navegador Chromium...
%PYTHON_CMD% -m playwright install chromium
if errorlevel 1 (
    echo [ERRO] Falha ao instalar Chromium.
    pause
    exit /b 1
)

:: Salva o caminho do Python para uso futuro
echo %PYTHON_CMD% > python_path.txt

echo.
echo ==========================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo   Execute: iniciar.bat
echo ==========================================
echo.
pause
