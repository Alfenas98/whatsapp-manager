@echo off
chcp 65001 >nul
title Gerando EXE...

echo.
echo Instalando PyInstaller...
pip install pyinstaller --quiet

echo.
echo Gerando executavel...
pyinstaller --onefile --console --name "WhatsAppManager" main.py

echo.
if exist "dist\WhatsAppManager.exe" (
    echo ==========================================
    echo   EXE GERADO COM SUCESSO!
    echo   Pasta: dist\WhatsAppManager.exe
    echo ==========================================
) else (
    echo [ERRO] Falha ao gerar o executavel.
)
echo.
pause
