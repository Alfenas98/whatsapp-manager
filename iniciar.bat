@echo off
chcp 65001 >nul
title WhatsApp Manager v1.0
cd /d "%~dp0"
python main.py
pause
