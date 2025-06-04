@echo off
cd /d %~dp0
pythonw extrair_chaves.py
timeout /t 5 >nul
start "" "chaves_extraidas.xlsx"
