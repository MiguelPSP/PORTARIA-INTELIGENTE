@echo off
title Sistema da Portaria - Interface Digital
cd /d "%~dp0"
echo Iniciando o sistema da portaria...
streamlit run app.py
pause