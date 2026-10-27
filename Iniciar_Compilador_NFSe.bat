@echo off
chcp 65001 >nul
title Compilador de NFS-e (XML para Excel)
cd /d "%~dp0"
start "" pythonw compilar_nfse.py || python compilar_nfse.py
