@echo off
cd /d "%~dp0"
title VAT Reconciliation Pro - Doi Chieu Thue GTGT
echo ===================================================
echo   KHOI DONG PHAN MEM DOI CHIEU THUE GTGT PRO
echo ===================================================
echo.

python vat_reconciler_gui.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LOI] Khong the khoi chay bang lenh python mac dinh.
    echo Dang thu khoi chay bang Python 3.12...
    "C:\Users\Admin\AppData\Local\Programs\Python\Python312\python.exe" vat_reconciler_gui.py
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [THAT BAI] Khong the khoi dong ung dung GUI.
        pause
    )
)
