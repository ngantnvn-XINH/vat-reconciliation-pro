@echo off
chcp 65001 >nul
title Vat Reconciliation Pro - Doi Chieu Thue GTGT
echo Đang khởi động phần mềm Đối Chiếu Thuế GTGT...
python vat_reconciler_gui.py
if %errorlevel% neq 0 (
    echo.
    echo [Lỗi] Không thể khởi động ứng dụng GUI!
    echo Vui lòng kiểm tra Python và cài đặt openpyxl bằng lệnh: pip install openpyxl
    pause
)
