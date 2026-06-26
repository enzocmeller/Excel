@echo off
setlocal
title USA Exports - Update, Refresh and PDF
echo Updating data, refreshing Excel, and building the combined PDF...
echo (Make sure USA_Exports.xlsx is CLOSED before continuing.)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_report.ps1"
echo.
pause
