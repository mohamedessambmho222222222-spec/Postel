@echo off
chcp 65001 >nul
cd /d "%~dp0"
title فحص أداة تسجيل البوستة
color 0E
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0health.ps1"
echo.
pause