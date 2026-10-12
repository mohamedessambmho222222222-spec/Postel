@echo off
chcp 65001 >nul
cd /d "%~dp0"
title تشغيل أداة تسجيل البوستة
color 0F

echo ============================================================
echo     تشغيل أداة تسجيل البوستة
echo ============================================================
echo.
echo   الأداة بتشتغل محلي على جهازك فقط (مفيش أي بيانات بتخرج).
echo   المتصفح هيفتح لوحده على:  http://127.0.0.1:8765
echo.
echo   * لو الصفحة مفتوحة خلاص: افتح الرابط ده في المتصفح وخلاص.
echo   * لإيقاف الأداة: اقفل النافذة دي.
echo.

set "PY="
where python >nul 2>nul
if not errorlevel 1 set "PY=python"

if not defined PY (
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)

if not defined PY (
  echo.
  echo [!] بايثون مش موجود.
  echo     شغّل التثبيت مرة واحدة: دوس على  تثبيت.bat
  echo.
  pause
  exit /b 1
)

%PY% postal_ocr.py
echo.
echo تم إيقاف الأداة.
pause
