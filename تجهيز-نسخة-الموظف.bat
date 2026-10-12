@echo off
chcp 65001 >nul
cd /d "%~dp0"
title تجهيز نسخة الموظف المرشحة للمراجعة
color 0A

echo ============================================================
echo     تجهيز نسخة الموظف - تسجيل البوستة
echo ============================================================
echo.
echo   ده بيعمل نسخة نضيفة للموظف في:  Desktop\PostalTool-Employee
echo   النسخة فيها: البرنامج + النموذج + السجل + 3 ازرار فقط
echo   ومفيهاش: ملفات الاختبار + النسخ المؤقتة + Skills المطور
echo.

set "DEST=%USERPROFILE%\Desktop\PostalTool-Employee"
if exist "%DEST%" (
  echo [!] المجلد موجود - تجهيز نسخة جديدة بجواره.
  set "DEST=%USERPROFILE%\Desktop\PostalTool-Employee-%RANDOM%-%RANDOM%"
)
mkdir "%DEST%"

echo [1/6] نسخ البرنامج الاساسي...
copy /y postal_ocr.py "%DEST%\" >nul
mkdir "%DEST%\ui"
copy /y ui\index.html "%DEST%\ui\" >nul
copy /y ui\postal.css "%DEST%\ui\" >nul
copy /y ui\postal.js "%DEST%\ui\" >nul
copy /y scan_wia.ps1 "%DEST%\" >nul
copy /y requirements.txt "%DEST%\" >nul
copy /y models_dl.py "%DEST%\" >nul

echo [2/6] نسخ النموذج العربي (8 ميجا)...
mkdir "%DEST%\models"
copy /y models\arabic_PP-OCRv5_rec_mobile.onnx "%DEST%\models\" >nul
copy /y models\ppocrv5_arabic_dict.txt "%DEST%\models\" >nul

echo [3/6] نسخ السجل الرسمي والشعار...
copy /y "البوستة.xlsx" "%DEST%\" >nul
copy /y "شركة_ميناء_القاهرة_الجوي.png" "%DEST%\" >nul

echo [4/6] نسخ ازرار التشغيل...
copy /y "تثبيت.bat" "%DEST%\" >nul
copy /y "تشغيل.bat" "%DEST%\" >nul
copy /y "فحص-البوستة.bat" "%DEST%\" >nul
copy /y install.ps1 "%DEST%\" >nul
copy /y health.ps1 "%DEST%\" >nul
copy /y launch-hidden.vbs "%DEST%\" >nul
copy /y "تعليمات-الموظف.md" "%DEST%\" >nul
copy /y "دليل-التشغيل.md" "%DEST%\" >nul

echo [5/6] عمل مجلدات الشغل الفاضية...
mkdir "%DEST%\inbox" 2>nul
mkdir "%DEST%\scans" 2>nul
mkdir "%DEST%\backups" 2>nul

echo [6/6] تنظيف ملفات المطور من النسخة...
del /q "%DEST%\test_*.py" 2>nul
del /q "%DEST%\test_*.txt" 2>nul
del /q "%DEST%\*_out.txt" 2>nul
del /q "%DEST%\verify.log" 2>nul

echo.
echo ============================================================
echo     خلص! النسخة جاهزة في:
echo     %DEST%
echo ============================================================
echo.
echo   خطوات النقل لجهاز الموظف:
echo     1) انسخ مجلد PostalTool-Employee على فلاشة
echo     2) على جهاز الموظف: الصقه في اي مكان
echo     3) دبل كليك: تثبيت.bat (مرة واحدة - 3-10 دقايق + نت)
echo     4) دبل كليك: فحص-البوستة.bat (للتأكد)
echo     5) اطبع: تعليمات-الموظف.md وحطها جنب الجهاز
echo.
pause
