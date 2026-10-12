#Requires -Version 5.1
<#
  فحص سريع: هل الأداة شغالة والكل نضيف؟
  دبل كليك على  فحص-البوستة.bat
#>
$ErrorActionPreference = 'Continue'
$B = $PSScriptRoot
# الأداة ممكن تكون متنقولة لمسار انجليزي وقت التثبيت (للاختصارات)
$installed = Join-Path $env:LOCALAPPDATA 'PostalTool'
if (Test-Path (Join-Path $installed 'postal_ocr.py')) { $B = $installed }
$URL = 'http://127.0.0.1:8765'
$py = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
if (-not (Test-Path $py)) {
    $c = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($c) { $py = $c.Source } else { $py = '' }
}

$script:bad = 0
function Line([bool]$ok, [string]$txt, [string]$hint = '') {
    if ($ok) { Write-Host "  [OK]  $txt" -ForegroundColor Green }
    else {
        Write-Host "  [!!]  $txt" -ForegroundColor Red
        if ($hint) { Write-Host "        -> $hint" -ForegroundColor DarkYellow }
    }
    if (-not $ok) { $script:bad++ }
}

Write-Host ''
Write-Host '============================================' 'Cyan'
Write-Host '  فحص سريع - أداة تسجيل البوستة' 'Cyan'
Write-Host '============================================' 'Cyan'

Line (Test-Path (Join-Path $B 'postal_ocr.py')) 'ملف الأداة موجود'
foreach ($uiFile in @('ui\index.html','ui\postal.css','ui\postal.js')) {
    Line (Test-Path (Join-Path $B $uiFile)) "ملف الواجهة موجود: $uiFile" 'أعد تجهيز أو تثبيت نسخة الموظف كاملة'
}
Line ([bool]$py -and (Test-Path $py)) 'بايثون متثبت' 'شغل تثبيت.bat مرة واحدة'
if ($py -and (Test-Path $py)) {
    $chk = & $py -c "import openpyxl,PIL,fitz,rapidocr,onnxruntime;print('ok')" 2>&1
    Line ("$chk" -match 'ok') 'المكتبات المطلوبة موجودة'
    $c2 = & $py -c "import sys;sys.path.insert(0,r'$B');import postal_ocr;print('ok')" 2>&1
    Line ("$c2" -match 'ok') 'كود الأداة سليم'
}
Line (Test-Path (Join-Path $B 'models\arabic_PP-OCRv5_rec_mobile.onnx')) 'نموذج القراءة العربي موجود' 'شغل تثبيت.bat'
Line (Test-Path (Join-Path $B 'البوستة.xlsx')) 'ملف السجل موجود (البوستة.xlsx)'
try {
    $r = Invoke-WebRequest -Uri ($URL + '/meta') -TimeoutSec 5 -UseBasicParsing
    Line ($r.StatusCode -eq 200) 'الأداة شغالة دلوقتي على 8765'
} catch {
    Line $false 'الأداة مش شغالة دلوقتي' 'دوس دبل كليك على ايقونة تسجيل البوستة في سطح المكتب'
}
Line ([bool](Get-ScheduledTask -TaskName 'PostalOCRStart' -ErrorAction SilentlyContinue)) 'التشغيل التلقائي بعد الريستارت متظبط'
$lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'تسجيل البوستة.lnk'
Line (Test-Path $lnk) 'اختصار سطح المكتب موجود'

Write-Host ''
if ($script:bad -eq 0) {
    Write-Host '  النتيجة: كل حاجة تمام.' -ForegroundColor Green
    Write-Host ''
    Start-Process $URL
} else {
    Write-Host "  النتيجة: في $script:bad بند محتاج تصحيح" -ForegroundColor Yellow
    Write-Host '  اسهل حل: دوس دبل كليك على  تثبيت.bat (تثبيت او اصلاح كامل)' -ForegroundColor Yellow
    Write-Host '  ثم شغل  فحص-البوستة.bat  تاني للتاكيد.' -ForegroundColor Gray
}
Write-Host ''
