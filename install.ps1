#Requires -Version 5.1
<#
  مثبّت أداة تسجيل البوستة — يشتغل مرة واحدة على أي جهاز ويندوز.
  بيعمل: بايثون (لو ناقص) + المكتبات + نموذج القراءة العربي + اختصارات +
          تشغيل تلقائي مع بدء ويندوز (بعد الريستارت كمان) + فحص نهائي.
  طريقة الاستخدام: دوس دبل كليك على  تثبيت.bat   (مش محتاج أي خبرة)
#>
param([switch]$Uninstall)

$ErrorActionPreference = 'Continue'
$B = $PSScriptRoot
$log = Join-Path $B 'install.log'
$PY_DEFAULT = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
$TASK = 'PostalOCRStart'
$URL = 'http://127.0.0.1:8765'

function Step([string]$m) {
    $line = "[$(Get-Date -Format HH:mm:ss)] $m"
    Write-Host $line
    $line | Add-Content -LiteralPath $log -Encoding UTF8
}
function Say([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c }

function Find-Python {
    if (Test-Path $PY_DEFAULT) { return $PY_DEFAULT }
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Programs\Python') -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | ForEach-Object {
            $p = Join-Path $_.FullName 'python.exe'
            if (Test-Path $p) { $p }
        } | Select-Object -First 1
}

# ---------------------------------------------------------------- إزالة (Uninstall)
if ($Uninstall) {
    Step "إزالة الأداة من بدء التشغيل"
    $startup = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
    Remove-Item -LiteralPath (Join-Path $startup 'تسجيل البوستة.lnk') -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('Desktop')) 'تسجيل البوستة.lnk') -Force -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TASK -Confirm:$false -ErrorAction SilentlyContinue
    Get-CimInstance Win32_Process -Filter "Name like 'python%'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like '*postal_ocr.py*' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Step "تم الإزالة (البيانات وملف الإكسل اتفضلوا زي ما هم)"
    Read-Host "اضغط Enter للخروج" | Out-Null
    exit 0
}

"START $(Get-Date -Format s)" | Set-Content -LiteralPath $log -Encoding UTF8
$ok = $true
Say ''
Say '============================================' 'Cyan'
Say '  تثبيت أداة تسجيل البوستة (محلية 100%)' 'Cyan'
Say '============================================' 'Cyan'
Say ''

Step "1ب) تهيئة مسار إنجليزي"
# نسخة من الشعار باسم انجليزي (ايقونات الاختصارات بتفضل شغالة)
$logoAr = Join-Path $B 'شركة_ميناء_القاهرة_الجوي.png'
$logoEn = Join-Path $B 'logo.png'
if ((Test-Path $logoAr) -and (-not (Test-Path $logoEn))) {
    Copy-Item -LiteralPath $logoAr -Destination $logoEn -Force
    Step "logo.png created"
}

# ---------------------------------------------------------------- 1) بايثون
Step "1/6 البحث عن بايثون"
$py = Find-Python
if (-not $py) {
    Say "  بايثون مش موجود — بننصّبه دلوقتي (دقيقة أو دقيقتين)..." 'Yellow'
    Step "installing python via winget"
    winget install -e --id Python.Python.3.12 --scope user --silent `
        --accept-package-agreements --accept-source-agreements --disable-interactivity 2>&1 |
        Add-Content -LiteralPath $log
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
    $py = Find-Python
}
if (-not $py) { $ok = $false; Step "FATAL: python missing" }
else { Step "python = $py"; Say "  [OK] بايثون: $py" 'Green' }

# ---------------------------------------------------------------- 1ب) مسار إنجليزي
# واختصارات ويندوز + التشغيل التلقائي بيحتاجوا مسار إنجليزي، فبننقل نسخة
# من الأداة لـ %LOCALAPPDATA%\PostalTool لو الفولدر الحالي فيه حروف عربية.
$APP = $B
if ($B.FullName -cmatch '[^\x00-\x7F]') {
    $dst = Join-Path $env:LOCALAPPDATA 'PostalTool'
    Step "app path has non-ASCII -> copying to $dst"
    try {
        New-Item -ItemType Directory -Path $dst -Force | Out-Null
        # Refresh program files on repair; do not overwrite the employee register.
        $programFiles = @('postal_ocr.py', 'scan_wia.ps1', 'requirements.txt',
                         'models_dl.py', 'install.ps1', 'health.ps1', 'launch-hidden.vbs',
                         'تشغيل.bat', 'تثبيت.bat', 'فحص-البوستة.bat',
                         'تعليمات-الموظف.md', 'دليل-التشغيل.md', 'logo.png',
                         'شركة_ميناء_القاهرة_الجوي.png')
        foreach ($programFile in $programFiles) {
            $sourceFile = Join-Path $B $programFile
            if (Test-Path -LiteralPath $sourceFile) {
                Copy-Item -LiteralPath $sourceFile -Destination $dst -Force -ErrorAction Stop
            }
        }
        $uiSource = Join-Path $B 'ui'
        if (Test-Path -LiteralPath $uiSource) {
            Copy-Item -LiteralPath $uiSource -Destination $dst -Recurse -Force -ErrorAction Stop
        } else {
            throw 'Missing ui folder: the employee package is incomplete.'
        }
        $modelsSource = Join-Path $B 'models'
        if (Test-Path -LiteralPath $modelsSource) {
            Copy-Item -LiteralPath $modelsSource -Destination $dst -Recurse -Force -ErrorAction Stop
        }
        foreach ($dataName in @('البوستة.xlsx', 'inbox', 'scans', 'backups')) {
            $dataSource = Join-Path $B $dataName
            $dataDestination = Join-Path $dst $dataName
            if ((Test-Path -LiteralPath $dataSource) -and (-not (Test-Path -LiteralPath $dataDestination))) {
                Copy-Item -LiteralPath $dataSource -Destination $dst -Recurse -ErrorAction Stop
            }
        }
        $APP = Get-Item $dst
        Say "  [OK] نسخة الأداة في مسار إنجليزي: $dst" 'Green'
        Say "     (دول عشان الاختصارات والتشغيل التلقائي تشتغل صح)" 'DarkGray'
    } catch {
        Step "copy failed: $($_.Exception.Message)"
        Say "  [!!] مقدرش ننقل الأداة لمسار إنجليزي — هيشتغل من الفولدر الحالي" 'Yellow'
    }
}
if ($py) {
    # مسار pythonw بالظبط (بحث عنه الـ VBS وقت التشغيل)
    $pythonw = Join-Path (Split-Path $py -Parent) 'pythonw.exe'
    if (Test-Path $pythonw) {
        Set-Content -LiteralPath (Join-Path $APP 'pythonw.txt') -Value $pythonw -Encoding ASCII
        Step "pythonw.txt = $pythonw"
    }
}
# ---------------------------------------------------------------- 2) المكتبات
if ($py) {
    Step "2/6 تنزيل المكتبات (مرة واحدة)"
    $pipLog = & $py -m pip install --disable-pip-version-check --no-input --upgrade -r (Join-Path $APP 'requirements.txt') 2>&1
    ($pipLog | Select-Object -Last 6) | Add-Content -LiteralPath $log
    $chk = & $py -c "import openpyxl,PIL,fitz,rapidocr,onnxruntime;print('imports ok')" 2>&1
    "$chk" | Add-Content -LiteralPath $log
    if ("$chk" -match 'imports ok') {
        Say "  [OK] المكتبات جاهزة" 'Green'
    } else {
        # غالبًا ملف مثبّت مفتوح أو صلاحيات — نحاول تثبيت للمستخدم نفسه
        Step "retrying pip with --user"
        (& $py -m pip install --disable-pip-version-check --no-input --user -r (Join-Path $APP 'requirements.txt') 2>&1 |
            Select-Object -Last 6) | Add-Content -LiteralPath $log
        $chk = & $py -c "import openpyxl,PIL,fitz,rapidocr,onnxruntime;print('imports ok')" 2>&1
        if ("$chk" -match 'imports ok') { Say "  [OK] المكتبات جاهزة" 'Green' }
        else { $ok = $false; Step "FATAL: imports failed" }
    }

    # ------------------------------------------------------------ 3) النموذج العربي
    Step "3/6 نموذج القراءة العربي"
    if (Test-Path (Join-Path $APP 'models\arabic_PP-OCRv5_rec_mobile.onnx')) {
        Say "  [OK] النموذج موجود" 'Green'
    } else {
        Say "  بنحمّل النموذج (مرة واحدة)" 'Yellow'
        (& $py (Join-Path $APP 'models_dl.py') 2>&1 | Select-Object -Last 6) | Add-Content -LiteralPath $log
        if (Test-Path (Join-Path $APP 'models\arabic_PP-OCRv5_rec_mobile.onnx')) {
            Say "  [OK] النموذج اتحمّل" 'Green'
        } else { $ok = $false; Step "FATAL: arabic model missing" }
    }

    if ($ok) {
        Step "تجهيز جميع نماذج القراءة قبل الاستخدام اليومي"
        Push-Location $APP
        try {
            $modelCheck = & $py -c "import postal_ocr; postal_ocr.ocr_engine(); postal_ocr.digits_engine(); print('ocr models ready')" 2>&1
            "$modelCheck" | Add-Content -LiteralPath $log
            if ($LASTEXITCODE -ne 0 -or "$modelCheck" -notmatch 'ocr models ready') {
                $ok = $false
                Step "FATAL: OCR initialization failed"
            }
        } finally { Pop-Location }
    }

    Step "4/6 فحص الكود"
    & $py -m py_compile (Join-Path $APP 'postal_ocr.py') 2>&1 | Add-Content -LiteralPath $log
    "compile-exit=$LASTEXITCODE" | Add-Content -LiteralPath $log
    if ($LASTEXITCODE -ne 0) { $ok = $false; Step "FATAL: syntax error" }
    else { Say "  [OK] الكود سليم" 'Green' }
}

# ---------------------------------------------------------------- 5) الاختصارات + التشغيل التلقائي
Step "5/6 الاختصارات والتشغيل التلقائي"
try {
    $ws = New-Object -ComObject WScript.Shell
    $desktop = [Environment]::GetFolderPath('Desktop')
    # ملاحظة: WSH بيكتب الاختصارات بترميز قديم، فبنعمله باسم إنجليزي
    # وبعدين بنسمّيه عربي (مستكشف ويندوز بيقبل الاسم العربي عادي).
    $asciiLnk = Join-Path $desktop 'PostalRegister.lnk'
    $arLnk = Join-Path $desktop 'تسجيل البوستة.lnk'
    $vbs = Join-Path $APP 'launch-hidden.vbs'
    $lnkArgs = '"' + $vbs + '"'

    $lnk = $ws.CreateShortcut($asciiLnk)
    $lnk.TargetPath = 'wscript.exe'
    $lnk.Arguments = $lnkArgs
    $lnk.WorkingDirectory = $APP
    $lnk.Description = 'تسجيل البوستة'
    $logo = Join-Path $APP 'logo.png'
    if (Test-Path $logo) { $lnk.IconLocation = $logo + ',0' }
    $lnk.Save()
    Move-Item -LiteralPath $asciiLnk -Destination $arLnk -Force
    Step "desktop shortcut ok ($arLnk)"

    $startup = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
    New-Item -ItemType Directory -Path $startup -Force | Out-Null
    $sl = $ws.CreateShortcut((Join-Path $startup 'PostalRegister.lnk'))
    $sl.TargetPath = 'wscript.exe'
    $sl.Arguments = $lnkArgs
    $sl.WorkingDirectory = $APP
    $sl.Save()
    Step "startup shortcut ok"

    # مهمة مجدولة: أدق من مجلد Startup وبتشتغل مع كل تسجيل دخول (بعد الريستارت كمان)
    try {
        $action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument $lnkArgs
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) `
            -MultipleInstances IgnoreNew
        Register-ScheduledTask -TaskName $TASK -Action $action -Trigger $trigger `
            -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
        Step "scheduled task ok"
        Say "  [OK] تشغيل تلقائي مع بدء ويندوز" 'Green'
    } catch {
        Step "scheduled task skipped: $($_.Exception.Message)"
        Say "  [OK] تشغيل تلقائي (عبر مجلد Startup)" 'Green'
    }
} catch {
    Step "shortcut error: $($_.Exception.Message)"
    $ok = $false
}

# ---------------------------------------------------------------- 6) تشغيل وفحص
if ($ok) {
    Step "6/6 تشغيل الأداة وفحصها"
    Start-Process -FilePath 'wscript.exe' -ArgumentList ('"' + (Join-Path $APP 'launch-hidden.vbs') + '"') -WindowStyle Hidden
    $live = $false
    for ($i = 0; $i -lt 45; $i++) {
        Start-Sleep -Seconds 2
        try {
            $r = Invoke-WebRequest -Uri ($URL + '/meta') -TimeoutSec 4 -UseBasicParsing
            if ($r.StatusCode -eq 200) { $live = $true; break }
        } catch { }
    }
    if ($live) { Say "  ✔ الأداة شغالة على $URL" 'Green' }
    else { $ok = $false; Step "FATAL: server not responding" }
}

# ---------------------------------------------------------------- النتيجة
Say ''
if ($ok) {
    Say '============================================' 'Green'
    Say '  التثبيت تم بنجاح' 'Green'
    Say '============================================' 'Green'
    Say ''
    Say '  الاستخدام اليومي: دوس دبل كليك على أيقونة «تسجيل البوستة»'
    Say '  على سطح المكتب — هيفتح الصفحة في المتصفح.'
    Say '  والأداة بتفتح لوحدها مع كل مرة تدخل فيها ويندوز.'
    Say ''
    Say "  سجل ($(Get-Date -Format HH:mm)): $log"
    Start-Process $URL
} else {
    Say '============================================' 'Red'
    Say '  في مشكلة — شوف التفاصيل في install.log' 'Red'
    Say '============================================' 'Red'
    Say ''
    Say '  الحلول الشائعة:'
    Say '   1) اقفل أي أدوات بايثون مفتوحة وجرّب تاني.'
    Say '   2) اتأكد إن الإنترنت شغال (لازم مرة واحدة عشان التحميل).'
    Say '   3) جرّب تشغّل الأداة يدوي: دوس على تشغيل.bat'
    Say ''
    Say "  سجل ($(Get-Date -Format HH:mm)): $log"
}
Read-Host "اضغط Enter للخروج" | Out-Null
