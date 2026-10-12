$ErrorActionPreference='Stop'
$taskRoot=$PSScriptRoot
$reviewPath=Join-Path $taskRoot 'المراجعة'
$env:POSTAL_XLSX=Join-Path $reviewPath 'سجل-المراجعة.xlsx'
$env:POSTAL_SCANS=Join-Path $reviewPath 'المحفوظات'
$env:POSTAL_INBOX=Join-Path $reviewPath 'استقبال-المسح'
$env:POSTAL_BACKUPS=Join-Path $reviewPath 'نسخ-احتياطية'
$env:POSTAL_PORT='8831'
$env:POSTAL_NO_BROWSER='1'
if (!(Test-Path -LiteralPath $env:POSTAL_XLSX)) { throw 'سجل المراجعة غير موجود؛ راجع تثبيت المشروع.' }
$connection=Get-NetTCPConnection -LocalPort 8831 -State Listen -ErrorAction SilentlyContinue
if (!$connection) {
    $pythonPath=Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\pythonw.exe'
    if (!(Test-Path -LiteralPath $pythonPath)) { $pythonPath=(Get-Command pythonw.exe).Source }
    Start-Process -FilePath $pythonPath -ArgumentList ('-B "'+(Join-Path $taskRoot 'postal_ocr.py')+'"') -WorkingDirectory $taskRoot -WindowStyle Hidden
}
Start-Process 'http://127.0.0.1:8831/'

