# WIA عبر COM: تعداد الأجهزة + مسح صامت. يطبع JSON للمسار الصادر.
param(
    [string]$Action = "list",
    [int]$Device = -1,
    [string]$Out = "",
    [int]$Dpi = 300,
    [string]$Color = "color",
    [string]$Source = "auto"
)

$ErrorActionPreference = "Stop"

function Get-WiaManager {
    $m = New-Object -ComObject WIA.DeviceManager
    if ($null -eq $m) { throw "WIA is not available on this machine." }
    return $m
}

function Device-Props($dev) {
    $p = @{}
    foreach ($pr in $dev.Properties) { $p[$pr.Name] = "$($pr.Value)" }
    return $p
}

function Find-Device($m, $index) {
    $devs = @()
    foreach ($info in $m.DeviceInfos) {
        if ($info.Type -eq 1) { $devs += $info }
    }
    if ($devs.Count -eq 0) { throw "NO_SCANNER" }
    if ($index -ge 0 -and $index -lt $devs.Count) { return $devs[$index] }
    return $devs[0]
}

function Device-Record($info, $index) {
    $name = ""
    $port = ""
    try {
        foreach ($pr in $info.Properties) {
            if ($pr.Name -eq "Name") { $name = "$($pr.Value)" }
            if ($pr.Name -eq "Port") { $port = "$($pr.Value)" }
        }
    } catch { }
    return @{ index = $index; name = $name; port = $port }
}

if ($Action -eq "list") {
    $m = Get-WiaManager
    $list = @()
    $i = 0
    foreach ($info in $m.DeviceInfos) {
        if ($info.Type -eq 1) {
            $list += (Device-Record $info $i)
            $i += 1
        }
    }
    @{ ok = $true; devices = $list } | ConvertTo-Json -Depth 4 -Compress
    exit 0
}

# ---- scan ----
$m = Get-WiaManager
$info = Find-Device $m $Device
$dev = $info.Connect()

if ($Action -eq "status") {
    $feederStatus = 0
    foreach ($pr in $dev.Properties) {
        if ($pr.Name -eq "Document Handling Status") { $feederStatus = [int]$pr.Value }
    }
    @{ ok = $true; ready = (($feederStatus -band 1) -ne 0) } | ConvertTo-Json -Compress
    exit 0
}

# اختيار مصدر الورق: مسطح أو مغذي (ADF)
try {
    $handling = $null
    foreach ($pr in $dev.Properties) {
        if ($pr.Name -eq "Document Handling Select") { $handling = $pr; break }
    }
    foreach ($item in $dev.Items) {
        if ($null -ne $handling) { break }
        foreach ($pr in $item.Properties) {
            if ($pr.Name -eq "Document Handling Select") { $handling = $pr }
        }
        if ($null -ne $handling) { break }
    }
    if ($null -ne $handling) {
        $FEEDER = 1; $FLATBED = 2
        if ($Source -eq "flatbed") { $handling.Value = $FLATBED }
        elseif ($Source -eq "feeder") { $handling.Value = $FEEDER }
        else {
            # auto: المغذي لو متاح وفيه ورق، غير كده المسطح
            try {
                $status = 0
                foreach ($pr in $dev.Properties) {
                    if ($pr.Name -eq "Document Handling Status") { $status = [int]$pr.Value }
                }
                if (($status -band $FEEDER) -ne 0) { $handling.Value = $FEEDER }
                else { $handling.Value = $FLATBED }
            } catch { $handling.Value = $FLATBED }
        }
    }
} catch { }

$item = $dev.Items[1]
foreach ($pr in $item.Properties) {
    if ($pr.Name -eq "Horizontal Resolution") { $pr.Value = $Dpi }
    if ($pr.Name -eq "Vertical Resolution") { $pr.Value = $Dpi }
    if ($pr.Name -eq "Current Intent") {
        if ($Color -eq "gray") { $pr.Value = 2 }
        elseif ($Color -eq "bw") { $pr.Value = 4 }
        else { $pr.Value = 1 }
    }
}

$img = $item.Transfer("{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")
if ([string]::IsNullOrWhiteSpace($Out)) {
    $Out = Join-Path ([IO.Path]::GetTempPath()) ("postal-scan-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".png")
}
if (Test-Path $Out) { Remove-Item -LiteralPath $Out -Force }
$img.SaveFile($Out)
@{ ok = $true; path = $Out; device = (Device-Record $info $Device).name } |
    ConvertTo-Json -Depth 3 -Compress
