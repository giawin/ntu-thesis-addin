# extract-vba.ps1 - trich VBA + customUI tu mot tep .dotm/.xlam goc.
#
# Dung Word/Excel COM de mo tep va xuat (export) cac module VBA ra tep .bas,
# dong thoi rut customUI/customUI.xml ra tep .xml. Day chinh la buoc "trich lai
# VBA" ma .github/workflows/sync-upstream.yml nhac toi.
#
# Dung:
#   .\extract-vba.ps1 -SourcePath <tep.dotm|tep.xlam> -OutDir <thu muc dich> [-AppKind word|excel]
#
# Vi du:
#   .\extract-vba.ps1 -SourcePath ..\..\core\vendor\upstream\ChuanHoa_VanBanDang.dotm `
#                     -OutDir ..\..\core\vendor\upstream\extracted
#
# Tep xuat ra:
#   <OutDir>\<TenTep>__<TenModule>.bas   (ma nguon VBA, ANSI/ASCII)
#   <OutDir>\<TenTep>__customUI.xml      (ribbon)
#
# Yeu cau: Windows + Microsoft Office (Word hoac Excel) da cai.
# Thoat: 0 = thanh cong, 1 = thieu tham so, 2 = khong tim thay tep,
#        3 = khong khoi dong duoc COM, 4 = khong xuat duoc module nao.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [ValidateSet('word','excel')][string]$AppKind = 'word'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SourcePath)) {
    Write-Error "Khong tim thay tep: $SourcePath"
    exit 2
}
$SourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path
$base = [IO.Path]::GetFileNameWithoutExtension($SourcePath)

# Office mac dinh chan truy cap vao VBA project. Bat lai cho HKCU.
$vbaKey = 'HKCU:\Software\Microsoft\Office\16.0\Common\Security'
try {
    if (-not (Test-Path $vbaKey)) { New-Item -Path $vbaKey -Force | Out-Null }
    New-ItemProperty -Path $vbaKey -Name 'AccessVBOM' -Value 1 `
        -PropertyType DWord -Force | Out-Null
} catch {
    Write-Warning "Khong dat duoc AccessVBOM (co the van chay duoc): $($_.Exception.Message)"
}

$app = $null
$doc = $null
$extracted = New-Object System.Collections.Generic.List[string]

try {
    if ($AppKind -eq 'word') {
        $app = New-Object -ComObject Word.Application
        $app.Visible = $false
        $app.DisplayAlerts = 0
        # 3 = msoAutomationSecurityForceDisable: khong chay macro khi mo tep.
        # Bat buoc, neu khong add-in goc se chay co che ban quyen cua tac gia.
        $app.AutomationSecurity = 3
        $doc = $app.Documents.Open($SourcePath, $false, $true)
        $components = $doc.VBProject.VBComponents
    } else {
        $app = New-Object -ComObject Excel.Application
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.AutomationSecurity = 3
        $doc = $app.Workbooks.Open($SourcePath, 0, $true)
        $components = $doc.VBProject.VBComponents
    }
} catch {
    Write-Error "Khong khoi dong duoc $AppKind COM hoac khong mo duoc tep: $($_.Exception.Message)"
    exit 3
}

# --- Xuat tung module VBA ---
$count = 0
foreach ($c in $components) {
    $ten = $c.Name
    $dich = Join-Path $OutDir "$base`__$ten.bas"
    try {
        # 0 = vbext_ct_StdModule, 1 = ClassModule, 2 = Form, 3 = ThisDocument...
        # Export hoat dong voi moi loai tru ThisDocument cua Word (khong xuat duoc).
        $c.Export($dich)
        if ((Test-Path -LiteralPath $dich) -and (Get-Item -LiteralPath $dich).Length -gt 0) {
            $kb = [math]::Round((Get-Item -LiteralPath $dich).Length / 1KB, 1)
            Write-Host ("  {0,-24} {1,8} KB  -> {2}" -f $ten, $kb, [IO.Path]::GetFileName($dich))
            $extracted.Add($dich)
            $count++
        }
    } catch {
        Write-Warning "  Bo qua module '$ten': $($_.Exception.Message)"
    }
}

# --- Dong va giai phong ---
try { if ($doc) { $doc.Close($false) } } catch {}
try { if ($app) { $app.Quit() } } catch {}
foreach ($o in @($doc, $app)) {
    if ($o) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {}
    }
}
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# --- Rut customUI tu goi OOXML (khong can COM) ---
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($SourcePath)
try {
    foreach ($tenPart in @('customUI/customUI.xml', 'customUI/customUI14.xml')) {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $tenPart }
        if ($entry) {
            # Bo hau to "14" de ten tep on dinh giua hai ban
            $nhan = if ($tenPart -like '*14*') { 'customUI14' } else { 'customUI' }
            $dich = Join-Path $OutDir "$base`__$nhan.xml"
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dich, $true)
            Write-Host ("  {0,-24} {1,8} KB  -> {2}" -f $tenPart, 
                [math]::Round((Get-Item -LiteralPath $dich).Length / 1KB, 1),
                [IO.Path]::GetFileName($dich))
            $extracted.Add($dich)
            $count++
        }
    }
} finally {
    $zip.Dispose()
}

Write-Host ""
if ($count -eq 0) {
    Write-Error "Khong xuat duoc tep nao tu $SourcePath"
    exit 4
}
Write-Host "OK -> $count tep trong $OutDir"
exit 0