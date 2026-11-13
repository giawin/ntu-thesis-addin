# build-all.ps1 - dung lai toan bo bo cai NTU chi bang MOT lenh.
#
# Chuoi cong viec (khong co buoc sua tay nao):
#   1. Trich VBA + ribbon tu hai tep goc trong src/core/vendor/upstream/
#   2. Ap gia tri ca nhan hoa NTU (personalize.mjs) - tat dinh, idempotent
#   3. Dong goi .dotm bang Word COM (build-word.vbs + pack-word.ps1)
#   4. Dung add-in NTU (build-ntu.vbs + pack-ntu.ps1)
#   5. Kiem chung ket qua (verify-word.vbs + kiem tra goi OOXML)
#   6. Gom bo cai vao bo-cai-dat/ (installer, huong dan, version.json)
#
# Dung:
#   powershell -ExecutionPolicy Bypass -File src/build/build-all.ps1
#   powershell -ExecutionPolicy Bypass -File src/build/build-all.ps1 -SkipNtu
#
# Yeu cau: Windows + Microsoft Word (de dung .dotm) + Node.js.
# Khong co Word thi buoc 3/4 that bai va script dung lai, KHONG gia tao ket qua.
#
# Thoat: 0 = tat ca dat, 1 = dung o buoc nao do (in ro buoc va loi).

[CmdletBinding()]
param(
    [switch]$SkipNtu,      # chi dung add-in Word (bo qua module NTU)
    [switch]$SkipVerify    # bo qua buoc kiem chung (khong khuyen dung)
)

$ErrorActionPreference = 'Stop'
$script:stepNo = 0
$script:sw = [Diagnostics.Stopwatch]::StartNew()

function Buoc([string]$ten) {
    $script:stepNo++
    Write-Host ""
    Write-Host ("=" * 68)
    Write-Host ("BUOC {0}: {1}" -f $script:stepNo, $ten)
    Write-Host ("=" * 68)
}

function Dung([string]$lyDo) {
    Write-Host ""
    Write-Host ("DUNG: {0}" -f $lyDo) -ForegroundColor Red
    Write-Host ("Da chay {0} buoc, {1:N1} giay." -f $script:stepNo, $script:sw.Elapsed.TotalSeconds)
    exit 1
}

# --- Xac dinh duong dan ---
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$build = $PSScriptRoot
$vendor = Join-Path $repo 'src\core\vendor\upstream'
$wordDir = Join-Path $repo 'src\core\word'
$ntuDir = Join-Path $repo 'src\modules\ntu-thesis'
$outDir = Join-Path $repo 'bo-cai-dat'

Write-Host "Kho     : $repo"
Write-Host "Ban goc : $vendor"

# --- Tim Node ---
$node = $null
foreach ($p in @(
    'C:\Program Files\nodejs\node.exe',
    "$env:LOCALAPPDATA\Programs\nodejs\node.exe",
    'C:\Program Files (x86)\nodejs\node.exe'
)) {
    if (Test-Path -LiteralPath $p) { $node = $p; break }
}
if (-not $node) {
    $cmd = Get-Command node -ErrorAction SilentlyContinue
    if ($cmd) { $node = $cmd.Source }
}
if (-not $node) { Dung 'Khong tim thay node.exe. Cai Node.js roi chay lai.' }
Write-Host "Node    : $node"

# --- Kiem tra tep goc ---
$upDotm = Join-Path $vendor 'ChuanHoa_VanBanDang.dotm'
$upXlam = Join-Path $vendor 'ChuanHoa_Excel.xlam'
foreach ($f in @($upDotm, $upXlam)) {
    if (-not (Test-Path -LiteralPath $f)) {
        Dung "Thieu tep goc: $f"
    }
}

# =========================================================================
Buoc 'Trich VBA tu tep goc'
# =========================================================================
$extract = Join-Path $build 'extract-vba.ps1'
$extractOut = Join-Path $env:TEMP 'ntu-build-extract'
if (Test-Path -LiteralPath $extractOut) {
    Remove-Item -LiteralPath $extractOut -Recurse -Force
}
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $extract `
    -SourcePath $upDotm -OutDir $extractOut -AppKind word
if ($LASTEXITCODE -ne 0) { Dung "extract-vba.ps1 that bai (exit $LASTEXITCODE)" }

$basGoc = Join-Path $extractOut 'ChuanHoa_VanBanDang__Module1.bas'
if (-not (Test-Path -LiteralPath $basGoc)) {
    Dung "Khong thay $basGoc sau khi trich"
}
# Cap nhat ban goc dung de doi chieu (giu nguyen ban chua ca nhan hoa)
Copy-Item -LiteralPath $basGoc -Destination (Join-Path $wordDir 'Module1.upstream.bas') -Force
Copy-Item -LiteralPath (Join-Path $extractOut 'ChuanHoa_VanBanDang__customUI.xml') `
    -Destination (Join-Path $wordDir 'customUI.upstream.xml') -Force
Write-Host "OK -> Module1.upstream.bas, customUI.upstream.xml"

# =========================================================================
Buoc 'Ap gia tri ca nhan hoa NTU'
# =========================================================================
Push-Location $build
try {
    & $node 'personalize.mjs' (Join-Path $wordDir 'Module1.upstream.bas') `
        (Join-Path $wordDir 'Module1.vba')
    if ($LASTEXITCODE -ne 0) { Dung "personalize.mjs that bai (exit $LASTEXITCODE)" }

    # Tu kiem tra: phai tao lai dung ban da kiem chung
    & $node 'personalize.test.mjs'
    if ($LASTEXITCODE -ne 0) { Dung 'personalize.test.mjs that bai' }
} finally {
    Pop-Location
}

# Ribbon Word: chep tu ban goc (chi bo BOM UTF-8), khong sua noi dung.
$uiSrc = Join-Path $wordDir 'customUI.upstream.xml'
$uiDst = Join-Path $wordDir 'customUI.xml'
$uiText = [IO.File]::ReadAllText($uiSrc, [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($uiDst, $uiText, (New-Object Text.UTF8Encoding($false)))
Write-Host "OK -> Module1.vba, customUI.xml"

# =========================================================================
Buoc 'Dong goi add-in Word'
# =========================================================================
$wordOut = Join-Path $wordDir 'dist'
foreach ($f in @(
    (Join-Path $wordDir 'Module1.vba'),
    (Join-Path $wordDir 'customUI.xml')
)) {
    if (-not (Test-Path -LiteralPath $f)) { Dung "Thieu $f" }
}
& cscript.exe //nologo (Join-Path $build 'word\build-word.vbs') $repo
if ($LASTEXITCODE -ne 0) { Dung "build-word.vbs that bai (exit $LASTEXITCODE)" }

& powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File (Join-Path $build 'word\pack-word.ps1')
if ($LASTEXITCODE -ne 0) { Dung "pack-word.ps1 that bai (exit $LASTEXITCODE)" }

# =========================================================================
if (-not $SkipNtu) {
    Buoc 'Dung add-in NTU'
    # =====================================================================
    foreach ($f in @(
        (Join-Path $ntuDir 'NTU.vba'),
        (Join-Path $ntuDir 'customUI.xml')
    )) {
        if (-not (Test-Path -LiteralPath $f)) { Dung "Thieu $f" }
    }

    $ntuWork = Join-Path $ntuDir 'dist_work'
    if (-not (Test-Path -LiteralPath $ntuWork)) {
        New-Item -ItemType Directory -Path $ntuWork -Force | Out-Null
    }

    Push-Location $build
    try {
        & $node 'vba2bas.mjs' (Join-Path $ntuDir 'NTU.vba') (Join-Path $ntuWork 'NTU.bas')
        if ($LASTEXITCODE -ne 0) { Dung "vba2bas.mjs that bai (exit $LASTEXITCODE)" }
    } finally {
        Pop-Location
    }

    & cscript.exe //nologo (Join-Path $build 'ntu\build-ntu.vbs') $repo
    if ($LASTEXITCODE -ne 0) { Dung "build-ntu.vbs that bai (exit $LASTEXITCODE)" }

    & powershell.exe -NoProfile -ExecutionPolicy Bypass `
        -File (Join-Path $build 'ntu\pack-ntu.ps1')
    if ($LASTEXITCODE -ne 0) { Dung "pack-ntu.ps1 that bai (exit $LASTEXITCODE)" }
}

# =========================================================================
Buoc 'Kiem chung'
# =========================================================================
$wordDotm = Join-Path $wordOut 'GiaTwinH_NTU_Thesis_Word.dotm'
if (-not (Test-Path -LiteralPath $wordDotm)) { Dung "Khong thay $wordDotm" }

if (-not $SkipVerify) {
    & cscript.exe //nologo (Join-Path $build 'word\verify-word.vbs') $wordDotm
    if ($LASTEXITCODE -ne 0) { Dung "verify-word.vbs that bai (exit $LASTEXITCODE)" }

    if (-not $SkipNtu) {
        $ntuVerify = Join-Path $ntuDir 'dist\NTUThesis.dotm'
        & cscript.exe //nologo (Join-Path $build 'ntu\verify-ntu.vbs') $ntuVerify
        if ($LASTEXITCODE -ne 0) { Dung "verify-ntu.vbs that bai (exit $LASTEXITCODE)" }
    }
}

# Kiem tra goi OOXML cua ca hai add-in
Add-Type -AssemblyName System.IO.Compression.FileSystem
function KiemGoi([string]$tep, [string[]]$batBuoc) {
    if (-not (Test-Path -LiteralPath $tep)) { return "THIEU TEP" }
    $zip = [IO.Compression.ZipFile]::OpenRead($tep)
    try {
        $co = $zip.Entries | ForEach-Object { $_.FullName }
        $thieu = $batBuoc | Where-Object { $co -notcontains $_ }
        if ($thieu) { return "THIEU PART: $($thieu -join ', ')" }
        return "OK ($($co.Count) part)"
    } finally { $zip.Dispose() }
}

$kqWord = KiemGoi $wordDotm @(
    'customUI/customUI.xml', 'word/vbaProject.bin', 'word/glossary/document.xml'
)
$ntuDotm = Join-Path $ntuDir 'dist\NTUThesis.dotm'
$kqNtu = if ($SkipNtu) { 'BO QUA' } else {
    KiemGoi $ntuDotm @('customUI/customUI.xml', 'word/vbaProject.bin')
}
Write-Host ("  Word .dotm : {0}" -f $kqWord)
Write-Host ("  NTU  .dotm : {0}" -f $kqNtu)
if ($kqWord -notlike 'OK*') {
    Dung "Goi Word .dotm khong dat: $kqWord"
}
if (-not $SkipNtu -and $kqNtu -notlike 'OK*') {
    Dung "Goi NTU .dotm khong dat: $kqNtu"
}

# =========================================================================
Buoc 'Gom bo cai dat'
# =========================================================================
if (Test-Path -LiteralPath $outDir) { Remove-Item -LiteralPath $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

Copy-Item -LiteralPath $wordDotm -Destination $outDir -Force
if (-not $SkipNtu -and (Test-Path -LiteralPath $ntuDotm)) {
    Copy-Item -LiteralPath $ntuDotm -Destination $outDir -Force
}
# Excel: NGUYEN BAN GOC cua tac gia, khong build lai, khong sua mot byte.
Copy-Item -LiteralPath $upXlam -Destination $outDir -Force

$installer = Join-Path $build 'installer'
Copy-Item -LiteralPath (Join-Path $installer 'CaiDat.bat') -Destination $outDir -Force
Copy-Item -LiteralPath (Join-Path $installer 'GoCaiDat.bat') -Destination $outDir -Force
$huongDan = Join-Path $repo 'HuongDan_CaiDat.txt'
if (Test-Path -LiteralPath $huongDan) {
    Copy-Item -LiteralPath $huongDan -Destination $outDir -Force
}

# version.json - manifest de add-in tu cap nhat
$base = 'https://github.com/giawin/ntu-thesis-addin/releases/latest/download'
function Sha([string]$p) {
    (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLower()
}
$version = @{
    version      = '1.0.0'
    notes_ascii  = 'Chuan hoa van ban Word va can chinh trang in Excel theo chuan NTU.'
    word_url     = "$base/GiaTwinH_NTU_Thesis_Word.dotm"
    word_sha256  = (Sha $wordDotm)
    word_size    = (Get-Item -LiteralPath $wordDotm).Length
    excel_url    = "$base/ChuanHoa_Excel.xlam"
    excel_sha256 = (Sha (Join-Path $outDir 'ChuanHoa_Excel.xlam'))
    excel_size   = (Get-Item -LiteralPath (Join-Path $outDir 'ChuanHoa_Excel.xlam')).Length
    published    = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
}
if (-not $SkipNtu -and (Test-Path -LiteralPath $ntuDotm)) {
    $version['ntu_url'] = "$base/NTUThesis.dotm"
    $version['ntu_sha256'] = (Sha $ntuDotm)
    $version['ntu_size'] = (Get-Item -LiteralPath $ntuDotm).Length
}
$json = $version | ConvertTo-Json -Depth 4
[IO.File]::WriteAllText((Join-Path $outDir 'version.json'), $json,
    (New-Object Text.UTF8Encoding($false)))

# --- Zip bo cai ---
$zipName = 'Bo_Cai_Dat_NTU_Thesis_Addin_v1.0.0.zip'
$zipPath = Join-Path $repo $zipName
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $outDir '*') -DestinationPath $zipPath -CompressionLevel Optimal

Write-Host ""
Write-Host ("=" * 68)
Write-Host "XONG - ket qua:"
Write-Host ("=" * 68)
Get-ChildItem -LiteralPath $outDir -File | ForEach-Object {
    Write-Host ("  {0,10:N0} B  {1}" -f $_.Length, $_.Name)
}
Write-Host ("  {0,10:N0} B  {1} (bo cai nop kem)" -f (Get-Item -LiteralPath $zipPath).Length, $zipName)
Write-Host ""
Write-Host ("Tong thoi gian: {0:N1} giay" -f $script:sw.Elapsed.TotalSeconds)
Write-Host "Gui nguoi dung: $zipPath"
exit 0