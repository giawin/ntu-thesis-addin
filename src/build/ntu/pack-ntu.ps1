# Inject the NTU ribbon into the .dotm produced by build-ntu.vbs, then move it to dist/.
# Pure ASCII on purpose (PowerShell 5.1 reads non-ASCII .ps1 as ANSI).
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$here   = $PSScriptRoot
$repo   = (Resolve-Path (Join-Path $here '..\..\..')).Path
$srcDir = Join-Path $repo 'src\modules\ntu-thesis'
$work   = Join-Path $srcDir 'dist_work'
$raw    = Join-Path $work 'NTUThesis-raw.dotm'
$dist   = Join-Path $srcDir 'dist'
$out    = Join-Path $dist 'NTUThesis.dotm'
$ribbon = Join-Path $srcDir 'customUI.xml'

if (-not (Test-Path $raw))    { throw "missing $raw - run build-ntu.vbs first" }
if (-not (Test-Path $ribbon)) { throw "missing $ribbon" }
New-Item -ItemType Directory -Force -Path $dist | Out-Null

function Put-Entry($zip, $name, $text) {
    $old = $zip.GetEntry($name)
    if ($old) { $old.Delete() }
    $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $s = $e.Open()
    $b = [System.Text.Encoding]::UTF8.GetBytes($text)
    $s.Write($b, 0, $b.Length)
    $s.Close()
}
function Get-EntryText($zip, $name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return '' }
    $sr = New-Object System.IO.StreamReader($e.Open())
    $t = $sr.ReadToEnd(); $sr.Close(); return $t
}

$ui = [System.IO.File]::ReadAllText($ribbon, [System.Text.Encoding]::UTF8)
try { $null = [xml]$ui } catch { throw "customUI.xml is not valid XML: $($_.Exception.Message)" }

Remove-Item $out -Force -ErrorAction SilentlyContinue
Copy-Item $raw $out -Force

$zip = [System.IO.Compression.ZipFile]::Open($out, 'Update')
try {
    Put-Entry $zip 'customUI/customUI.xml'   $ui
    Put-Entry $zip 'customUI/customUI14.xml' $ui

    $rels = Get-EntryText $zip '_rels/.rels'
    if ($rels -notmatch 'rIdCustomUI') {
        $add = '<Relationship Id="rIdCustomUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml" />' +
               '<Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml" />'
        Put-Entry $zip '_rels/.rels' ($rels -replace '(</Relationships>)', ($add + '$1'))
    }
    $ct = Get-EntryText $zip '[Content_Types].xml'
    if ($ct -notmatch 'Extension="xml"') {
        Put-Entry $zip '[Content_Types].xml' ($ct -replace '(<Types[^>]*>)', '$1<Default Extension="xml" ContentType="application/xml" />')
    }
}
finally { $zip.Dispose() }

Write-Host ("OK -> {0} ({1} bytes)" -f $out, (Get-Item $out).Length)