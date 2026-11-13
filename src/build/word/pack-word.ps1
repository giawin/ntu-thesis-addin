# Inject the personalized ribbon into the .dotm produced by build-word.vbs.
# Word keeps customUI out of its own save path, so this must run after every build.
# This file is intentionally pure ASCII: Windows PowerShell 5.1 reads non-ASCII
# .ps1 as ANSI and breaks the parser.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$here     = $PSScriptRoot
$repo     = (Resolve-Path (Join-Path $here '..\..\..')).Path
$dist     = Join-Path $repo 'src\core\word\dist'
$out      = Join-Path $dist 'GiaTwinH_NTU_Thesis_Word.dotm'
$ribbon   = Join-Path $repo 'src\core\word\customUI.xml'
$upstream = Join-Path $repo 'src\core\vendor\upstream\ChuanHoa_VanBanDang.dotm'

if (-not (Test-Path $out))      { throw "missing $out - run build-word.vbs first" }
if (-not (Test-Path $ribbon))   { throw "missing $ribbon" }
if (-not (Test-Path $upstream)) { throw "missing upstream reference $upstream" }

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

# The building blocks (glossary) carry the sample-document templates used by the
# Ribbon menus. If Word dropped them during SaveAs2, copy them back from the
# original byte for byte so the menu behaviour stays identical.
$need = @()
$z = [System.IO.Compression.ZipFile]::OpenRead($out)
try {
    foreach ($n in @('word/glossary/document.xml','word/glossary/_rels/document.xml.rels','word/glossary/settings.xml','word/glossary/styles.xml','word/glossary/fontTable.xml','word/glossary/webSettings.xml')) {
        if (-not $z.GetEntry($n)) { $need += $n }
    }
}
finally { $z.Dispose() }

if ($need.Count -gt 0) {
    Write-Host ("restoring {0} glossary part(s) from the original" -f $need.Count)
    $src = [System.IO.Compression.ZipFile]::OpenRead($upstream)
    $dst = [System.IO.Compression.ZipFile]::Open($out, 'Update')
    try {
        foreach ($n in $need) {
            $e = $src.GetEntry($n)
            if (-not $e) { Write-Warning "upstream also lacks $n"; continue }
            $ne = $dst.CreateEntry($n, [System.IO.Compression.CompressionLevel]::Optimal)
            $inS = $e.Open(); $outS = $ne.Open()
            $inS.CopyTo($outS)
            $outS.Close(); $inS.Close()
        }
        # glossary relationships must be registered in the document rels
        $drels = Get-EntryText $dst 'word/_rels/document.xml.rels'
        if ($drels -and $drels -notmatch 'glossary') {
            $add = '<Relationship Id="rIdGlossary" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml" />'
            Put-Entry $dst 'word/_rels/document.xml.rels' ($drels -replace '(</Relationships>)', ($add + '$1'))
        }
        $dct = Get-EntryText $dst '[Content_Types].xml'
        if ($dct -and $dct -notmatch 'glossaryDocument') {
            $ov = '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml" />'
            Put-Entry $dst '[Content_Types].xml' ($dct -replace '(</Types>)', ($ov + '$1'))
        }
    }
    finally { $dst.Dispose(); $src.Dispose() }
}

$size = (Get-Item $out).Length
Write-Host ("OK -> {0} ({1} bytes)" -f $out, $size)
$z2 = [System.IO.Compression.ZipFile]::OpenRead($out)
try { $z2.Entries | Select-Object -Expand FullName } finally { $z2.Dispose() }
