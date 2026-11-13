# Kiem chung (ASCII-only script; reads files as UTF-8 explicitly).
$ErrorActionPreference = 'Continue'
$d = (Split-Path -Parent $PSCommandPath)
$node = (Get-Command node -ErrorAction SilentlyContinue).Source
$conv = (Join-Path $d '..\..\build\vba2bas.mjs')
$out = New-Object System.Collections.Generic.List[string]

$out.Add('### LENH 1 - dem nut trong customUI.xml')
[xml]$x = Get-Content "$d\customUI.xml" -Raw -Encoding UTF8
$out.Add('buttons=' + $x.SelectNodes('//*[local-name()="button"]').Count)
$out.Add('buttons_via_tag=' + $x.GetElementsByTagName('button').Count)
$out.Add('bare_//button=' + $x.SelectNodes('//button').Count + '   (0 vi default xmlns khong duoc dang ky voi XPath)')
$acts = @($x.SelectNodes('//*[local-name()="button"]') | ForEach-Object { $_.onAction })
$out.Add('onAction=' + ($acts -join ', '))
$out.Add('tabLabel=' + $x.SelectSingleNode('//*[local-name()="tab"]').getAttribute('label'))
$out.Add('onLoad=' + $x.DocumentElement.getAttribute('onLoad'))

$out.Add('### LENH 2 - moi onAction phai co Sub tuong ung')
foreach ($n in $acts) {
  $hit = (Select-String -Path "$d\NTU.vba" -Pattern ("Sub\s+" + $n + "\s*\(") -Encoding UTF8).Count -gt 0
  $out.Add("$n -> $hit")
}

$out.Add('### LENH 3 - vba2bas.mjs')
$o = & $node $conv "$d\NTU.vba" "$env:TEMP\NTU_test.bas" 2>&1
foreach ($l in $o) { $out.Add([string]$l) }
$out.Add("exit=$LASTEXITCODE")

$out.Add('### LENH 4 - .bas sinh ra phai thuan ASCII')
$b = [IO.File]::ReadAllBytes("$env:TEMP\NTU_test.bas")
$bad = @($b | Where-Object { $_ -gt 127 })
$out.Add("bas_bytes=$($b.Length)  non-ascii=$($bad.Count)")

$out.Add('### LENH 5 - khong con ten goi cu')
$h = @(Select-String -Path "$d\NTU.vba" -Pattern 'OfficePartySuite|CuongBC195' -Encoding UTF8)
$out.Add("hits=$($h.Count)")
foreach ($l in $h) { $out.Add("$($l.LineNumber): $($l.Line.Trim())") }

$out.Add('### BO SUNG - tinh toan ven tep nguon')
$raw = [IO.File]::ReadAllBytes("$d\NTU.vba")
$s = [Text.Encoding]::UTF8.GetString($raw)
$out.Add("ntu_bytes=$($raw.Length)  lines=" + (Get-Content "$d\NTU.vba" -Encoding UTF8).Count)
$out.Add("bom=" + ($raw[0] -eq 0xEF) + "  first3=$($raw[0]),$($raw[1]),$($raw[2])")
$out.Add("high_codepoints=" + (($s.ToCharArray() | Where-Object { [int]$_ -gt 127 }).Count))
$out.Add("replacement_char=" + ([regex]::Matches($s, [string][char]0xFFFD).Count))
$procs = @(Select-String -Path "$d\NTU.vba" -Pattern '^\s*(Public|Private)\s+(Sub|Function)\s+(\w+)' -Encoding UTF8 | ForEach-Object { $_.Matches[0].Groups[3].Value })
$out.Add("procedures=" + $procs.Count)
$dups = @($procs | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
$out.Add("duplicate_procedures=" + $(if ($dups.Count) { $dups -join ',' } else { '(none)' }))
$out.Add("cu_xml_bytes=" + (Get-Item "$d\customUI.xml").Length)
$out.Add("APP_const=" + ((Select-String -Path "$d\NTU.vba" -Pattern 'Private Const APP' -Encoding UTF8).Line.Trim()))

$res = ($out -join "`r`n")
$res
$res | Set-Content "$env:TEMP\ntu_verify_out.txt" -Encoding UTF8