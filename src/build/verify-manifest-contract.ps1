# Manifest contract check: every key the VBA add-ins actually read must be
# produced by the release workflow, and the workflow must not emit keys nobody
# reads.
#
# Why this exists: v1.0.0 shipped `xlam_url` while the original Excel add-in
# reads `excel_url`. Nothing crashed -- the manifest was still valid JSON, the
# key was simply empty, and Excel's "Cap nhat" button died silently. This turns
# that class of bug into a red CI run.
#
# Keys ending in `_win` are optional: the VBA tries them first and falls back to
# the plain key when they are absent, so they are reported but never fatal.
#
# Run: powershell.exe -NoProfile -ExecutionPolicy Bypass -File src/build/verify-manifest-contract.ps1

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSCommandPath
$repo = Split-Path -Parent (Split-Path -Parent $here)

function Fail([string]$msg) {
  Write-Host "FAIL: $msg" -ForegroundColor Red
  exit 1
}

# --- 1. Keys read by the VBA, extracted from source (never hand-copied) ---
# Every add-in: Word core, NTU module, and the untouched upstream Excel (which
# reads excel_* and carries the author license API keys we must ignore).
$vbaFiles = @(
  (Join-Path $repo 'src/core/word/Module1.vba'),
  (Join-Path $repo 'src/modules/ntu-thesis/NTU.vba'),
  (Join-Path $repo 'src/core/excel/Module1.upstream.bas')
)
$vbaKeys = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($f in $vbaFiles) {
  if (-not (Test-Path -LiteralPath $f)) { Fail "missing $f" }
  $text = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
  # ParseJsonField(js, "key")  and  NTU_ParseJson(js, "key"). Restrict to the
  # manifest namespace: the same helper also parses the author license API
  # replies (status/message/khoa/key), which are not manifest fields.
  foreach ($m in [regex]::Matches($text, '(?:ParseJsonField|NTU_ParseJson)\s*\([^,]+,\s*"(version|notes_ascii|(?:word|ntu|excel)_[a-z0-9_]+)"')) {
    [void]$vbaKeys.Add($m.Groups[1].Value)
  }
}
if ($vbaKeys.Count -eq 0) { Fail 'extracted no keys from VBA - the regex has drifted' }

# --- 2. Keys produced by the release workflow (parse the @{...} block) ---
$releaseYml = Join-Path $repo '.github/workflows/release.yml'
if (-not (Test-Path -LiteralPath $releaseYml)) { Fail "missing $releaseYml" }
$y = [IO.File]::ReadAllText($releaseYml, [Text.Encoding]::UTF8)

$block = [regex]::Match($y, '\$man\s*=\s*\[ordered\]@\{(?<body>[\s\S]*?)\n\s*\}')
if (-not $block.Success) { Fail 'could not find the $man = [ordered]@{...} block in release.yml' }

$manKeys = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($m in [regex]::Matches($block.Groups['body'].Value, '(?m)^\s*([a-z0-9_]+)\s*=')) {
  [void]$manKeys.Add($m.Groups[1].Value)
}
if ($manKeys.Count -eq 0) { Fail 'parsed the manifest block but found no keys' }

# --- 3. Required keys: everything the VBA reads, minus optional _win fallbacks ---
$required = @($vbaKeys | Where-Object { $_ -notmatch '_win$' } | Sort-Object)
$missing = @($required | Where-Object { -not $manKeys.Contains($_) })
if ($missing.Count -gt 0) {
  Write-Host 'VBA reads these keys but the manifest does not produce them:' -ForegroundColor Red
  $missing | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
  Fail "manifest is missing $($missing.Count) required key(s) - auto-update would break"
}

# --- 4. No key nobody reads (catches misnamed keys such as xlam_*) ---
$extra = @($manKeys | Where-Object { -not $vbaKeys.Contains($_) } | Sort-Object)
if ($extra.Count -gt 0) {
  Write-Host 'Manifest produces keys no add-in reads:' -ForegroundColor Red
  $extra | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
  Fail 'unexpected key - usually a misnamed key (e.g. xlam_* instead of excel_*)'
}

# --- 5. The in-workflow verification must cover every required key ---
$guard = [regex]::Match($y, 'foreach\s*\(\$k in (?<list>[^\r\n]+)\)')
if (-not $guard.Success) { Fail 'could not find the key guard loop in release.yml' }
$guarded = [regex]::Matches($guard.Groups['list'].Value, "'([a-z0-9_]+)'") | ForEach-Object { $_.Groups[1].Value }
$unguarded = @($required | Where-Object { $guarded -notcontains $_ })
if ($unguarded.Count -gt 0) {
  Write-Host 'The release guard loop does not check these keys:' -ForegroundColor Red
  $unguarded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
  Fail 'the guard must cover every required key, otherwise an empty value ships'
}

Write-Host "OK: $($required.Count) required keys, $($manKeys.Count) produced - contract matches."
Write-Host "    required: $($required -join ', ')"
if ($vbaKeys.Count -ne $required.Count) {
  $opt = @($vbaKeys | Where-Object { $_ -match '_win$' } | Sort-Object)
  Write-Host "    optional fallbacks (not required): $($opt -join ', ')"
}
exit 0