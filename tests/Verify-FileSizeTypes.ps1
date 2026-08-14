$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot '..\FormatCore.ps1'
$source = Get-Content -LiteralPath $scriptPath -Raw

if ($source -match '\[ValidateRange\(1024, 2147483647\)\]\[int\]\$Threshold') {
    throw 'Threshold sigue limitado a Int32.'
}

foreach ($pattern in @('\[long\]\$Threshold', '\[long\]\$OriginalSize')) {
    if ($source -notmatch $pattern) { throw "Falta el tipo long requerido: $pattern" }
}

if ($source -match 'param\(\[string\]\$FilePath,\[int\]\$OriginalSize') {
    throw 'OriginalSize no debe ser int.'
}

Write-Output 'PASS: los tamanos de archivo y umbrales usan Int64.'
