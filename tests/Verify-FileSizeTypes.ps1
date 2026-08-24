$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot '..\FormatCore.ps1'
$source = Get-Content -LiteralPath $scriptPath -Raw

if ($source -match '\[ValidateRange\(1024, 2147483647\)\]\[int\]\$Threshold') {
    throw 'Threshold is still limited to Int32.'
}

foreach ($pattern in @('\[long\]\$Threshold', '\[long\]\$OriginalSize')) {
    if ($source -notmatch $pattern) { throw "Missing required long type: $pattern" }
}

if ($source -match 'param\(\[string\]\$FilePath,\[int\]\$OriginalSize') {
    throw 'OriginalSize must not be int.'
}

Write-Output 'PASS: file sizes and thresholds use Int64.'
