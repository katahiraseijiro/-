# CSV から pptx を生成（Python 3 必須）
param(
    [Parameter(Mandatory = $true)][string]$InputCsv,
    [Parameter(Mandatory = $true)][string]$Output,
    [string]$Title = ""
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

$python = $null
foreach ($name in @("python", "python3")) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if ($c) { $python = $name; break }
}
if (-not $python) {
    Write-Error "Python 3 が PATH にありません。インストール後に再実行してください。"
}

& $python -m pip install -q -r (Join-Path $here "requirements-pptx.txt")
$py = Join-Path $here "csv_to_pptx.py"
if ($Title) {
    & $python $py -i $InputCsv -o $Output --title $Title
} else {
    & $python $py -i $InputCsv -o $Output
}
