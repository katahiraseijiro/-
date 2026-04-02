param(
    [string]$ProductName = "Product",
    [string]$Target = "Target audience",
    [string]$Pain = "Main pain point",
    [string]$Benefit = "Main benefit",
    [string]$Offer = "Offer",
    [string]$Cta = "CTA",
    [string]$Tone = "Friendly and trustworthy",
    [string]$OutputCsv = "canva_generated_copy.csv"
)

$ErrorActionPreference = "Stop"

$templatePath = Join-Path $PSScriptRoot "canva_copy_prompt.txt"
if (!(Test-Path $templatePath)) {
    throw "Prompt file not found: $templatePath"
}

$prompt = Get-Content -Path $templatePath -Raw
$prompt = $prompt.Replace("{商品名}", $ProductName)
$prompt = $prompt.Replace("{ターゲット}", $Target)
$prompt = $prompt.Replace("{悩み}", $Pain)
$prompt = $prompt.Replace("{ベネフィット}", $Benefit)
$prompt = $prompt.Replace("{オファー}", $Offer)
$prompt = $prompt.Replace("{CTA}", $Cta)
$prompt = $prompt.Replace("{トーン}", $Tone)

# Pass prompt to Claude Code CLI and save CSV
$csv = & "$env:USERPROFILE\.local\bin\claude.exe" --print $prompt
if ([string]::IsNullOrWhiteSpace($csv)) {
    throw "Claude returned empty output."
}
if ($csv -match "Not logged in") {
    throw "Claude CLI is not logged in. Run 'claude' once and complete login, then rerun this script."
}
if ($csv -notmatch "^""Headline"",""Subheadline"",""Body"",""CTA"",""Hook"",""VisualIdea""") {
    throw "Claude output is not in expected CSV format. Check your prompt/template."
}

$outPath = Join-Path $PSScriptRoot $OutputCsv
[System.IO.File]::WriteAllText($outPath, $csv, [System.Text.UTF8Encoding]::new($false))
Write-Host "Generated:" $outPath
