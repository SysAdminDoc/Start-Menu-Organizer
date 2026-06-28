#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'StartMenuOrganizerPro.ps1'
$content = Get-Content -LiteralPath $scriptPath -Raw
$tokens = $null
$parseErrors = $null
[System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors) | Out-Null
if ($parseErrors.Count -gt 0) {
    $messages = $parseErrors | ForEach-Object { $_.Message }
    throw "Parser errors found:`n$($messages -join "`n")"
}

foreach ($requiredFunction in @('Start-BackgroundWorker', 'Stop-BackgroundWorker', 'Set-WorkerUiState', 'Refresh-Items')) {
    if ($content -notmatch "function\s+$requiredFunction\b") {
        throw "Failed: missing $requiredFunction."
    }
}

if ($content -notmatch 'Start-BackgroundWorker\s+-Name\s+"Scanning Start Menu\.\.\."') {
    throw 'Failed: Refresh-Items does not dispatch the scan through Start-BackgroundWorker.'
}

if ($content -notmatch '\$btnCancelWork\.Add_Click\(\{\s*Stop-BackgroundWorker\s*\}\)') {
    throw 'Failed: Cancel button is not wired to Stop-BackgroundWorker.'
}

if ($content -notmatch '\$Window\.Dispatcher\.Invoke') {
    throw 'Failed: worker completion does not marshal UI updates through the dispatcher.'
}

'Async scan tests passed.'
