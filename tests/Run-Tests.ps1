#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'StartMenuOrganizerPro.ps1'
$settingsPath = Join-Path $repoRoot 'PSScriptAnalyzerSettings.psd1'

$tokens = $null
$parseErrors = $null
[System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors) | Out-Null
if ($parseErrors.Count -gt 0) {
    $messages = $parseErrors | ForEach-Object { $_.Message }
    throw "Parser errors found:`n$($messages -join "`n")"
}

$plainTests = @(
    'RestoreSafety.Tests.ps1',
    'UndoJournal.Tests.ps1',
    'GuardedFileOperation.Tests.ps1',
    'OperationPlan.Tests.ps1',
    'AsyncScan.Tests.ps1',
    'ShortcutValidation.Tests.ps1',
    'SettingsPersistence.Tests.ps1',
    'Packaging.Tests.ps1',
    'Logging.Tests.ps1',
    'AccessibilityLocalization.Tests.ps1',
    'ProfileTarget.Tests.ps1',
    'SafetyGuards.Tests.ps1'
)

foreach ($test in $plainTests) {
    & (Join-Path $PSScriptRoot $test) | Out-Host
}

Invoke-Pester -Path (Join-Path $PSScriptRoot 'StartMenuOrganizer.Pester.Tests.ps1') -CI

$analysis = @(Invoke-ScriptAnalyzer -Path $scriptPath -Settings $settingsPath)
if ($analysis.Count -gt 0) {
    $analysis | Format-Table -AutoSize | Out-String | Write-Error
    throw "PSScriptAnalyzer reported $($analysis.Count) finding(s)."
}

'All local tests passed.'
