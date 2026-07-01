#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'StartMenuOrganizerPro.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    $messages = $parseErrors | ForEach-Object { $_.Message }
    throw "Parser errors found:`n$($messages -join "`n")"
}

$helperFunctions = @(
    'Get-ConfigurationSnapshot',
    'Set-ObservableCollection',
    'Apply-ConfigurationSnapshot',
    'Write-AtomicFile',
    'Read-WithBackupFallback',
    'Save-ApplicationConfiguration',
    'Load-ApplicationConfiguration'
)

foreach ($functionName in $helperFunctions) {
    $functionAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq $functionName
    }, $true)

    if (-not $functionAst) {
        throw "Failed: helper function not found: $functionName"
    }

    Invoke-Expression $functionAst.Extent.Text
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'Info',
        [string]$OperationId = $null
    )
}

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerSettings_$([System.Guid]::NewGuid().ToString('N'))"
try {
    [System.IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $script:Config = @{
        SettingsSchema = 1
        ConfigFile = Join-Path $testRoot 'config.json'
    }
    $script:IsLoadingConfiguration = $false
    $cmbScope = [PSCustomObject]@{ SelectedIndex = 3 }
    $txtProfileRoot = [PSCustomObject]@{ Text = (Join-Path $testRoot 'Profiles\Alice') }
    $script:JunkPatterns = [System.Collections.ObjectModel.ObservableCollection[string]]@('*junk*')
    $script:ProtectedFolders = [System.Collections.ObjectModel.ObservableCollection[string]]@('Startup')
    $script:Categories = [ordered]@{
        Utilities = @('Tool*')
    }

    Save-ApplicationConfiguration
    if (-not (Test-Path -LiteralPath $script:Config.ConfigFile)) {
        throw 'Failed: config file was not written.'
    }

    $saved = Get-Content -LiteralPath $script:Config.ConfigFile -Raw | ConvertFrom-Json
    if ($saved.SchemaVersion -ne 1 -or $saved.ScopeIndex -ne 3 -or $saved.ProfileRoot -ne $txtProfileRoot.Text) {
        throw 'Failed: saved config schema, scope, or profile root was incorrect.'
    }

    $script:JunkPatterns = [System.Collections.ObjectModel.ObservableCollection[string]]@('*changed*')
    $script:ProtectedFolders = [System.Collections.ObjectModel.ObservableCollection[string]]@('Changed')
    $script:Categories = [ordered]@{ Changed = @('Changed*') }
    $cmbScope.SelectedIndex = 0
    $txtProfileRoot.Text = ''

    Load-ApplicationConfiguration
    if ($script:JunkPatterns[0] -ne '*junk*') {
        throw 'Failed: junk patterns were not loaded.'
    }
    if ($script:ProtectedFolders[0] -ne 'Startup') {
        throw 'Failed: protected folders were not loaded.'
    }
    if (-not $script:Categories.Contains('Utilities')) {
        throw 'Failed: categories were not loaded.'
    }
    if ($cmbScope.SelectedIndex -ne 3) {
        throw 'Failed: saved scope was not loaded.'
    }
    if ($txtProfileRoot.Text -ne (Join-Path $testRoot 'Profiles\Alice')) {
        throw 'Failed: saved profile root was not loaded.'
    }

    Set-Content -LiteralPath $script:Config.ConfigFile -Value '{ invalid json'
    Load-ApplicationConfiguration
    $badFiles = @(Get-ChildItem -LiteralPath $testRoot -Filter 'config.json.bad.*')
    if ($badFiles.Count -ne 1) {
        throw 'Failed: invalid config was not moved aside.'
    }

    # --- Atomic write creates .bak and recovers from corrupt primary ---
    $script:JunkPatterns = [System.Collections.ObjectModel.ObservableCollection[string]]@('*original*')
    $script:Config.ConfigFile = Join-Path $testRoot 'atomic-config.json'
    Save-ApplicationConfiguration
    if (-not (Test-Path -LiteralPath "$($script:Config.ConfigFile).bak" -ErrorAction SilentlyContinue)) {
        Save-ApplicationConfiguration
    }
    $bakExists = Test-Path -LiteralPath "$($script:Config.ConfigFile).bak"
    if (-not $bakExists) {
        throw 'Failed: .bak file was not created after second save.'
    }

    Set-Content -LiteralPath $script:Config.ConfigFile -Value '{ truncated'
    $script:JunkPatterns = [System.Collections.ObjectModel.ObservableCollection[string]]@('*overwritten*')
    Load-ApplicationConfiguration
    if ($script:JunkPatterns[0] -ne '*original*') {
        throw "Failed: backup fallback did not recover original config. Got: $($script:JunkPatterns[0])"
    }

    # --- Atomic write cleans up .tmp on failure ---
    $tmpPath = "$($script:Config.ConfigFile).tmp"
    if (Test-Path -LiteralPath $tmpPath) {
        throw 'Failed: orphan .tmp file exists before test.'
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Settings persistence tests passed.'
