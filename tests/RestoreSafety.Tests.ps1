#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'StartMenuOrganizerPro.ps1'
$content = Get-Content -LiteralPath $scriptPath -Raw

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    $messages = $parseErrors | ForEach-Object { $_.Message }
    throw "Parser errors found:`n$($messages -join "`n")"
}

function Assert-Contains {
    param(
        [string]$Name,
        [string]$Text,
        [string]$Pattern
    )

    if ($Text -notmatch $Pattern) {
        throw "Failed: $Name"
    }
}

function Assert-NotContains {
    param(
        [string]$Name,
        [string]$Text,
        [string]$Pattern
    )

    if ($Text -match $Pattern) {
        throw "Failed: $Name"
    }
}

$restoreStart = $content.IndexOf('function Restore-Backup')
$restoreEnd = $content.IndexOf('function Export-Configuration')
if ($restoreStart -lt 0 -or $restoreEnd -lt $restoreStart) {
    throw 'Failed: Restore-Backup function boundaries not found.'
}

$restoreBody = $content.Substring($restoreStart, $restoreEnd - $restoreStart)

Assert-NotContains 'Restore-Backup must not delete user/system Start Menu with wildcard paths.' $restoreBody 'Remove-Item\s+-Path\s+"\$\(\$Config\.(UserStartMenu|SystemStartMenu)\)\\\*"'
Assert-NotContains 'Restore-Backup must not copy backup contents with wildcard paths.' $restoreBody 'Copy-Item\s+-Path\s+"\$(userBackup|systemBackup)\\\*"'
Assert-Contains 'Restore-Backup must build validated restore plans before applying them.' $restoreBody 'New-RestorePlan'
Assert-Contains 'Restore-Backup must execute the reviewed restore plan.' $restoreBody 'Restore-DirectoryFromPlan'
Assert-Contains 'Restore-Backup must preserve a pre-restore rollback snapshot.' $restoreBody '_restore_rollback'
Assert-Contains 'Restore execution must attempt rollback on failure.' $content 'Invoke-RestoreRollback\s+-Plan\s+\$Plan'
Assert-Contains 'Directory clearing must use literal paths.' $content 'Remove-Item\s+-LiteralPath\s+\$child\.FullName'
Assert-Contains 'Restore planning must stage backup sources.' $content 'StagedSource\s*=\s*\$stagedSource'
Assert-Contains 'Restore planning must persist rollback paths.' $content 'RollbackPath\s*=\s*\$rollbackPath'

$helperFunctions = @(
    'Get-NormalizedPath',
    'Test-ApprovedRestoreTarget',
    'Copy-DirectoryContents',
    'Clear-DirectoryContents',
    'New-RestorePlan',
    'Invoke-RestoreRollback',
    'Restore-DirectoryFromPlan'
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
        [string]$Level = 'Info'
    )
}

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerRestoreSafety_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $targetPath = Join-Path $testRoot 'live\User'
    $systemPath = Join-Path $testRoot 'live\System'
    $backupPath = Join-Path $testRoot 'backup\User'
    $stagingRoot = Join-Path $testRoot 'staging'
    $rollbackRoot = Join-Path $testRoot 'rollback'

    [System.IO.Directory]::CreateDirectory($targetPath) | Out-Null
    [System.IO.Directory]::CreateDirectory($systemPath) | Out-Null
    [System.IO.Directory]::CreateDirectory($backupPath) | Out-Null
    [System.IO.Directory]::CreateDirectory($stagingRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($rollbackRoot) | Out-Null

    Set-Content -LiteralPath (Join-Path $targetPath 'current.lnk') -Value 'current'
    Set-Content -LiteralPath (Join-Path $backupPath 'restored.lnk') -Value 'restored'

    $script:Config = @{
        UserStartMenu = $targetPath
        SystemStartMenu = $systemPath
    }

    $plan = New-RestorePlan -ScopeName 'User' -BackupPath $backupPath -TargetPath $targetPath -StagingRoot $stagingRoot -RollbackRoot $rollbackRoot
    $result = Restore-DirectoryFromPlan -Plan $plan

    if (-not $result.Success) {
        throw "Failed: restore helper returned failure: $($result.Message)"
    }
    if (-not (Test-Path -LiteralPath (Join-Path $targetPath 'restored.lnk'))) {
        throw 'Failed: restored file was not copied to target.'
    }
    if (Test-Path -LiteralPath (Join-Path $targetPath 'current.lnk')) {
        throw 'Failed: previous live file remained after restore.'
    }
    if (-not (Test-Path -LiteralPath (Join-Path $plan.RollbackPath 'current.lnk'))) {
        throw 'Failed: rollback snapshot did not preserve previous live file.'
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Restore safety tests passed.'
