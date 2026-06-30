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
    'Get-NormalizedPath',
    'Test-ProtectedFolder',
    'Test-PathWithinRoot',
    'Get-DefaultUserProfileRoot',
    'Test-ProfileTargetPath',
    'Get-ProfileProgramsPath',
    'New-StartMenuScope',
    'Get-ConfiguredProfileScope',
    'Get-DefaultUserScope',
    'Get-StartMenuScopes',
    'Get-StartMenuPaths',
    'Get-ApprovedMutationRoots',
    'Test-MutationRootRequiresAdministrator',
    'Get-ApprovedMutationRoot',
    'Ensure-JournalStorage',
    'New-OperationId',
    'New-UndoBackupCopy',
    'Invoke-GuardedFileOperation',
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
        [string]$Level = 'Info',
        [string]$OperationId = $null
    )
}

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerProfile_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $userRoot = Join-Path $testRoot 'current-user\Programs'
    $systemRoot = Join-Path $testRoot 'system\Programs'
    $profileRoot = Join-Path $testRoot 'profiles\Alice'
    $defaultRoot = Join-Path $testRoot 'profiles\Default'
    $profilePrograms = Join-Path $profileRoot 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs'
    $defaultPrograms = Join-Path $defaultRoot 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs'
    $undoRoot = Join-Path $testRoot 'undo'

    foreach ($path in @($userRoot, $systemRoot, $profilePrograms, $defaultPrograms, $undoRoot)) {
        [System.IO.Directory]::CreateDirectory($path) | Out-Null
    }

    $script:Config = @{
        UserStartMenu = $userRoot
        SystemStartMenu = $systemRoot
        ProfileRoot = $profileRoot
        DefaultProfileRoot = $defaultRoot
        UndoFile = Join-Path $undoRoot 'undo.json'
        UndoBackupRoot = Join-Path $undoRoot 'UndoBackups'
    }
    $script:ProtectedFolders = @('Startup')

    $resolvedProfile = Get-ProfileProgramsPath -ProfileRootOrProgramsPath $profileRoot
    if ((Get-NormalizedPath $resolvedProfile.ProgramsPath) -ne (Get-NormalizedPath $profilePrograms)) {
        throw "Failed: profile root did not resolve to Programs path: $($resolvedProfile.ProgramsPath)"
    }
    $resolvedDirect = Get-ProfileProgramsPath -ProfileRootOrProgramsPath $profilePrograms
    if ((Get-NormalizedPath $resolvedDirect.ProfileRoot) -ne (Get-NormalizedPath $profileRoot)) {
        throw 'Failed: direct Programs path did not resolve back to the profile root.'
    }

    $driveRootRejected = $false
    try {
        Get-ProfileProgramsPath -ProfileRootOrProgramsPath ([System.IO.Path]::GetPathRoot($testRoot)) | Out-Null
    }
    catch {
        $driveRootRejected = $true
    }
    if (-not $driveRootRejected) {
        throw 'Failed: drive root was accepted as a profile root.'
    }

    $cmbScope = [PSCustomObject]@{ SelectedIndex = 3 }
    $profileScope = @(Get-StartMenuScopes)[0]
    if ($profileScope.ScopeName -ne 'Profile' -or
        $profileScope.BackupName -ne 'Profile' -or
        -not $profileScope.RequiresAdmin -or
        (Get-NormalizedPath $profileScope.Path) -ne (Get-NormalizedPath $profilePrograms)) {
        throw 'Failed: selected profile scope metadata was incorrect.'
    }

    $cmbScope.SelectedIndex = 4
    $defaultScope = @(Get-StartMenuScopes)[0]
    if ($defaultScope.ScopeName -ne 'DefaultUser' -or
        $defaultScope.BackupName -ne 'DefaultUser' -or
        -not $defaultScope.RequiresAdmin -or
        (Get-NormalizedPath $defaultScope.Path) -ne (Get-NormalizedPath $defaultPrograms)) {
        throw 'Failed: default user scope metadata was incorrect.'
    }

    $approvedRoots = @(Get-ApprovedMutationRoots)
    if ($approvedRoots -notcontains (Get-NormalizedPath $profilePrograms)) {
        throw 'Failed: approved mutation roots did not include the selected profile Start Menu.'
    }
    if ($approvedRoots -notcontains (Get-NormalizedPath $defaultPrograms)) {
        throw 'Failed: approved mutation roots did not include the default user Start Menu.'
    }

    $profileShortcut = Join-Path $profilePrograms 'delete-me.lnk'
    Set-Content -LiteralPath $profileShortcut -Value 'shortcut'
    $script:IsAdmin = $false
    $denied = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $profileShortcut
    if ($denied.Result -ne 'Failed' -or $denied.Message -notmatch 'Administrator privileges') {
        throw "Failed: non-admin profile mutation was not denied. $($denied.Result) $($denied.Message)"
    }
    if (-not (Test-Path -LiteralPath $profileShortcut)) {
        throw 'Failed: denied profile mutation removed the source file.'
    }

    $script:IsAdmin = $true
    $allowed = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $profileShortcut
    if ($allowed.Result -ne 'Success' -or (Test-Path -LiteralPath $profileShortcut)) {
        throw 'Failed: elevated profile mutation did not succeed.'
    }

    Set-Content -LiteralPath (Join-Path $profilePrograms 'current.lnk') -Value 'current'
    $backupPath = Join-Path $testRoot 'backup\Profile'
    [System.IO.Directory]::CreateDirectory($backupPath) | Out-Null
    Set-Content -LiteralPath (Join-Path $backupPath 'restored.lnk') -Value 'restored'
    $stagingRoot = Join-Path $testRoot 'staging'
    $rollbackRoot = Join-Path $testRoot 'rollback'
    [System.IO.Directory]::CreateDirectory($stagingRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($rollbackRoot) | Out-Null

    $plan = New-RestorePlan -ScopeName 'Profile' -BackupPath $backupPath -TargetPath $profilePrograms -StagingRoot $stagingRoot -RollbackRoot $rollbackRoot
    $result = Restore-DirectoryFromPlan -Plan $plan
    if (-not $result.Success) {
        throw "Failed: profile restore failed: $($result.Message)"
    }
    if (-not (Test-Path -LiteralPath (Join-Path $profilePrograms 'restored.lnk'))) {
        throw 'Failed: profile restore did not copy backup content.'
    }
    if (-not (Test-Path -LiteralPath (Join-Path $plan.RollbackPath 'current.lnk'))) {
        throw 'Failed: profile restore did not preserve rollback content.'
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Profile target tests passed.'
