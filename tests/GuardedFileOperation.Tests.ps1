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
    'Get-ApprovedMutationRoots',
    'Test-MutationRootRequiresAdministrator',
    'Get-ApprovedMutationRoot',
    'Test-ReparsePoint',
    'Test-PathContainsReparsePoint',
    'Ensure-JournalStorage',
    'New-OperationId',
    'New-UndoBackupCopy',
    'Invoke-GuardedFileOperation'
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

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerGuardedOps_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $userRoot = Join-Path $testRoot 'user'
    $systemRoot = Join-Path $testRoot 'system'
    $undoRoot = Join-Path $testRoot 'undo'
    [System.IO.Directory]::CreateDirectory($userRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($systemRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($undoRoot) | Out-Null

    $script:Config = @{
        UserStartMenu = $userRoot
        SystemStartMenu = $systemRoot
        UndoFile = Join-Path $undoRoot 'undo.json'
        UndoBackupRoot = Join-Path $undoRoot 'UndoBackups'
    }
    $script:ProtectedFolders = @('Startup')

    $previewPath = Join-Path $userRoot 'preview-delete.lnk'
    Set-Content -LiteralPath $previewPath -Value 'preview'
    $preview = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $previewPath -Preview $true
    if ($preview.Result -ne 'Preview' -or -not (Test-Path -LiteralPath $previewPath)) {
        throw 'Failed: preview delete did not leave the source intact.'
    }

    $crossRootSource = Join-Path $userRoot 'cross-root.lnk'
    $crossRootDestination = Join-Path $systemRoot 'cross-root.lnk'
    Set-Content -LiteralPath $crossRootSource -Value 'cross-root'
    $crossRoot = Invoke-GuardedFileOperation -Action 'Move' -SourcePath $crossRootSource -DestinationPath $crossRootDestination
    if ($crossRoot.Result -ne 'Failed' -or $crossRoot.Message -notmatch 'same approved root') {
        throw "Failed: cross-root move was not rejected. Result: $($crossRoot.Result) $($crossRoot.Message)"
    }

    $collisionSource = Join-Path $userRoot 'collision-source.lnk'
    $collisionDestination = Join-Path $userRoot 'collision-destination.lnk'
    Set-Content -LiteralPath $collisionSource -Value 'source'
    Set-Content -LiteralPath $collisionDestination -Value 'destination'
    $collision = Invoke-GuardedFileOperation -Action 'Move' -SourcePath $collisionSource -DestinationPath $collisionDestination
    if ($collision.Result -ne 'Failed' -or $collision.Message -notmatch 'Destination already exists') {
        throw "Failed: collision was not rejected. Result: $($collision.Result) $($collision.Message)"
    }
    if (-not (Test-Path -LiteralPath $collisionSource)) {
        throw 'Failed: collision failure moved the source.'
    }

    $deletePath = Join-Path $userRoot 'delete-with-backup.lnk'
    Set-Content -LiteralPath $deletePath -Value 'delete'
    $delete = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $deletePath -OperationId (New-OperationId) -RegisterRollback
    if ($delete.Result -ne 'Success') {
        throw "Failed: guarded delete did not succeed. $($delete.Message)"
    }
    if (Test-Path -LiteralPath $deletePath) {
        throw 'Failed: guarded delete left the source in place.'
    }
    if (-not (Test-Path -LiteralPath $delete.BackupPath)) {
        throw 'Failed: guarded delete did not create a rollback backup.'
    }

    $protectedPath = Join-Path $userRoot 'Startup\do-not-delete.lnk'
    [System.IO.Directory]::CreateDirectory((Split-Path $protectedPath -Parent)) | Out-Null
    Set-Content -LiteralPath $protectedPath -Value 'protected'
    $protectedDelete = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $protectedPath
    if ($protectedDelete.Result -ne 'Failed' -or $protectedDelete.Message -notmatch 'protected') {
        throw "Failed: protected delete was not rejected. $($protectedDelete.Result) $($protectedDelete.Message)"
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Guarded file operation tests passed.'
