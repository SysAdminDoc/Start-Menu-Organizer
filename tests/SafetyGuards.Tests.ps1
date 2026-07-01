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
    'Test-PathWithinRoot',
    'Test-ReparsePoint',
    'Test-PathContainsReparsePoint',
    'Test-ProtectedFolder',
    'Get-DefaultUserProfileRoot',
    'Test-ProfileTargetPath',
    'Get-ProfileProgramsPath',
    'New-StartMenuScope',
    'Get-ConfiguredProfileScope',
    'Get-DefaultUserScope',
    'Get-ApprovedMutationRoots',
    'Test-MutationRootRequiresAdministrator',
    'Get-ApprovedMutationRoot',
    'Ensure-JournalStorage',
    'New-OperationId',
    'New-UndoBackupCopy',
    'Invoke-GuardedFileOperation',
    'Test-ApprovedRestoreTarget',
    'Copy-DirectoryContents'
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

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerSafetyGuards_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $userRoot = Join-Path $testRoot 'user'
    $systemRoot = Join-Path $testRoot 'system'
    $undoRoot = Join-Path $testRoot 'undo'
    $outsideDir = Join-Path $testRoot 'outside'
    [System.IO.Directory]::CreateDirectory($userRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($systemRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($undoRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($outsideDir) | Out-Null

    $script:Config = @{
        UserStartMenu   = $userRoot
        SystemStartMenu = $systemRoot
        UndoFile        = Join-Path $undoRoot 'undo.json'
        UndoBackupRoot  = Join-Path $undoRoot 'UndoBackups'
        ProfileRoot     = ''
        DefaultProfileRoot = ''
    }
    $script:ProtectedFolders = @('Startup')
    $script:IsAdmin = $true

    # --- Test-ReparsePoint on normal file ---
    $normalFile = Join-Path $userRoot 'normal.lnk'
    Set-Content -LiteralPath $normalFile -Value 'test'
    if (Test-ReparsePoint $normalFile) {
        throw 'Failed: Test-ReparsePoint returned true for a normal file.'
    }

    # --- Test-ReparsePoint on normal directory ---
    $normalDir = Join-Path $userRoot 'normaldir'
    [System.IO.Directory]::CreateDirectory($normalDir) | Out-Null
    if (Test-ReparsePoint $normalDir) {
        throw 'Failed: Test-ReparsePoint returned true for a normal directory.'
    }

    # --- Test-ReparsePoint on junction ---
    $junctionTarget = Join-Path $outsideDir 'target'
    [System.IO.Directory]::CreateDirectory($junctionTarget) | Out-Null
    Set-Content -LiteralPath (Join-Path $junctionTarget 'secret.txt') -Value 'sensitive data'

    $junctionPath = Join-Path $userRoot 'junction-link'
    cmd /c mklink /J "$junctionPath" "$junctionTarget" 2>&1 | Out-Null
    if (-not (Test-Path -LiteralPath $junctionPath)) {
        throw 'Failed: could not create junction for testing.'
    }

    if (-not (Test-ReparsePoint $junctionPath)) {
        throw 'Failed: Test-ReparsePoint returned false for a junction.'
    }

    # --- Test-PathContainsReparsePoint detects junction in path ---
    $fileUnderJunction = Join-Path $junctionPath 'secret.txt'
    if (-not (Test-PathContainsReparsePoint -Path $fileUnderJunction -Root $userRoot)) {
        throw 'Failed: Test-PathContainsReparsePoint did not detect junction in path.'
    }

    # --- Test-PathContainsReparsePoint returns false for normal path ---
    if (Test-PathContainsReparsePoint -Path $normalFile -Root $userRoot) {
        throw 'Failed: Test-PathContainsReparsePoint returned true for a normal path.'
    }

    # --- Guarded delete refuses junction ---
    $deleteJunction = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $junctionPath
    if ($deleteJunction.Result -ne 'Failed' -or $deleteJunction.Message -notmatch 'reparse point') {
        throw "Failed: delete did not reject junction. Result: $($deleteJunction.Result) $($deleteJunction.Message)"
    }

    # --- Guarded delete refuses file under junction ---
    $deleteUnderJunction = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $fileUnderJunction
    if ($deleteUnderJunction.Result -ne 'Failed' -or $deleteUnderJunction.Message -notmatch 'reparse point') {
        throw "Failed: delete did not reject file under junction. Result: $($deleteUnderJunction.Result) $($deleteUnderJunction.Message)"
    }

    # --- Guarded move refuses source under junction ---
    $moveTarget = Join-Path $userRoot 'moved.lnk'
    $moveUnderJunction = Invoke-GuardedFileOperation -Action 'Move' -SourcePath $fileUnderJunction -DestinationPath $moveTarget
    if ($moveUnderJunction.Result -ne 'Failed' -or $moveUnderJunction.Message -notmatch 'reparse point') {
        throw "Failed: move did not reject source under junction. Result: $($moveUnderJunction.Result) $($moveUnderJunction.Message)"
    }

    # --- Guarded move refuses destination under junction ---
    $normalSource = Join-Path $userRoot 'tomove.lnk'
    Set-Content -LiteralPath $normalSource -Value 'move-me'
    $destUnderJunction = Join-Path $junctionPath 'moved.lnk'
    $moveToJunction = Invoke-GuardedFileOperation -Action 'Move' -SourcePath $normalSource -DestinationPath $destUnderJunction
    if ($moveToJunction.Result -ne 'Failed' -or $moveToJunction.Message -notmatch 'reparse point') {
        throw "Failed: move did not reject destination under junction. Result: $($moveToJunction.Result) $($moveToJunction.Message)"
    }

    # --- Guarded delete on normal file still works ---
    $normalDelete = Join-Path $userRoot 'delete-me.lnk'
    Set-Content -LiteralPath $normalDelete -Value 'delete'
    $normalResult = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $normalDelete
    if ($normalResult.Result -ne 'Success') {
        throw "Failed: normal delete was rejected: $($normalResult.Message)"
    }
    if (Test-Path -LiteralPath $normalDelete) {
        throw 'Failed: file was not actually deleted.'
    }

    # --- Copy-DirectoryContents skips junctions ---
    $copySource = Join-Path $testRoot 'copysrc'
    $copyDest = Join-Path $testRoot 'copydst'
    [System.IO.Directory]::CreateDirectory($copySource) | Out-Null

    $realSubdir = Join-Path $copySource 'real'
    [System.IO.Directory]::CreateDirectory($realSubdir) | Out-Null
    Set-Content -LiteralPath (Join-Path $realSubdir 'file.txt') -Value 'real content'

    $junctionInCopy = Join-Path $copySource 'link-to-outside'
    cmd /c mklink /J "$junctionInCopy" "$junctionTarget" 2>&1 | Out-Null

    $copiedCount = Copy-DirectoryContents -SourcePath $copySource -DestinationPath $copyDest
    if ($copiedCount -ne 1) {
        throw "Failed: Copy-DirectoryContents copied $copiedCount items (expected 1, junction should be skipped)."
    }
    if (Test-Path -LiteralPath (Join-Path $copyDest 'link-to-outside')) {
        throw 'Failed: Copy-DirectoryContents copied the junction.'
    }
    if (-not (Test-Path -LiteralPath (Join-Path $copyDest 'real'))) {
        throw 'Failed: Copy-DirectoryContents did not copy the real subdirectory.'
    }

    # --- Junction target is untouched ---
    if (-not (Test-Path -LiteralPath (Join-Path $junctionTarget 'secret.txt'))) {
        throw 'Failed: junction target was modified by the tests.'
    }

    'SafetyGuards: All tests passed.'
}
finally {
    cmd /c rmdir /S /Q "$testRoot" 2>&1 | Out-Null
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
