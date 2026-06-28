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
    'Get-ApprovedMutationRoot',
    'Ensure-JournalStorage',
    'New-OperationId',
    'New-UndoBackupCopy',
    'Invoke-GuardedFileOperation',
    'New-JournalItem',
    'Save-UndoJournal',
    'Load-UndoJournal',
    'Add-JournalEntry',
    'Invoke-JournaledDelete',
    'Restore-JournalBackup',
    'Invoke-Undo'
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

function Refresh-Items {}

$dispatcher = New-Object PSObject
$dispatcher | Add-Member -MemberType ScriptMethod -Name Invoke -Value {
    param($action)
    $action.Invoke()
}
$Window = [PSCustomObject]@{ Dispatcher = $dispatcher }
$btnUndo = [PSCustomObject]@{ IsEnabled = $false }
$script:UndoStack = [System.Collections.Generic.List[PSObject]]::new()

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerUndoJournal_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $liveRoot = Join-Path $testRoot 'live'
    $undoRoot = Join-Path $testRoot 'undo'
    $sourcePath = Join-Path $liveRoot 'delete-me.lnk'

    [System.IO.Directory]::CreateDirectory($liveRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($undoRoot) | Out-Null
    Set-Content -LiteralPath $sourcePath -Value 'shortcut'

    $script:Config = @{
        UndoFile = Join-Path $undoRoot 'undo.json'
        UndoBackupRoot = Join-Path $undoRoot 'UndoBackups'
        MaxUndoSteps = 50
        UserStartMenu = $liveRoot
        SystemStartMenu = Join-Path $testRoot 'system-start'
    }
    $script:ProtectedFolders = @('Startup')

    $operationId = New-OperationId
    $record = Invoke-JournaledDelete -Path $sourcePath -OperationId $operationId
    Add-JournalEntry -Type 'Delete' -Description 'Delete one item' -Items @($record) -OperationId $operationId

    if (Test-Path -LiteralPath $sourcePath) {
        throw 'Failed: source file still exists after journaled delete.'
    }

    $journal = @(Get-Content -LiteralPath $script:Config.UndoFile -Raw | ConvertFrom-Json)
    if ($journal.Count -ne 1) {
        throw "Failed: expected one journal entry, found $($journal.Count)."
    }
    $journalItem = @($journal[0].Items)[0]
    foreach ($propertyName in @('OriginalPath', 'NewPath', 'BackupPath', 'ActionType', 'Timestamp', 'Result')) {
        if ($journalItem.PSObject.Properties.Name -notcontains $propertyName) {
            throw "Failed: journal item missing $propertyName."
        }
    }
    if (-not (Test-Path -LiteralPath $journalItem.BackupPath)) {
        throw 'Failed: durable undo backup was not created.'
    }

    $script:UndoStack.Clear()
    $btnUndo.IsEnabled = $false
    Load-UndoJournal

    if ($script:UndoStack.Count -ne 1) {
        throw "Failed: expected one loaded undo entry, found $($script:UndoStack.Count)."
    }
    if (-not $btnUndo.IsEnabled) {
        throw 'Failed: undo button was not enabled after loading journal.'
    }

    Invoke-Undo

    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw 'Failed: undo did not restore the deleted file.'
    }
    if ($script:UndoStack.Count -ne 0) {
        throw 'Failed: undo stack was not cleared after undo.'
    }

    $remainingJournal = Get-Content -LiteralPath $script:Config.UndoFile -Raw
    if ($remainingJournal.Trim() -ne '[]') {
        throw "Failed: undo journal was not emptied after undo. Value: $remainingJournal"
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Undo journal tests passed.'
