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
    'Remove-OldLogFiles',
    'Initialize-FileLogging',
    'Get-ExceptionDetails',
    'Write-StructuredLogEntry',
    'Write-CrashLog',
    'Write-Log',
    'Ensure-JournalStorage',
    'New-OperationId',
    'New-JournalItem',
    'Save-UndoJournal',
    'Add-JournalEntry'
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

Add-Type -AssemblyName PresentationFramework

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerLogging_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $logRoot = Join-Path $testRoot 'logs'
    $undoRoot = Join-Path $testRoot 'undo'
    [System.IO.Directory]::CreateDirectory($logRoot) | Out-Null

    1..5 | ForEach-Object {
        $oldLog = Join-Path $logRoot "StartMenuOrganizer-old$_.jsonl"
        Set-Content -LiteralPath $oldLog -Value "{}"
        (Get-Item -LiteralPath $oldLog).LastWriteTime = (Get-Date).AddDays(-$_)
    }

    $script:Config = @{
        Version = 'test'
        LogRoot = $logRoot
        MaxLogFiles = 3
        UndoFile = Join-Path $undoRoot 'undo.json'
        UndoBackupRoot = Join-Path $undoRoot 'UndoBackups'
        MaxUndoSteps = 50
    }
    $script:IsAdmin = $false
    $script:SessionId = 'test-session'
    $script:LogFilePath = $null
    $script:UndoStack = [System.Collections.Generic.List[PSObject]]::new()

    $Window = [PSCustomObject]@{ Dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher }
    $txtLog = [System.Windows.Controls.TextBlock]::new()
    $svLog = New-Object PSObject
    $svLog | Add-Member -MemberType ScriptMethod -Name ScrollToEnd -Value {}
    $btnUndo = [PSCustomObject]@{ IsEnabled = $false }

    $logPath = Initialize-FileLogging
    if (-not (Test-Path -LiteralPath $logPath)) {
        throw "Failed: log file was not initialized: $logPath"
    }
    $rotatedLogs = @(Get-ChildItem -LiteralPath $logRoot -Filter 'StartMenuOrganizer-*.jsonl')
    if ($rotatedLogs.Count -gt 3) {
        throw "Failed: log rotation kept $($rotatedLogs.Count) JSONL logs."
    }

    Write-Log 'Direct log message' 'Success' -OperationId 'op-direct'
    $directEntry = @(Get-Content -LiteralPath $logPath | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json })[-1]
    if ($directEntry.Message -ne 'Direct log message' -or
        $directEntry.Level -ne 'Success' -or
        $directEntry.OperationId -ne 'op-direct' -or
        $directEntry.SessionId -ne 'test-session') {
        throw 'Failed: direct structured log entry did not include expected fields.'
    }

    $journalItem = New-JournalItem -ActionType 'Delete' -OriginalPath 'C:\Temp\shortcut.lnk' -Result 'Success' -OperationId 'op-journal'
    Add-JournalEntry -Type 'Delete' -Description 'Delete shortcut' -Items @($journalItem) -OperationId 'op-journal'
    $journalEntry = @(Get-Content -LiteralPath $logPath | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json } |
        Where-Object { $_.Message -eq 'Journal recorded: Delete shortcut' })[-1]
    if ($journalEntry.OperationId -ne 'op-journal') {
        throw 'Failed: journal log entry did not include the operation ID.'
    }

    try {
        throw 'crash test'
    }
    catch {
        $crashPath = Write-CrashLog -Exception $_ -Context 'Test crash' -OperationId 'op-crash'
    }

    if (-not (Test-Path -LiteralPath $crashPath)) {
        throw 'Failed: crash log file was not created.'
    }
    $crashText = Get-Content -LiteralPath $crashPath -Raw
    if ($crashText -notmatch 'Test crash' -or $crashText -notmatch 'op-crash' -or $crashText -notmatch 'crash test') {
        throw 'Failed: crash log did not include context, operation ID, and exception message.'
    }
    $crashEntry = @(Get-Content -LiteralPath $logPath | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json } |
        Where-Object { $_.OperationId -eq 'op-crash' })[-1]
    if (-not $crashEntry.CrashLogPath -or $crashEntry.Level -ne 'Error') {
        throw 'Failed: structured crash entry did not include crash path and error level.'
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Logging tests passed.'
