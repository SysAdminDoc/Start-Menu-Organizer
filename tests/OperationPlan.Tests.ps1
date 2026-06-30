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
    'Ensure-JournalStorage',
    'New-OperationId',
    'New-UndoBackupCopy',
    'Invoke-GuardedFileOperation',
    'New-JournalItem',
    'Save-UndoJournal',
    'Add-JournalEntry',
    'Invoke-JournaledDelete',
    'Invoke-JournaledMove',
    'Invoke-JournaledRename',
    'New-PlanOperation',
    'Set-CurrentOperationPlan',
    'Update-PlanStatus',
    'Show-OperationPlanSummary',
    'Clear-OperationPlan',
    'Invoke-OperationPlanSynchronously',
    'Start-OperationPlanWorker',
    'Invoke-CurrentOperationPlan'
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

function Show-Progress {
    param(
        [int]$Value = 0,
        [int]$Maximum = 100,
        [bool]$Visible = $true
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
$txtPlanStatus = [PSCustomObject]@{ Text = '' }
$btnExportPlan = [PSCustomObject]@{ IsEnabled = $false }
$btnExecutePlan = [PSCustomObject]@{ IsEnabled = $false }
$btnClearPlan = [PSCustomObject]@{ IsEnabled = $false }
$script:UndoStack = [System.Collections.Generic.List[PSObject]]::new()
$script:CurrentOperationPlan = $null
$script:UseSynchronousPlanExecution = $true

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerPlan_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $userRoot = Join-Path $testRoot 'user'
    $systemRoot = Join-Path $testRoot 'system'
    $undoRoot = Join-Path $testRoot 'undo'
    [System.IO.Directory]::CreateDirectory($userRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($systemRoot) | Out-Null
    [System.IO.Directory]::CreateDirectory($undoRoot) | Out-Null

    $script:Config = @{
        Version = 'test'
        UserStartMenu = $userRoot
        SystemStartMenu = $systemRoot
        UndoFile = Join-Path $undoRoot 'undo.json'
        UndoBackupRoot = Join-Path $undoRoot 'UndoBackups'
        MaxUndoSteps = 50
    }
    $script:ProtectedFolders = @('Startup')

    $deletePath = Join-Path $userRoot 'delete.lnk'
    $moveSource = Join-Path $userRoot 'move.lnk'
    $moveDestination = Join-Path $userRoot 'Category\move.lnk'
    $renameSource = Join-Path $userRoot 'rename-old.lnk'
    $renameNewName = 'rename-new.lnk'
    $renameDestination = Join-Path $userRoot $renameNewName

    Set-Content -LiteralPath $deletePath -Value 'delete'
    Set-Content -LiteralPath $moveSource -Value 'move'
    Set-Content -LiteralPath $renameSource -Value 'rename'

    $operations = @(
        (New-PlanOperation -Action 'Delete' -SourcePath $deletePath -DisplayName 'delete'),
        (New-PlanOperation -Action 'Move' -SourcePath $moveSource -DestinationPath $moveDestination -DisplayName 'move'),
        (New-PlanOperation -Action 'Rename' -SourcePath $renameSource -NewName $renameNewName -DisplayName 'rename')
    )

    Set-CurrentOperationPlan -Name 'Mixed temp plan' -Operations $operations
    if ($txtPlanStatus.Text -notmatch '3 operation') {
        throw "Failed: plan status did not show operation count: $($txtPlanStatus.Text)"
    }
    if (-not $btnExecutePlan.IsEnabled -or -not $btnExportPlan.IsEnabled) {
        throw 'Failed: plan controls were not enabled.'
    }

    $roundTrip = $script:CurrentOperationPlan | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    if (@($roundTrip.Operations).Count -ne 3) {
        throw 'Failed: operation plan JSON round-trip lost operations.'
    }

    Invoke-CurrentOperationPlan

    if (Test-Path -LiteralPath $deletePath) {
        throw 'Failed: plan delete did not remove the source.'
    }
    if (-not (Test-Path -LiteralPath $moveDestination)) {
        throw 'Failed: plan move did not create the reviewed destination.'
    }
    if (-not (Test-Path -LiteralPath $renameDestination)) {
        throw 'Failed: plan rename did not create the reviewed name.'
    }
    if ($script:CurrentOperationPlan) {
        throw 'Failed: current plan was not cleared after execution.'
    }
    if (-not (Test-Path -LiteralPath $script:Config.UndoFile)) {
        throw 'Failed: plan execution did not write the undo journal.'
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Operation plan tests passed.'
