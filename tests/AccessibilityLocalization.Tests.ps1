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

function Invoke-ScriptAssignment {
    param([string]$LeftHandSide)

    $assignmentAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq $LeftHandSide
    }, $true)

    if (-not $assignmentAst) {
        throw "Failed: assignment not found: $LeftHandSide"
    }

    Invoke-Expression $assignmentAst.Extent.Text
}

function Invoke-XamlAssignment {
    $assignmentAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -match '\$XAML$'
    }, $true)

    if (-not $assignmentAst) {
        throw 'Failed: XAML assignment not found.'
    }

    return [xml](Invoke-Expression $assignmentAst.Right.Extent.Text)
}

$helperFunctions = @(
    'Import-UiStringFile',
    'Load-UiStrings',
    'Get-UiString',
    'Convert-HexColorToRgb',
    'Convert-SrgbChannelToLinear',
    'Get-RelativeLuminance',
    'Get-ContrastRatio',
    'Test-ThemeContrast'
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

function Write-StructuredLogEntry {
    param(
        [string]$Message,
        [string]$Level = 'Info',
        [string]$OperationId = $null,
        [object]$Exception = $null,
        [string]$CrashLogPath = $null
    )
}

Invoke-ScriptAssignment -LeftHandSide '$script:DefaultUiStrings'
Invoke-ScriptAssignment -LeftHandSide '$script:UiTabOrder'
Invoke-ScriptAssignment -LeftHandSide '$script:ThemeContrastPairs'

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerA11y_$([System.Guid]::NewGuid().ToString('N'))"
try {
    [System.IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $script:Config = @{
        Version = 'test'
        LocalizationRoot = $testRoot
    }

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    $XAML = Invoke-XamlAssignment
    $reader = [System.Xml.XmlNodeReader]::new($XAML)
    $loadedWindow = [Windows.Markup.XamlReader]::Load($reader)
    foreach ($requiredControl in @('txtAppTitle', 'tabActions', 'cmbScopeUser', 'cmbScopeProfile', 'cmbScopeDefaultUser', 'txtCleanupHeader', 'btnDeleteSelected', 'txtProfileRoot', 'btnBrowseProfileRoot', 'txtLog')) {
        if (-not $loadedWindow.FindName($requiredControl)) {
            throw "Failed: XAML did not load named control $requiredControl"
        }
    }

    @{
        'Window.Title' = 'Localized Start Menu Organizer {0}'
        'btnBackup.Content' = 'Localized Backup'
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $testRoot 'strings.json')

    Load-UiStrings
    if ((Get-UiString -Key 'btnBackup.Content') -ne 'Localized Backup') {
        throw 'Failed: localization override was not loaded.'
    }
    if ((Get-UiString -Key 'Window.Title' -FormatArgs @('1.2.3')) -ne 'Localized Start Menu Organizer 1.2.3') {
        throw 'Failed: formatted localization string was not applied.'
    }

    $requiredStringKeys = @(
        'txtSearch.AutomationName',
        'txtSearch.HelpText',
        'btnBackup.AutomationName',
        'btnDeleteSelected.AutomationName',
        'dgItems.AutomationName',
        'txtLog.AutomationName',
        'txtProfileRoot.AutomationName',
        'btnBrowseProfileRoot.AutomationName',
        'btnUseDefaultProfileRoot.AutomationName',
        'Status.AdminFullAccess',
        'Status.StandardUser'
    )
    foreach ($key in $requiredStringKeys) {
        if (-not $script:DefaultUiStrings.Contains($key)) {
            throw "Failed: missing UI string key $key"
        }
    }

    $requiredTabStops = @('txtSearch', 'btnBackup', 'cmbScope', 'dgItems', 'btnDeleteSelected', 'txtFindText', 'lstJunkPatterns', 'txtProfileRoot', 'btnBrowseProfileRoot', 'btnClearLog')
    foreach ($tabStop in $requiredTabStops) {
        if ($script:UiTabOrder -notcontains $tabStop) {
            throw "Failed: missing tab stop $tabStop"
        }
    }
    if (($script:UiTabOrder | Select-Object -Unique).Count -ne $script:UiTabOrder.Count) {
        throw 'Failed: tab order contains duplicate controls.'
    }

    $contrastFailures = @(Test-ThemeContrast | Where-Object { -not $_.Pass })
    if ($contrastFailures.Count -gt 0) {
        $summary = $contrastFailures | ForEach-Object { "$($_.Name)=$($_.Ratio)" }
        throw "Failed: contrast checks did not pass: $($summary -join ', ')"
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Accessibility and localization tests passed.'
