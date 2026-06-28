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
    'Get-ShortcutTarget',
    'Test-ShortcutBroken',
    'Test-ProtectedFolder'
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

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerShortcutValidation_$([System.Guid]::NewGuid().ToString('N'))"
try {
    $basePath = Join-Path $testRoot 'Programs'
    $startupPath = Join-Path $basePath 'Startup'
    $appsPath = Join-Path $basePath 'Apps'
    [System.IO.Directory]::CreateDirectory($startupPath) | Out-Null
    [System.IO.Directory]::CreateDirectory($appsPath) | Out-Null
    $script:ProtectedFolders = @('Startup')

    if (-not (Test-ProtectedFolder -Path (Join-Path $startupPath 'Keep.lnk') -BasePath $basePath)) {
        throw 'Failed: Startup child was not detected as protected.'
    }
    if (Test-ProtectedFolder -Path (Join-Path $appsPath 'Move.lnk') -BasePath $basePath) {
        throw 'Failed: non-protected folder was detected as protected.'
    }

    $webUrl = Join-Path $appsPath 'Web.url'
    Set-Content -LiteralPath $webUrl -Value @('[InternetShortcut]', 'URL=https://example.com')
    if ((Get-ShortcutTarget -ShortcutPath $webUrl) -ne 'https://example.com') {
        throw 'Failed: .url target was not parsed.'
    }
    if (Test-ShortcutBroken -ShortcutPath $webUrl) {
        throw 'Failed: https .url was marked broken.'
    }

    $missingFileUrl = Join-Path $appsPath 'MissingFile.url'
    Set-Content -LiteralPath $missingFileUrl -Value @('[InternetShortcut]', 'URL=file:///C:/definitely/missing/start-menu-target.txt')
    if (-not (Test-ShortcutBroken -ShortcutPath $missingFileUrl)) {
        throw 'Failed: missing file URL was not marked broken.'
    }

    $appRef = Join-Path $appsPath 'App.appref-ms'
    Set-Content -LiteralPath $appRef -Value 'appref'
    if ((Get-ShortcutTarget -ShortcutPath $appRef) -ne $appRef) {
        throw 'Failed: appref-ms target did not classify as itself.'
    }
    if (Test-ShortcutBroken -ShortcutPath $appRef) {
        throw 'Failed: existing appref-ms file was marked broken.'
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Shortcut validation tests passed.'
