#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'StartMenuOrganizerPro.ps1'
$buildPath = Join-Path $repoRoot 'tools\Build-Release.ps1'
$iconPngPath = Join-Path $repoRoot 'assets\brand\start-menu-organizer-1024.png'
$iconPath = Join-Path $repoRoot 'assets\brand\start-menu-organizer.ico'

$scriptText = Get-Content -LiteralPath $scriptPath -Raw
$buildText = Get-Content -LiteralPath $buildPath -Raw

if ($scriptText -notmatch '\[switch\]\$Demo' -or
    $scriptText -notmatch 'function Get-DemoStartMenuItems' -or
    $scriptText -notmatch 'No Start Menu files are being read or changed') {
    throw 'Failed: safe demo mode was not wired into the product.'
}

if ($scriptText -notmatch 'IsChecked="True" IsEnabled="False"' -or
    $scriptText -notmatch 'Every change is prepared as a reviewable plan') {
    throw 'Failed: the review-first workflow is not visible and locked on.'
}

foreach ($forbiddenPattern in @('"YesNo"', '<Window\.InputBindings>', '\.Add_KeyDown\(', 'Ctrl\+', 'ToolTip="F5"')) {
    if ($scriptText -match $forbiddenPattern) {
        throw "Failed: forbidden confirmation or keyboard-only UI remained: $forbiddenPattern"
    }
}

if ($buildText -match 'Get-FileHash' -or
    $buildText -match 'PackageIdentifier:\s*SysAdminDoc\.StartMenuOrganizer' -or
    $buildText -match 'ManifestType:\s*singleton') {
    throw 'Failed: release build still depends on Get-FileHash or authors winget metadata.'
}

if (-not (Test-Path -LiteralPath $iconPngPath -PathType Leaf) -or
    -not (Test-Path -LiteralPath $iconPath -PathType Leaf)) {
    throw 'Failed: the application icon family is incomplete.'
}

Add-Type -AssemblyName System.Drawing
$bitmap = [System.Drawing.Bitmap]::new($iconPngPath)
try {
    if ($bitmap.Width -ne 1024 -or $bitmap.Height -ne 1024) {
        throw "Failed: primary icon must be 1024x1024, got $($bitmap.Width)x$($bitmap.Height)."
    }
    if ($bitmap.GetPixel(0, 0).A -ne 0) {
        throw 'Failed: primary icon does not have a transparent outer canvas.'
    }
}
finally {
    $bitmap.Dispose()
}

'Marketing experience tests passed.'
