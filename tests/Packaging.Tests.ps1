#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$buildScript = Join-Path $repoRoot 'tools\Build-Release.ps1'
$outputRoot = Join-Path $env:TEMP "StartMenuOrganizerPackage_$([System.Guid]::NewGuid().ToString('N'))"
$defaultOutputRoot = Join-Path $repoRoot 'dist'

try {
    $result = & $buildScript -OutputRoot $outputRoot

    if ($result.Version -ne '0.15.0') {
        throw "Failed: package version was $($result.Version)."
    }
    if (-not (Test-Path -LiteralPath $result.Artifact)) {
        throw "Failed: package artifact was not created: $($result.Artifact)"
    }
    if ((Split-Path $result.Artifact -Leaf) -ne 'StartMenuOrganizer-v0.15.0.zip') {
        throw "Failed: artifact filename was incorrect: $($result.Artifact)"
    }
    if ((Split-Path (Split-Path $result.Artifact -Parent) -Leaf) -ne (Split-Path $outputRoot -Leaf)) {
        throw "Failed: explicit output root was not honored: $($result.Artifact)"
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($result.Artifact)
    try {
        $entries = @($zip.Entries | ForEach-Object { $_.FullName })
    }
    finally {
        $zip.Dispose()
    }

    $expectedEntries = @(
        'StartMenuOrganizerPro.ps1',
        'Install-StartMenuOrganizer.ps1',
        'Uninstall-StartMenuOrganizer.ps1',
        'README.md',
        'LICENSE',
        'assets/brand/start-menu-organizer.ico',
        'assets/brand/start-menu-organizer-512.png',
        'assets/screenshots/01-start-menu-workspace.png',
        'assets/social-preview.png'
    )
    foreach ($entry in $expectedEntries) {
        if ($entries -notcontains $entry) {
            throw "Failed: package was missing $entry. Entries: $($entries -join ', ')"
        }
    }

    $installScript = Get-Content -LiteralPath (Join-Path $result.PackageRoot 'Install-StartMenuOrganizer.ps1') -Raw
    if ($installScript -notmatch 'CreateShortcut' -or
        $installScript -notmatch 'ExecutionPolicy Bypass' -or
        $installScript -notmatch 'Uninstall-StartMenuOrganizer\.ps1' -or
        $installScript -notmatch 'SkipShortcut' -or
        $installScript -notmatch 'start-menu-organizer\.ico' -or
        $installScript -notmatch '\.start-menu-organizer-install') {
        throw 'Failed: installer did not contain shortcut, icon, marker, and uninstall wiring.'
    }

    $uninstallScript = Get-Content -LiteralPath (Join-Path $result.PackageRoot 'Uninstall-StartMenuOrganizer.ps1') -Raw
    if ($uninstallScript -notmatch 'Start Menu Organizer\.lnk' -or
        $uninstallScript -notmatch 'Remove-Item' -or
        $uninstallScript -notmatch 'Refusing to remove an unrecognized install directory') {
        throw 'Failed: uninstaller did not validate and remove the shortcut and install root.'
    }

    $metadataFiles = @(Get-ChildItem -LiteralPath $result.PackageMetadata -File)
    if ($metadataFiles.Name -contains 'SysAdminDoc.StartMenuOrganizer.yaml') {
        throw 'Failed: the build authored a prohibited winget manifest.'
    }

    $installRoot = Join-Path $outputRoot 'install-validation'
    $installerPath = Join-Path $result.PackageRoot 'Install-StartMenuOrganizer.ps1'
    & $installerPath -InstallRoot $installRoot -SkipShortcut
    & $installerPath -InstallRoot $installRoot -SkipShortcut
    if (-not (Test-Path -LiteralPath (Join-Path $installRoot 'assets\brand\start-menu-organizer.ico')) -or
        -not (Test-Path -LiteralPath (Join-Path $installRoot '.start-menu-organizer-install')) -or
        (Test-Path -LiteralPath (Join-Path $installRoot 'assets\assets'))) {
        throw 'Failed: repeated isolated install did not preserve the expected branded layout.'
    }

    & (Join-Path $installRoot 'Uninstall-StartMenuOrganizer.ps1') -InstallRoot $installRoot
    if (Test-Path -LiteralPath $installRoot) {
        throw 'Failed: isolated uninstall did not remove the validated install directory.'
    }

    $defaultResult = & $buildScript
    if (-not (Test-Path -LiteralPath $defaultResult.Artifact)) {
        throw "Failed: default package artifact was not created: $($defaultResult.Artifact)"
    }
    if ((Split-Path (Split-Path $defaultResult.Artifact -Parent) -Leaf) -ne 'dist') {
        throw "Failed: default output root was not dist: $($defaultResult.Artifact)"
    }
}
finally {
    if (Test-Path -LiteralPath $outputRoot) {
        Remove-Item -LiteralPath $outputRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $defaultOutputRoot) {
        Remove-Item -LiteralPath $defaultOutputRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

'Packaging tests passed.'
