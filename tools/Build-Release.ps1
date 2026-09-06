#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputRoot
)

$ErrorActionPreference = 'Stop'

function Get-Sha256Hash {
    param([Parameter(Mandatory)][string]$Path)

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        try {
            $bytes = $sha256.ComputeHash($stream)
            return ([System.BitConverter]::ToString($bytes)).Replace('-', '')
        }
        finally {
            $sha256.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $repoRoot 'dist'
}
$OutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
if ($OutputRoot -eq [System.IO.Path]::GetPathRoot($OutputRoot)) {
    throw 'OutputRoot cannot be a drive root.'
}

$appScript = Join-Path $repoRoot 'StartMenuOrganizerPro.ps1'
$readme = Join-Path $repoRoot 'README.md'
$license = Join-Path $repoRoot 'LICENSE'
$assetsRoot = Join-Path $repoRoot 'assets'
$brandRoot = Join-Path $assetsRoot 'brand'
$icon = Join-Path $brandRoot 'start-menu-organizer.ico'

if (-not (Test-Path -LiteralPath $appScript)) {
    throw "Application script not found: $appScript"
}
if (-not (Test-Path -LiteralPath $icon)) {
    throw "Application icon not found: $icon"
}

$scriptText = Get-Content -LiteralPath $appScript -Raw
$versionMatch = [regex]::Match($scriptText, 'Version\s*=\s*"(?<version>[^"]+)"')
if (-not $versionMatch.Success) {
    throw 'Could not read application version from StartMenuOrganizerPro.ps1.'
}

$version = $versionMatch.Groups['version'].Value
$packageName = "StartMenuOrganizer-v$version"
$stageRoot = Join-Path $OutputRoot $packageName
$artifactPath = Join-Path $OutputRoot "$packageName.zip"
$artifactManifest = Join-Path $OutputRoot "$packageName.zip.sha256"
$pkgMetaDir = Join-Path $OutputRoot 'package-metadata'

[System.IO.Directory]::CreateDirectory($OutputRoot) | Out-Null
$staleArtifacts = @(Get-ChildItem -LiteralPath $OutputRoot -Force | Where-Object {
    $_.Name -match '^StartMenuOrganizer-v\d+\.\d+\.\d+(\.zip(\.sha256)?)?$'
})
foreach ($staleArtifact in $staleArtifacts) {
    if ([System.IO.Path]::GetFullPath((Split-Path -Parent $staleArtifact.FullName)) -ne $OutputRoot) {
        throw "Stale artifact escaped OutputRoot: $($staleArtifact.FullName)"
    }
    Remove-Item -LiteralPath $staleArtifact.FullName -Recurse -Force
}
foreach ($generatedPath in @($stageRoot, $artifactPath, $artifactManifest, $pkgMetaDir)) {
    $resolvedGeneratedPath = [System.IO.Path]::GetFullPath($generatedPath)
    if (-not $resolvedGeneratedPath.StartsWith($OutputRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Generated path escaped OutputRoot: $resolvedGeneratedPath"
    }
    if (Test-Path -LiteralPath $resolvedGeneratedPath) {
        Remove-Item -LiteralPath $resolvedGeneratedPath -Recurse -Force
    }
}
[System.IO.Directory]::CreateDirectory($stageRoot) | Out-Null

Copy-Item -LiteralPath $appScript -Destination (Join-Path $stageRoot 'StartMenuOrganizerPro.ps1') -Force
Copy-Item -LiteralPath $readme -Destination (Join-Path $stageRoot 'README.md') -Force
Copy-Item -LiteralPath $license -Destination (Join-Path $stageRoot 'LICENSE') -Force
Copy-Item -LiteralPath $assetsRoot -Destination (Join-Path $stageRoot 'assets') -Recurse -Force

$installScript = @'
#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'Programs\Start Menu Organizer'),
    [switch]$SkipShortcut
)

$ErrorActionPreference = 'Stop'
$InstallRoot = [System.IO.Path]::GetFullPath($InstallRoot)
if ($InstallRoot -eq [System.IO.Path]::GetPathRoot($InstallRoot)) {
    throw 'InstallRoot cannot be a drive root.'
}
$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceScript = Join-Path $packageRoot 'StartMenuOrganizerPro.ps1'
$sourceAssets = Join-Path $packageRoot 'assets'
if (-not (Test-Path -LiteralPath $sourceScript)) {
    throw "StartMenuOrganizerPro.ps1 was not found beside the installer."
}
if (-not (Test-Path -LiteralPath $sourceAssets -PathType Container)) {
    throw "Brand assets were not found beside the installer."
}

[System.IO.Directory]::CreateDirectory($InstallRoot) | Out-Null
Copy-Item -LiteralPath $sourceScript -Destination (Join-Path $InstallRoot 'StartMenuOrganizerPro.ps1') -Force
$installedAssets = Join-Path $InstallRoot 'assets'
if (Test-Path -LiteralPath $installedAssets) {
    Remove-Item -LiteralPath $installedAssets -Recurse -Force
}
Copy-Item -LiteralPath $sourceAssets -Destination $installedAssets -Recurse -Force
[System.IO.File]::WriteAllText((Join-Path $InstallRoot '.start-menu-organizer-install'), 'Start Menu Organizer installation marker', [System.Text.UTF8Encoding]::new($false))

$shortcutDir = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs\Start Menu Organizer'
$shortcutPath = Join-Path $shortcutDir 'Start Menu Organizer.lnk'
if (-not $SkipShortcut) {
    [System.IO.Directory]::CreateDirectory($shortcutDir) | Out-Null
    $shell = New-Object -ComObject WScript.Shell
    try {
        $shortcut = $shell.CreateShortcut($shortcutPath)
        try {
            $shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
            $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$InstallRoot\StartMenuOrganizerPro.ps1`""
            $shortcut.WorkingDirectory = $InstallRoot
            $shortcut.Description = 'Start Menu Organizer'
            $shortcut.IconLocation = Join-Path $InstallRoot 'assets\brand\start-menu-organizer.ico'
            $shortcut.Save()
        }
        finally {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
        }
    }
    finally {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
}

$uninstallPath = Join-Path $InstallRoot 'Uninstall-StartMenuOrganizer.ps1'
Copy-Item -LiteralPath (Join-Path $packageRoot 'Uninstall-StartMenuOrganizer.ps1') -Destination $uninstallPath -Force

Write-Host "Installed Start Menu Organizer to $InstallRoot"
if ($SkipShortcut) {
    Write-Host 'Start Menu shortcut skipped.'
}
else {
    Write-Host "Shortcut created at $shortcutPath"
}
Write-Host "Uninstall with: powershell -NoProfile -ExecutionPolicy Bypass -File `"$uninstallPath`""
'@

$uninstallScript = @'
#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'Programs\Start Menu Organizer')
)

$ErrorActionPreference = 'Stop'
$InstallRoot = [System.IO.Path]::GetFullPath($InstallRoot)
if ($InstallRoot -eq [System.IO.Path]::GetPathRoot($InstallRoot)) {
    throw 'InstallRoot cannot be a drive root.'
}
$installMarker = Join-Path $InstallRoot '.start-menu-organizer-install'
$shortcutDir = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs\Start Menu Organizer'
$shortcutPath = Join-Path $shortcutDir 'Start Menu Organizer.lnk'

if (Test-Path -LiteralPath $shortcutPath) {
    Remove-Item -LiteralPath $shortcutPath -Force
}
if ((Test-Path -LiteralPath $shortcutDir) -and -not (Get-ChildItem -LiteralPath $shortcutDir -Force -ErrorAction SilentlyContinue)) {
    Remove-Item -LiteralPath $shortcutDir -Force
}
if (Test-Path -LiteralPath $InstallRoot) {
    if (-not (Test-Path -LiteralPath $installMarker -PathType Leaf)) {
        throw "Refusing to remove an unrecognized install directory: $InstallRoot"
    }
    Remove-Item -LiteralPath $InstallRoot -Recurse -Force
}

Write-Host 'Start Menu Organizer was uninstalled.'
'@

[System.IO.File]::WriteAllText((Join-Path $stageRoot 'Install-StartMenuOrganizer.ps1'), $installScript, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $stageRoot 'Uninstall-StartMenuOrganizer.ps1'), $uninstallScript, [System.Text.UTF8Encoding]::new($false))

$signingCert = $null
try {
    $signingCert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction Stop | Select-Object -First 1
}
catch {
    Write-Verbose "No code signing certificate found."
}

$ps1Files = Get-ChildItem -LiteralPath $stageRoot -Filter '*.ps1' -Force
if ($signingCert) {
    foreach ($ps1 in $ps1Files) {
        Set-AuthenticodeSignature -FilePath $ps1.FullName -Certificate $signingCert -TimestampServer 'http://timestamp.digicert.com' -ErrorAction Stop | Out-Null
    }
    Write-Host "Scripts signed with: $($signingCert.Subject)"
}
else {
    Write-Host 'No code signing certificate available. Scripts are unsigned.'
}

$manifestPath = Join-Path $stageRoot 'SHA256SUMS.txt'
$manifestLines = @()
foreach ($file in (Get-ChildItem -LiteralPath $stageRoot -File -Recurse -Force | Sort-Object FullName)) {
    $hash = Get-Sha256Hash -Path $file.FullName
    $relativePath = $file.FullName.Substring($stageRoot.Length + 1).Replace('\', '/')
    $manifestLines += "$hash  $relativePath"
}
[System.IO.File]::WriteAllText($manifestPath, ($manifestLines -join "`n"), [System.Text.UTF8Encoding]::new($false))

$packageItems = Get-ChildItem -LiteralPath $stageRoot -Force
Compress-Archive -LiteralPath $packageItems.FullName -DestinationPath $artifactPath -Force
if (-not (Test-Path -LiteralPath $artifactPath)) {
    throw "Artifact was not created: $artifactPath"
}

$artifactHash = Get-Sha256Hash -Path $artifactPath
[System.IO.File]::WriteAllText($artifactManifest, "$artifactHash  $packageName.zip", [System.Text.UTF8Encoding]::new($false))

[System.IO.Directory]::CreateDirectory($pkgMetaDir) | Out-Null

$scoopManifest = [ordered]@{
    version = $version
    description = 'Start Menu cleanup and organization tool for Windows'
    homepage = 'https://github.com/SysAdminDoc/Start-Menu-Organizer'
    license = 'MIT'
    url = "https://github.com/SysAdminDoc/Start-Menu-Organizer/releases/download/v$version/$packageName.zip"
    hash = $artifactHash
    installer = @{
        script = 'powershell -NoProfile -ExecutionPolicy Bypass -File "$dir\Install-StartMenuOrganizer.ps1"'
    }
    uninstaller = @{
        script = 'powershell -NoProfile -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Programs\Start Menu Organizer\Uninstall-StartMenuOrganizer.ps1"'
    }
}
$scoopJson = $scoopManifest | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText((Join-Path $pkgMetaDir 'start-menu-organizer.json'), $scoopJson, [System.Text.UTF8Encoding]::new($false))

$chocoNuspec = @"
<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://schemas.microsoft.com/packaging/2015/06/nuspec.xsd">
  <metadata>
    <id>start-menu-organizer</id>
    <version>$version</version>
    <title>Start Menu Organizer</title>
    <authors>SysAdminDoc</authors>
    <projectUrl>https://github.com/SysAdminDoc/Start-Menu-Organizer</projectUrl>
    <licenseUrl>https://github.com/SysAdminDoc/Start-Menu-Organizer/blob/main/LICENSE</licenseUrl>
    <requireLicenseAcceptance>false</requireLicenseAcceptance>
    <description>Clean up junk, detect broken shortcuts, remove duplicates, and organize your Windows Start Menu.</description>
    <summary>Windows Start Menu management tool</summary>
    <tags>start-menu windows cleanup organizer shortcuts</tags>
  </metadata>
</package>
"@
[System.IO.File]::WriteAllText((Join-Path $pkgMetaDir 'start-menu-organizer.nuspec'), $chocoNuspec, [System.Text.UTF8Encoding]::new($false))

Write-Host "Package metadata generated in $pkgMetaDir"

[PSCustomObject]@{
    Version = $version
    Artifact = $artifactPath
    ArtifactHash = $artifactHash
    ManifestPath = $manifestPath
    PackageMetadata = $pkgMetaDir
    Signed = [bool]$signingCert
    PackageRoot = $stageRoot
}
