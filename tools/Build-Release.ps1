#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputRoot
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $repoRoot 'dist'
}

$appScript = Join-Path $repoRoot 'StartMenuOrganizerPro.ps1'
$readme = Join-Path $repoRoot 'README.md'
$license = Join-Path $repoRoot 'LICENSE'

if (-not (Test-Path -LiteralPath $appScript)) {
    throw "Application script not found: $appScript"
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

if (Test-Path -LiteralPath $OutputRoot) {
    Remove-Item -LiteralPath $OutputRoot -Recurse -Force
}
[System.IO.Directory]::CreateDirectory($stageRoot) | Out-Null

Copy-Item -LiteralPath $appScript -Destination (Join-Path $stageRoot 'StartMenuOrganizerPro.ps1') -Force
Copy-Item -LiteralPath $readme -Destination (Join-Path $stageRoot 'README.md') -Force
Copy-Item -LiteralPath $license -Destination (Join-Path $stageRoot 'LICENSE') -Force

$installScript = @'
#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'Programs\Start Menu Organizer')
)

$ErrorActionPreference = 'Stop'
$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceScript = Join-Path $packageRoot 'StartMenuOrganizerPro.ps1'
if (-not (Test-Path -LiteralPath $sourceScript)) {
    throw "StartMenuOrganizerPro.ps1 was not found beside the installer."
}

[System.IO.Directory]::CreateDirectory($InstallRoot) | Out-Null
Copy-Item -LiteralPath $sourceScript -Destination (Join-Path $InstallRoot 'StartMenuOrganizerPro.ps1') -Force

$shortcutDir = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs\Start Menu Organizer'
[System.IO.Directory]::CreateDirectory($shortcutDir) | Out-Null
$shortcutPath = Join-Path $shortcutDir 'Start Menu Organizer.lnk'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$InstallRoot\StartMenuOrganizerPro.ps1`""
$shortcut.WorkingDirectory = $InstallRoot
$shortcut.Description = 'Start Menu Organizer'
$shortcut.Save()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null

$uninstallPath = Join-Path $InstallRoot 'Uninstall-StartMenuOrganizer.ps1'
Copy-Item -LiteralPath (Join-Path $packageRoot 'Uninstall-StartMenuOrganizer.ps1') -Destination $uninstallPath -Force

Write-Host "Installed Start Menu Organizer to $InstallRoot"
Write-Host "Shortcut created at $shortcutPath"
Write-Host "Uninstall with: powershell -NoProfile -ExecutionPolicy Bypass -File `"$uninstallPath`""
'@

$uninstallScript = @'
#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'Programs\Start Menu Organizer')
)

$ErrorActionPreference = 'Stop'
$shortcutDir = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs\Start Menu Organizer'
$shortcutPath = Join-Path $shortcutDir 'Start Menu Organizer.lnk'

if (Test-Path -LiteralPath $shortcutPath) {
    Remove-Item -LiteralPath $shortcutPath -Force
}
if ((Test-Path -LiteralPath $shortcutDir) -and -not (Get-ChildItem -LiteralPath $shortcutDir -Force -ErrorAction SilentlyContinue)) {
    Remove-Item -LiteralPath $shortcutDir -Force
}
if (Test-Path -LiteralPath $InstallRoot) {
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
foreach ($file in (Get-ChildItem -LiteralPath $stageRoot -File -Force | Sort-Object Name)) {
    $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
    $manifestLines += "$hash  $($file.Name)"
}
[System.IO.File]::WriteAllText($manifestPath, ($manifestLines -join "`n"), [System.Text.UTF8Encoding]::new($false))

$packageItems = Get-ChildItem -LiteralPath $stageRoot -Force
Compress-Archive -LiteralPath $packageItems.FullName -DestinationPath $artifactPath -Force
if (-not (Test-Path -LiteralPath $artifactPath)) {
    throw "Artifact was not created: $artifactPath"
}

$artifactHash = (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash
$artifactManifest = Join-Path $OutputRoot "$packageName.zip.sha256"
[System.IO.File]::WriteAllText($artifactManifest, "$artifactHash  $packageName.zip", [System.Text.UTF8Encoding]::new($false))

[PSCustomObject]@{
    Version = $version
    Artifact = $artifactPath
    ArtifactHash = $artifactHash
    ManifestPath = $manifestPath
    Signed = [bool]$signingCert
    PackageRoot = $stageRoot
}
