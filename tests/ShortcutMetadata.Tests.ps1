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
    'Get-ShortcutMetadata',
    'Get-ShortcutRiskFlags',
    'Get-ShortcutTarget',
    'Test-ShortcutBroken'
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

$script:WScriptShell = New-Object -ComObject WScript.Shell

$testRoot = Join-Path $env:TEMP "StartMenuOrganizerMetadata_$([System.Guid]::NewGuid().ToString('N'))"
try {
    [System.IO.Directory]::CreateDirectory($testRoot) | Out-Null

    # --- .lnk metadata extraction ---
    $lnkPath = Join-Path $testRoot 'test-notepad.lnk'
    $lnk = $script:WScriptShell.CreateShortcut($lnkPath)
    $lnk.TargetPath = "$env:SystemRoot\notepad.exe"
    $lnk.Arguments = '/A testfile.txt'
    $lnk.WorkingDirectory = $env:TEMP
    $lnk.Description = 'Test shortcut'
    $lnk.Save()

    $meta = Get-ShortcutMetadata -ShortcutPath $lnkPath
    if ($meta.Target -ne "$env:SystemRoot\notepad.exe") {
        throw "Failed: .lnk Target expected '$env:SystemRoot\notepad.exe', got '$($meta.Target)'"
    }
    if ($meta.Arguments -ne '/A testfile.txt') {
        throw "Failed: .lnk Arguments expected '/A testfile.txt', got '$($meta.Arguments)'"
    }
    if ($meta.WorkingDir -ne $env:TEMP) {
        throw "Failed: .lnk WorkingDir expected '$env:TEMP', got '$($meta.WorkingDir)'"
    }
    if ($meta.Description -ne 'Test shortcut') {
        throw "Failed: .lnk Description expected 'Test shortcut', got '$($meta.Description)'"
    }

    # --- .url metadata extraction ---
    $urlPath = Join-Path $testRoot 'test-site.url'
    Set-Content -LiteralPath $urlPath -Value @(
        '[InternetShortcut]'
        'URL=https://example.com'
        'IconFile=C:\icon.ico'
        'HotKey=0'
    )

    $urlMeta = Get-ShortcutMetadata -ShortcutPath $urlPath
    if ($urlMeta.Target -ne 'https://example.com') {
        throw "Failed: .url Target expected 'https://example.com', got '$($urlMeta.Target)'"
    }
    if ($urlMeta.IconLocation -ne 'C:\icon.ico') {
        throw "Failed: .url IconLocation expected 'C:\icon.ico', got '$($urlMeta.IconLocation)'"
    }

    # --- Risk: ScriptHost detection ---
    $cmdMeta = [PSCustomObject]@{
        Target = 'C:\Windows\System32\cmd.exe'
        Arguments = '/c echo hello'
        WorkingDir = ''
        Description = ''
        Hotkey = ''
        IconLocation = ''
        WindowStyle = 0
    }
    $cmdRisk = @(Get-ShortcutRiskFlags -Metadata $cmdMeta)
    if ($cmdRisk -notcontains 'ScriptHost') {
        throw "Failed: cmd.exe should flag ScriptHost, got: $($cmdRisk -join ', ')"
    }

    # --- Risk: NetworkTarget detection ---
    $uncMeta = [PSCustomObject]@{
        Target = '\\server\share\app.exe'
        Arguments = ''
        WorkingDir = ''
        Description = ''
        Hotkey = ''
        IconLocation = ''
        WindowStyle = 0
    }
    $uncRisk = @(Get-ShortcutRiskFlags -Metadata $uncMeta)
    if ($uncRisk -notcontains 'NetworkTarget') {
        throw "Failed: UNC path should flag NetworkTarget, got: $($uncRisk -join ', ')"
    }

    # --- Risk: HiddenExecution detection ---
    $hiddenMeta = [PSCustomObject]@{
        Target = 'C:\Windows\System32\powershell.exe'
        Arguments = '-EncodedCommand ZQBjAGgAbwAgAGgAZQBsAGwAbwA='
        WorkingDir = ''
        Description = ''
        Hotkey = ''
        IconLocation = ''
        WindowStyle = 0
    }
    $hiddenRisk = @(Get-ShortcutRiskFlags -Metadata $hiddenMeta)
    if ($hiddenRisk -notcontains 'HiddenExecution') {
        throw "Failed: -EncodedCommand should flag HiddenExecution, got: $($hiddenRisk -join ', ')"
    }
    if ($hiddenRisk -notcontains 'ScriptHost') {
        throw "Failed: powershell.exe should also flag ScriptHost, got: $($hiddenRisk -join ', ')"
    }

    # --- Risk: LongArguments detection ---
    $longArgsMeta = [PSCustomObject]@{
        Target = 'C:\app.exe'
        Arguments = ('x' * 300)
        WorkingDir = ''
        Description = ''
        Hotkey = ''
        IconLocation = ''
        WindowStyle = 0
    }
    $longRisk = @(Get-ShortcutRiskFlags -Metadata $longArgsMeta)
    if ($longRisk -notcontains 'LongArguments') {
        throw "Failed: 300-char arguments should flag LongArguments, got: $($longRisk -join ', ')"
    }

    # --- Risk: WebTarget detection ---
    $webMeta = [PSCustomObject]@{
        Target = 'https://malicious.example.com/payload.exe'
        Arguments = ''
        WorkingDir = ''
        Description = ''
        Hotkey = ''
        IconLocation = ''
        WindowStyle = 0
    }
    $webRisk = @(Get-ShortcutRiskFlags -Metadata $webMeta)
    if ($webRisk -notcontains 'WebTarget') {
        throw "Failed: https URL should flag WebTarget, got: $($webRisk -join ', ')"
    }

    # --- Risk: ScriptTarget detection ---
    $scriptMeta = [PSCustomObject]@{
        Target = 'C:\scripts\setup.bat'
        Arguments = ''
        WorkingDir = ''
        Description = ''
        Hotkey = ''
        IconLocation = ''
        WindowStyle = 0
    }
    $scriptRisk = @(Get-ShortcutRiskFlags -Metadata $scriptMeta)
    if ($scriptRisk -notcontains 'ScriptTarget') {
        throw "Failed: .bat target should flag ScriptTarget, got: $($scriptRisk -join ', ')"
    }

    # --- No risk for normal shortcut ---
    $safeMeta = [PSCustomObject]@{
        Target = 'C:\Program Files\App\app.exe'
        Arguments = ''
        WorkingDir = ''
        Description = ''
        Hotkey = ''
        IconLocation = ''
        WindowStyle = 0
    }
    $safeRisk = @(Get-ShortcutRiskFlags -Metadata $safeMeta)
    if ($safeRisk.Count -gt 0) {
        throw "Failed: normal exe should have no risk flags, got: $($safeRisk -join ', ')"
    }

    'ShortcutMetadata: All tests passed.'
}
finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}
