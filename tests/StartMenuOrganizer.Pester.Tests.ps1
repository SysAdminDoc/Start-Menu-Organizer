BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:ScriptPath = Join-Path $script:RepoRoot 'StartMenuOrganizerPro.ps1'
}

Describe 'Start Menu Organizer static gates' {
    It 'parses as PowerShell' {
        $tokens = $null
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$parseErrors) | Out-Null
        $parseErrors.Count | Should -Be 0
    }

    It 'keeps the app version aligned with README' {
        $scriptText = Get-Content -LiteralPath $script:ScriptPath -Raw
        $readmeText = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'README.md') -Raw
        $scriptText | Should -Match 'Version\s+=\s+"0\.14\.0"'
        $readmeText | Should -Match 'Version-v0\.14\.0'
    }
}
