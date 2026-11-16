[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Pester shares variables between BeforeAll/BeforeEach and It blocks.')]
param()

BeforeAll {
    $repoRoot = Split-Path -Path $PSScriptRoot -Parent

    function Add-TestFile {
        param([string]$Path, [string]$Content = 'x')
        $null = New-Item -ItemType File -Path $Path -Value $Content -Force
    }
}

Describe 'Script hygiene' {
    It '<_> parses without errors' -ForEach @('Common.ps1', 'MovingFiles.ps1', 'DirectoryReorder.ps1', 'DirectoryDistributionChecker.ps1') {
        $tokens = $null
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path -Path $repoRoot -ChildPath $_), [ref]$tokens, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It '<_> dot-sources Common.ps1 and avoids Write-Host, exit and hardcoded drive defaults' -ForEach @('MovingFiles.ps1', 'DirectoryReorder.ps1', 'DirectoryDistributionChecker.ps1') {
        $content = Get-Content -LiteralPath (Join-Path -Path $repoRoot -ChildPath $_) -Raw
        $content | Should -Match "\.\s+\(Join-Path \`$PSScriptRoot 'Common\.ps1'\)"
        $content | Should -Not -Match 'Write-Host'
        $content | Should -Not -Match '(?m)^\s*exit\b'
        $content | Should -Not -Match 'F:\\|YourName'
    }

    It 'MovingFiles.ps1 no longer exposes -Parallel' {
        (Get-Command (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1')).Parameters.Keys | Should -Not -Contain 'Parallel'
    }

    It 'DirectoryReorder.ps1 has no duplicate document extensions' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path -Path $repoRoot -ChildPath 'DirectoryReorder.ps1'), [ref]$null, [ref]$null)
        $assignment = $ast.Find({
                param($node)
                $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $node.Left.Extent.Text -eq '$DocumentExtensions'
            }, $true)
        $assignment | Should -Not -BeNullOrEmpty
        $extensions = @($assignment.Right.FindAll({ param($node) $node -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true) | ForEach-Object Value)
        $extensions | Should -Contain '.tex'
        @($extensions | Group-Object | Where-Object Count -gt 1) | Should -BeNullOrEmpty
    }
}

Describe 'MovingFiles.ps1' {
    BeforeEach {
        $root = Join-Path -Path $TestDrive -ChildPath ([guid]::NewGuid().ToString('N'))
        $params = @{
            SourceDir          = Join-Path -Path $root -ChildPath 'Downloads'
            PersonalDir        = Join-Path -Path $root -ChildPath 'Personal'
            MusicDir           = Join-Path -Path $root -ChildPath 'Music'
            PhotosDir          = Join-Path -Path $root -ChildPath 'Photos'
            LogPath            = Join-Path -Path $root -ChildPath 'MovingFiles.log'
            ReportPath         = Join-Path -Path $root -ChildPath 'MovingFiles_Report_test.json'
            InformationAction  = 'SilentlyContinue'
            WarningAction      = 'SilentlyContinue'
        }
        Add-TestFile (Join-Path -Path $params.SourceDir -ChildPath 'report.pdf')
        Add-TestFile (Join-Path -Path $params.SourceDir -ChildPath 'song.mp3')
        Add-TestFile (Join-Path -Path $params.SourceDir -ChildPath 'photo.jpg')
        Add-TestFile ([System.IO.Path]::Combine($params.SourceDir, 'nested', 'deep.jpg'))
        Add-TestFile (Join-Path -Path $params.SourceDir -ChildPath 'archive.zip')
    }

    It 'throws when required paths are missing' {
        { & (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1') -LogPath (Join-Path -Path $root -ChildPath 'x.log') -InformationAction SilentlyContinue } |
            Should -Throw "*'SourceDir' is required*"
    }

    It 'moves files by category and records them in the report' {
        & (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1') @params

        Test-Path (Join-Path -Path $params.PersonalDir -ChildPath 'report.pdf') | Should -BeTrue
        Test-Path (Join-Path -Path $params.MusicDir -ChildPath 'song.mp3') | Should -BeTrue
        Test-Path (Join-Path -Path $params.PhotosDir -ChildPath 'photo.jpg') | Should -BeTrue
        Test-Path (Join-Path -Path $params.SourceDir -ChildPath 'archive.zip') | Should -BeTrue
        Test-Path ([System.IO.Path]::Combine($params.SourceDir, 'nested', 'deep.jpg')) | Should -BeTrue

        $report = Get-Content -LiteralPath $params.ReportPath -Raw | ConvertFrom-Json
        $report.Summary.TotalFilesMoved | Should -Be 3
        $report.Summary.TotalFilesFailed | Should -Be 0
        $report.Sections.Photo.Moved | Should -Be 1
    }

    It 'recurses and keeps relative folders with -PreserveFolderStructure' {
        & (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1') @params -PreserveFolderStructure

        Test-Path ([System.IO.Path]::Combine($params.PhotosDir, 'nested', 'deep.jpg')) | Should -BeTrue
        $report = Get-Content -LiteralPath $params.ReportPath -Raw | ConvertFrom-Json
        $report.Summary.TotalFilesMoved | Should -Be 4
        $report.Sections.Photo.Moved | Should -Be 2
    }

    It 'resolves relative paths against the current location' {
        Push-Location -LiteralPath $root
        try {
            & (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1') -SourceDir 'Downloads' -PersonalDir 'Personal' -MusicDir 'Music' `
                -PhotosDir 'Photos' -PreserveFolderStructure -LogPath $params.LogPath -ReportPath $params.ReportPath `
                -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        }
        finally {
            Pop-Location
        }
        Test-Path ([System.IO.Path]::Combine($params.PhotosDir, 'nested', 'deep.jpg')) | Should -BeTrue
        Test-Path (Join-Path -Path $params.PersonalDir -ChildPath 'report.pdf') | Should -BeTrue
    }

    It 'reads paths from a config file and keeps the old -FotosDir alias' {
        $configPath = Join-Path -Path $root -ChildPath 'config.json'
        @{ MovingFiles = @{ SourceDir = $params.SourceDir; PersonalDir = $params.PersonalDir; MusicDir = $params.MusicDir; PhotosDir = 'unused' } } |
            ConvertTo-Json | Set-Content -LiteralPath $configPath

        & (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1') -ConfigPath $configPath -FotosDir $params.PhotosDir -LogPath $params.LogPath `
            -ReportPath $params.ReportPath -InformationAction SilentlyContinue -WarningAction SilentlyContinue

        Test-Path (Join-Path -Path $params.PhotosDir -ChildPath 'photo.jpg') | Should -BeTrue
        Test-Path (Join-Path -Path $params.PersonalDir -ChildPath 'report.pdf') | Should -BeTrue
    }

    It 'accepts legacy keys in a top-level config file' {
        $configPath = Join-Path -Path $root -ChildPath 'legacy.json'
        @{ SourceDir = $params.SourceDir; PersonalDir = $params.PersonalDir; MusicDir = $params.MusicDir; FotosDir = $params.PhotosDir } |
            ConvertTo-Json | Set-Content -LiteralPath $configPath

        & (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1') -ConfigPath $configPath -LogPath $params.LogPath `
            -ReportPath $params.ReportPath -InformationAction SilentlyContinue -WarningAction SilentlyContinue

        Test-Path (Join-Path -Path $params.PhotosDir -ChildPath 'photo.jpg') | Should -BeTrue
    }

    It 'changes nothing with -WhatIf' {
        & (Join-Path -Path $repoRoot -ChildPath 'MovingFiles.ps1') @params -WhatIf
        @(Get-ChildItem -LiteralPath $params.SourceDir -Recurse -File).Count | Should -Be 5
        Test-Path $params.PersonalDir | Should -BeFalse
    }
}

Describe 'DirectoryReorder.ps1' {
    BeforeEach {
        $root = Join-Path -Path $TestDrive -ChildPath ([guid]::NewGuid().ToString('N'))
        $params = @{
            SourceDir         = $root
            PersonalDir       = Join-Path -Path $root -ChildPath 'Personal'
            PhotosDir         = Join-Path -Path $root -ChildPath 'Photos'
            MusicDir          = Join-Path -Path $root -ChildPath 'Music'
            LogPath           = Join-Path -Path $TestDrive -ChildPath "$([guid]::NewGuid().ToString('N')).log"
            ReportPath        = Join-Path -Path $TestDrive -ChildPath "DirectoryReorder_Report_$([guid]::NewGuid().ToString('N')).json"
            InformationAction = 'SilentlyContinue'
            WarningAction     = 'SilentlyContinue'
        }
        Add-TestFile ([System.IO.Path]::Combine($root, 'docs', 'paper.tex')) 'abc'
        Add-TestFile ([System.IO.Path]::Combine($params.PhotosDir, '2020', 'trip', 'p.jpg')) 'abcd'
        Add-TestFile ([System.IO.Path]::Combine($params.MusicDir, 'Artist', 'Album', 'song.mp3')) 'abcde'
        Add-TestFile (Join-Path -Path $params.MusicDir -ChildPath '.DS_Store')
    }

    It 'reorganizes documents, photos and music and totals the report sizes' {
        & (Join-Path -Path $repoRoot -ChildPath 'DirectoryReorder.ps1') @params

        Test-Path (Join-Path -Path $params.PersonalDir -ChildPath 'paper.tex') | Should -BeTrue
        Test-Path (Join-Path -Path $params.PhotosDir -ChildPath 'p.jpg') | Should -BeTrue
        Test-Path ([System.IO.Path]::Combine($params.MusicDir, 'Album', 'song.mp3')) | Should -BeTrue
        Test-Path (Join-Path -Path $params.MusicDir -ChildPath '.DS_Store') | Should -BeFalse

        $report = Get-Content -LiteralPath $params.ReportPath -Raw | ConvertFrom-Json
        $expectedBytes = $report.Sections.Documents.Size + $report.Sections.Photos.Size + $report.Sections.Music.Size
        $report.Summary.TotalSizeBytes | Should -Be $expectedBytes
        $report.Summary.TotalSizeGB | Should -Be ([Math]::Round($expectedBytes / 1GB, 2))
        $report.Summary.TotalErrors | Should -Be 0
    }

    It 'accepts the legacy config keys' {
        $configPath = Join-Path -Path $TestDrive -ChildPath "$([guid]::NewGuid().ToString('N')).json"
        @{ SourceRoot = $root; PersonalFolder = $params.PersonalDir; PhotoRoot = $params.PhotosDir; MusicRoot = $params.MusicDir } |
            ConvertTo-Json | Set-Content -LiteralPath $configPath
        & (Join-Path -Path $repoRoot -ChildPath 'DirectoryReorder.ps1') -ConfigPath $configPath -LogPath $params.LogPath `
            -ReportPath $params.ReportPath -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        Test-Path (Join-Path -Path $params.PersonalDir -ChildPath 'paper.tex') | Should -BeTrue
    }

    It 'accepts the legacy parameter names' {
        & (Join-Path -Path $repoRoot -ChildPath 'DirectoryReorder.ps1') -SourceRoot $root -PersonalFolder $params.PersonalDir -PhotoRoot $params.PhotosDir `
            -MusicRoot $params.MusicDir -LogPath $params.LogPath -ReportPath $params.ReportPath -DryRun `
            -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        Test-Path ([System.IO.Path]::Combine($root, 'docs', 'paper.tex')) | Should -BeTrue
    }
}

Describe 'DirectoryDistributionChecker.ps1' {
    It 'aggregates sizes and file counts per directory in a single pass' {
        $root = Join-Path -Path $TestDrive -ChildPath 'ddc'
        Add-TestFile (Join-Path -Path $root -ChildPath 'a.pdf') ('a' * 100)
        Add-TestFile ([System.IO.Path]::Combine($root, 'sub', 'b.mp3')) ('b' * 50)
        Add-TestFile ([System.IO.Path]::Combine($root, 'sub', 'deep', 'c.jpg')) ('c' * 25)

        $result = & (Join-Path -Path $repoRoot -ChildPath 'DirectoryDistributionChecker.ps1') -Path $root -PassThru `
            -LogPath (Join-Path -Path $TestDrive -ChildPath 'ddc.log') -InformationAction SilentlyContinue

        $rootEntry = $result | Where-Object Directory -eq (Get-Item -LiteralPath $root).FullName
        $rootEntry.FileCount | Should -Be 3
        $rootEntry.SizeBytes | Should -Be 175
        ($result | Where-Object Directory -like '*sub').SizeBytes | Should -Be 75
        ($result | Where-Object Directory -like '*deep').FileCount | Should -Be 1
    }

    It 'keeps directories that differ only by case apart on case-sensitive file systems' -Skip:([System.IO.Path]::DirectorySeparatorChar -eq '\') {
        $root = Join-Path -Path $TestDrive -ChildPath 'ddc-case'
        Add-TestFile ([System.IO.Path]::Combine($root, 'Dir', 'a.txt')) ('a' * 10)
        Add-TestFile ([System.IO.Path]::Combine($root, 'dir', 'b.txt')) ('b' * 20)

        $result = & (Join-Path -Path $repoRoot -ChildPath 'DirectoryDistributionChecker.ps1') -Path $root -PassThru `
            -LogPath (Join-Path -Path $TestDrive -ChildPath 'ddc-case.log') -InformationAction SilentlyContinue

        ($result | Where-Object { $_.Directory -ceq (Join-Path -Path $root -ChildPath 'Dir') }).SizeBytes | Should -Be 10
        ($result | Where-Object { $_.Directory -ceq (Join-Path -Path $root -ChildPath 'dir') }).SizeBytes | Should -Be 20
    }

    It 'exports a JSON report' {
        $root = Join-Path -Path $TestDrive -ChildPath 'ddc-export'
        Add-TestFile (Join-Path -Path $root -ChildPath 'a.txt')
        $out = Join-Path -Path $TestDrive -ChildPath 'ddc-out'
        & (Join-Path -Path $repoRoot -ChildPath 'DirectoryDistributionChecker.ps1') -Path $root -ExportFormat JSON -OutputPath $out `
            -LogPath (Join-Path -Path $TestDrive -ChildPath 'ddc2.log') -InformationAction SilentlyContinue | Out-Null
        @(Get-ChildItem -LiteralPath $out -Filter '*.json').Count | Should -Be 1
    }
}
