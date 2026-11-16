[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Pester shares variables between BeforeAll/BeforeEach and It blocks.')]
param()

BeforeAll {
    . ([System.IO.Path]::Combine($PSScriptRoot, '..', 'Common.ps1'))
    $InformationPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'
    Set-LogFile -Path (Join-Path -Path $TestDrive -ChildPath 'common.log')
}

Describe 'Format-FileSize' {
    It 'formats <Bytes> as <Expected>' -TestCases @(
        @{ Bytes = 0; Expected = '0 B' }
        @{ Bytes = 512; Expected = '512 B' }
        @{ Bytes = 1KB; Expected = ('{0:N2} KB' -f 1) }
        @{ Bytes = 1.5MB; Expected = ('{0:N2} MB' -f 1.5) }
        @{ Bytes = 2GB; Expected = ('{0:N2} GB' -f 2) }
        @{ Bytes = 3TB; Expected = ('{0:N2} TB' -f 3) }
    ) {
        Format-FileSize -Bytes $Bytes | Should -Be $Expected
    }
}

Describe 'Test-PathSafe' {
    It 'returns $false for null, empty or whitespace paths' {
        Test-PathSafe -Path $null | Should -BeFalse
        Test-PathSafe -Path '' | Should -BeFalse
        Test-PathSafe -Path '   ' | Should -BeFalse
    }

    It 'returns $true for an existing path and $false for a missing one' {
        Test-PathSafe -Path $TestDrive | Should -BeTrue
        Test-PathSafe -Path (Join-Path -Path $TestDrive -ChildPath 'does-not-exist') | Should -BeFalse
    }

    It 'treats wildcard characters literally' {
        $null = New-Item -ItemType File -Path (Join-Path -Path $TestDrive -ChildPath 'a.txt')
        Test-PathSafe -Path (Join-Path -Path $TestDrive -ChildPath '*.txt') | Should -BeFalse
    }
}

Describe 'ConvertTo-FullPath' {
    It 'resolves relative paths against the PowerShell location, not the process directory' {
        Push-Location -LiteralPath $TestDrive
        try {
            ConvertTo-FullPath -Path 'not-yet-created' | Should -Be (Join-Path -Path $TestDrive -ChildPath 'not-yet-created')
        }
        finally {
            Pop-Location
        }
    }
}

Describe 'New-FolderIfMissing' {
    It 'creates missing nested folders' {
        $path = [System.IO.Path]::Combine($TestDrive, 'x', 'y', 'z')
        New-FolderIfMissing -Path $path | Should -BeTrue
        Test-Path -LiteralPath $path -PathType Container | Should -BeTrue
    }

    It 'does not create anything with -WhatIf' {
        $path = Join-Path -Path $TestDrive -ChildPath 'whatif-folder'
        New-FolderIfMissing -Path $path -WhatIf | Should -BeTrue
        Test-Path -LiteralPath $path | Should -BeFalse
    }
}

Describe 'Write-Log' {
    It 'appends timestamped lines to the log file' {
        $log = Join-Path -Path $TestDrive -ChildPath 'write-log.log'
        Set-LogFile -Path $log
        Write-Log -Message 'hello world' -Level 'Success'
        Get-Content -LiteralPath $log | Should -Match '^\[\d{4}-\d{2}-\d{2} [\d:.]+\] \[Success\] hello world$'
    }

    It 'writes Info messages to the information stream' {
        $record = Write-Log -Message 'info message' -InformationAction Continue 6>&1
        "$record" | Should -Match '\[Info\] info message$'
    }

    It 'writes Error messages as non-terminating errors even when ErrorActionPreference is Stop' {
        $ErrorActionPreference = 'Stop'
        { Write-Log -Message 'boom' -Level 'Error' 2>$null } | Should -Not -Throw
    }
}

Describe 'Import-ScriptConfig and Resolve-Setting' {
    BeforeAll {
        $configPath = Join-Path -Path $TestDrive -ChildPath 'config.json'
        @{
            MovingFiles = @{ SourceDir = 'from-config'; TimeoutSeconds = 42 }
        } | ConvertTo-Json | Set-Content -LiteralPath $configPath
    }

    It 'returns an empty hashtable when no path is given' {
        (Import-ScriptConfig -Path '' -Section 'MovingFiles').Count | Should -Be 0
    }

    It 'throws for a missing configuration file' {
        { Import-ScriptConfig -Path (Join-Path -Path $TestDrive -ChildPath 'nope.json') } | Should -Throw '*not found*'
    }

    It 'reads the requested section' {
        $config = Import-ScriptConfig -Path $configPath -Section 'MovingFiles'
        $config.SourceDir | Should -Be 'from-config'
        $config.TimeoutSeconds | Should -Be 42
    }

    It 'prefers explicitly bound parameters over the config file' {
        $config = @{ SourceDir = 'from-config' }
        $bound = @{ SourceDir = 'from-parameter' }
        Resolve-Setting -Name 'SourceDir' -BoundParameters $bound -Config $config | Should -Be 'from-parameter'
    }

    It 'works with a real $PSBoundParameters dictionary' {
        function Test-Bound { param([string]$SourceDir) Resolve-Setting -Name 'SourceDir' -BoundParameters $PSBoundParameters -Config @{ SourceDir = 'cfg' } }
        Test-Bound -SourceDir 'param' | Should -Be 'param'
        Test-Bound | Should -Be 'cfg'
    }

    It 'accepts legacy config keys via -Alias' {
        Resolve-Setting -Name 'PhotosDir' -Alias 'FotosDir', 'PhotoRoot' -Config @{ PhotoRoot = 'legacy' } | Should -Be 'legacy'
        Resolve-Setting -Name 'PhotosDir' -Alias 'FotosDir' -Config @{ PhotosDir = 'new'; FotosDir = 'old' } | Should -Be 'new'
    }

    It 'falls back to the default and throws for missing required settings' {
        Resolve-Setting -Name 'Missing' -Default 7 | Should -Be 7
        { Resolve-Setting -Name 'Missing' -Required } | Should -Throw "*'Missing' is required*"
    }
}

Describe 'Report helpers' {
    BeforeEach {
        $report = Initialize-Report -Sections 'Documents', 'Photos', 'Music' -ExtraCounters @{ Music = @('Duplicates') } -Metadata @{ Script = 'Test' }
    }

    It 'initializes sections, extra counters and metadata' {
        $report.Script | Should -Be 'Test'
        $report.Status | Should -Be 'Running'
        $report.Sections.Keys | Should -Be @('Documents', 'Photos', 'Music')
        $report.Sections.Music.Duplicates | Should -Be 0
        $report.Errors.Count | Should -Be 0
    }

    It 'increments counters and sizes' {
        Update-Report -Report $report -Section 'Photos' -Metric 'Moved' -Size 100
        Update-Report -Report $report -Section 'Photos' -Metric 'Moved' -Size 50
        Update-Report -Report $report -Section 'Music' -Metric 'Duplicates'
        $report.Sections.Photos.Moved | Should -Be 2
        $report.Sections.Photos.Size | Should -Be 150
        $report.Sections.Music.Duplicates | Should -Be 1
    }

    It 'records the passed -ErrorMessage (not the automatic $Error variable)' {
        $Error.Clear()
        Update-Report -Report $report -ErrorMessage 'disk on fire' -NoLog
        $report.Errors.Count | Should -Be 1
        $report.Errors[0].Message | Should -Be 'disk on fire'
    }

    It 'does not record an error when no -ErrorMessage is passed, even if $Error is populated' {
        try { throw 'unrelated earlier failure' } catch { $null = $_ }
        $Error.Count | Should -BeGreaterThan 0
        Update-Report -Report $report -Section 'Documents' -Metric 'Processed'
        $report.Errors.Count | Should -Be 0
    }

    It 'records warnings' {
        Update-Report -Report $report -WarningMessage 'careful' -NoLog
        $report.Warnings[0].Message | Should -Be 'careful'
    }

    It 'divides the summed size of all sections by 1GB' {
        Update-Report -Report $report -Section 'Documents' -Metric 'Moved' -Size 1GB
        Update-Report -Report $report -Section 'Photos' -Metric 'Moved' -Size 2GB
        Update-Report -Report $report -Section 'Music' -Metric 'Moved' -Size 1GB
        $summary = Get-ReportSummary -Report $report
        $summary.TotalSizeBytes | Should -Be 4GB
        $summary.TotalSizeGB | Should -Be 4
        $summary.TotalFilesMoved | Should -Be 3
    }

    It 'saves a JSON report with a summary' {
        $path = [System.IO.Path]::Combine($TestDrive, 'reports', 'Test_Report_1.json')
        Update-Report -Report $report -Section 'Documents' -Metric 'Moved' -Size 10
        Save-Report -Report $report -Path $path
        $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $json.Status | Should -Be 'Completed'
        $json.Summary.TotalFilesMoved | Should -Be 1
        $json.Summary.TotalSizeBytes | Should -Be 10
    }
}

Describe 'Move-FileSafely' {
    BeforeEach {
        $source = Join-Path -Path $TestDrive -ChildPath "src-$([guid]::NewGuid().ToString('N'))"
        $target = Join-Path -Path $TestDrive -ChildPath "dst-$([guid]::NewGuid().ToString('N'))"
        $null = New-Item -ItemType Directory -Path $source
        $file = New-Item -ItemType File -Path (Join-Path -Path $source -ChildPath 'file.txt') -Value 'content'
        $destination = [System.IO.Path]::Combine($target, 'sub', 'file.txt')
        $report = Initialize-Report -Sections 'Docs'
    }

    It 'moves same-volume files with Move-Item and updates the report' {
        Mock Copy-Item { throw 'Copy-Item should not be used for same-volume moves' }
        Move-FileSafely -File $file -DestinationPath $destination -Report $report -Section 'Docs' | Should -Be 'Moved'
        Test-Path -LiteralPath $file.FullName | Should -BeFalse
        Get-Content -LiteralPath $destination | Should -Be 'content'
        $report.Sections.Docs.Processed | Should -Be 1
        $report.Sections.Docs.Moved | Should -Be 1
        $report.Sections.Docs.Size | Should -Be $file.Length
        Should -Invoke Copy-Item -Times 0
    }

    It 'uses copy + verify + delete for cross-volume moves' {
        Mock Test-SameVolume { $false }
        Move-FileSafely -File $file -DestinationPath $destination -Report $report -Section 'Docs' | Should -Be 'Moved'
        Test-Path -LiteralPath $file.FullName | Should -BeFalse
        Get-Content -LiteralPath $destination | Should -Be 'content'
        @(Get-ChildItem -LiteralPath (Split-Path $destination) -Filter '*.tmp').Count | Should -Be 0
    }

    It 'keeps the source and cleans up the temp file when cross-volume verification fails' {
        Mock Test-SameVolume { $false }
        Mock Get-Item { [pscustomobject]@{ Length = -1 } }
        Move-FileSafely -File $file -DestinationPath $destination -Report $report -Section 'Docs' 2>$null | Should -Be 'Failed'
        Test-Path -LiteralPath $file.FullName | Should -BeTrue
        Test-Path -LiteralPath $destination | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Split-Path $destination) -Force).Count | Should -Be 0
        $report.Sections.Docs.Failed | Should -Be 1
        $report.Errors[0].Message | Should -BeLike '*Copy verification failed*'
    }

    It 'skips existing destinations unless -Force is used' {
        $null = New-Item -ItemType File -Path $destination -Value 'old' -Force
        Move-FileSafely -File $file -DestinationPath $destination -Report $report -Section 'Docs' | Should -Be 'Skipped'
        Get-Content -LiteralPath $destination | Should -Be 'old'
        $report.Sections.Docs.Skipped | Should -Be 1

        Move-FileSafely -File $file -DestinationPath $destination -Report $report -Section 'Docs' -Force | Should -Be 'Moved'
        Get-Content -LiteralPath $destination | Should -Be 'content'
    }

    It 'does not move anything with -WhatIf' {
        Move-FileSafely -File $file -DestinationPath $destination -Report $report -Section 'Docs' -WhatIf | Should -Be 'WhatIf'
        Test-Path -LiteralPath $file.FullName | Should -BeTrue
        Test-Path -LiteralPath $destination | Should -BeFalse
    }
}
