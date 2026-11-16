<#
.SYNOPSIS
    Organizes documents, photos and music on a drive or directory tree.

.DESCRIPTION
    SECTION 1: Document Organization
        - Moves document files found under SourceDir into PersonalDir (flattened)
        - Files already inside PersonalDir are left alone

    SECTION 2: Photo Organization
        - Consolidates photos from subfolders into the PhotosDir root
        - Moves non-photo files into PhotosDir\Others
        - Removes empty directories

    SECTION 3: Music Library Management
        - Removes macOS artifacts (__MACOSX / _MACOSX folders, .DS_Store files)
        - Moves nested album folders to the MusicDir root
        - Moves duplicate albums into MusicDir\Duplicates (timestamped on name collision)
        - Removes folders that contain no music

    Folder settings are resolved in this order: explicit parameter > config file > error.
    There are no hardcoded default folders: pass them as parameters or via -ConfigPath
    (see config.example.json, section "DirectoryReorder").

.PARAMETER SourceDir
    Root directory to scan for documents. Alias: SourceRoot.

.PARAMETER PersonalDir
    Destination for documents. Alias: PersonalFolder.

.PARAMETER PhotosDir
    Photo library root. Alias: PhotoRoot.

.PARAMETER MusicDir
    Music library root. Alias: MusicRoot.

.PARAMETER ConfigPath
    Path to a JSON configuration file. Explicit parameters override values from the file.

.PARAMETER LogPath
    Path of the log file (default: script directory).

.PARAMETER ReportPath
    Path of the JSON execution report (default: script directory).

.PARAMETER TimeoutSeconds
    Maximum seconds for the whole run (default: 300).

.PARAMETER DryRun
    Same as -WhatIf.

.EXAMPLE
    .\DirectoryReorder.ps1 -SourceDir "D:\" -PersonalDir "D:\Personal" -PhotosDir "D:\Photos" -MusicDir "D:\Music" -WhatIf

.EXAMPLE
    .\DirectoryReorder.ps1 -ConfigPath ".\config.json" -Verbose

.EXAMPLE
    .\DirectoryReorder.ps1 -ConfigPath ".\config.json" -SourceDir "D:\Data" -LogPath "C:\Logs\organizer.log"
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [Alias('SourceRoot')]
    [string]$SourceDir,

    [Parameter(Mandatory = $false)]
    [Alias('PersonalFolder')]
    [string]$PersonalDir,

    [Parameter(Mandatory = $false)]
    [Alias('PhotoRoot')]
    [string]$PhotosDir,

    [Parameter(Mandatory = $false)]
    [Alias('MusicRoot')]
    [string]$MusicDir,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path $PSScriptRoot "DirectoryReorder_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"),

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = (Join-Path $PSScriptRoot "DirectoryReorder_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"),

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 3600)]
    [int]$TimeoutSeconds = 300,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not $PSBoundParameters.ContainsKey('InformationAction')) {
    $InformationPreference = 'Continue'
}
if ($DryRun) {
    $WhatIfPreference = $true
}

#region Configuration

$config = Import-ScriptConfig -Path $ConfigPath -Section 'DirectoryReorder'
$SourceDir = Resolve-Setting -Name 'SourceDir' -BoundParameters $PSBoundParameters -Config $config -Required
$PersonalDir = Resolve-Setting -Name 'PersonalDir' -BoundParameters $PSBoundParameters -Config $config -Required
$PhotosDir = Resolve-Setting -Name 'PhotosDir' -BoundParameters $PSBoundParameters -Config $config -Required
$MusicDir = Resolve-Setting -Name 'MusicDir' -BoundParameters $PSBoundParameters -Config $config -Required
$TimeoutSeconds = Resolve-Setting -Name 'TimeoutSeconds' -BoundParameters $PSBoundParameters -Config $config -Default $TimeoutSeconds

$OthersDir = Join-Path $PhotosDir 'Others'
$DuplicatesDir = Join-Path $MusicDir 'Duplicates'

Set-LogFile -Path $LogPath

#endregion

#region Helper Functions

function Get-NormalizedPath {
    param([string]$Path)
    return [System.IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
}

function Test-PathUnder {
    <# Returns $true when $Path equals $Root or is inside it. #>
    param([string]$Path, [string]$Root)

    $p = Get-NormalizedPath -Path $Path
    $r = Get-NormalizedPath -Path $Root
    if ($p.Equals($r, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $p.StartsWith($r + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
           $p.StartsWith($r + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-PathEqual {
    param([string]$Path, [string]$OtherPath)
    return (Get-NormalizedPath -Path $Path).Equals((Get-NormalizedPath -Path $OtherPath), [System.StringComparison]::OrdinalIgnoreCase)
}

function Remove-ItemSafely {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [switch]$Recurse,

        [Parameter(Mandatory = $false)]
        [string]$Description = 'item'
    )

    if (-not (Test-PathSafe -Path $Path)) {
        return $false
    }

    if (-not $PSCmdlet.ShouldProcess($Path, "Remove $Description")) {
        Write-Log -Message "[WHAT IF] Would remove $Description : $Path"
        return $true
    }

    try {
        Remove-Item -LiteralPath $Path -Force -Recurse:$Recurse -ErrorAction Stop -WhatIf:$false -Confirm:$false
        Write-Log -Message "Removed $Description : $Path" -Level 'Success'
        return $true
    }
    catch {
        Update-Report -Report $Report -ErrorMessage "Failed to remove $Path : $_"
        return $false
    }
}

function Get-FileSafely {
    <# Recursive file listing that logs inaccessible paths instead of aborting. #>
    param([string]$Path, [switch]$Force)

    $scanErrors = $null
    $items = @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force:$Force -ErrorAction SilentlyContinue -ErrorVariable scanErrors)
    foreach ($scanError in $scanErrors) {
        Update-Report -Report $Report -WarningMessage "Could not scan: $($scanError.TargetObject) ($($scanError.Exception.Message))"
    }
    return $items
}

function Assert-NotTimedOut {
    if ($timeoutTimer.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
        throw "Timeout reached after $TimeoutSeconds seconds"
    }
}

#endregion

#region Extension Definitions

$DocumentExtensions = @(
    '.doc', '.docx', '.pdf', '.xls', '.xlsx', '.csv',
    '.ppt', '.pptx', '.txt', '.rtf', '.odt', '.ods',
    '.odp', '.md', '.tex', '.wpd', '.wps'
)

$PhotoExtensions = @(
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
    '.heic', '.webp', '.svg', '.ico', '.raw', '.cr2',
    '.nef', '.arw', '.dng', '.psd', '.ai', '.eps'
)

$MusicExtensions = @(
    '.mp3', '.flac', '.wav', '.aac', '.m4a', '.ogg',
    '.wma', '.opus', '.aiff', '.ape', '.dsf', '.dff'
)

#endregion

#region Main Execution

$Report = Initialize-Report -Sections 'Documents', 'Photos', 'Music', 'Cleanup' `
    -ExtraCounters @{ Music = @('Duplicates'); Cleanup = @('EmptyDirsRemoved', 'MacOSItemsRemoved') } `
    -Metadata ([ordered]@{ Script = 'DirectoryReorder'; SourceDirectory = $SourceDir })

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$timeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Write-Log -Message '=============================================='
    Write-Log -Message 'DIRECTORY REORDER STARTED' -Level 'Success'
    Write-Log -Message "Source directory: $SourceDir"
    Write-Log -Message "Personal directory: $PersonalDir"
    Write-Log -Message "Photos directory: $PhotosDir"
    Write-Log -Message "Music directory: $MusicDir"
    Write-Log -Message "Log file: $LogPath"
    Write-Log -Message "Report file: $ReportPath"
    Write-Log -Message "Timeout: $TimeoutSeconds seconds"
    Write-Log -Message '=============================================='

    if ($WhatIfPreference) {
        Write-Log -Message 'WHAT IF MODE - No changes will be made' -Level 'Warning'
    }

    if (-not (Test-PathSafe -Path $SourceDir)) {
        throw "Source directory does not exist: $SourceDir"
    }

    #===========================================================================
    # SECTION 1: Document Organization
    #===========================================================================

    Write-Log -Message 'SECTION 1: Organizing documents...'

    if (New-FolderIfMissing -Path $PersonalDir -Description 'Personal documents folder' -Report $Report) {
        try {
            $docFiles = @(Get-FileSafely -Path $SourceDir |
                Where-Object { $DocumentExtensions -contains $_.Extension.ToLowerInvariant() } |
                Where-Object { -not (Test-PathUnder -Path $_.FullName -Root $PersonalDir) })

            Write-Log -Message "Found $($docFiles.Count) document(s)"

            foreach ($file in $docFiles) {
                Assert-NotTimedOut
                $null = Move-FileSafely -File $file -DestinationPath (Join-Path $PersonalDir $file.Name) -Report $Report -Section 'Documents' -Description 'Personal folder'
            }
        }
        catch {
            Update-Report -Report $Report -ErrorMessage "Document section error: $_"
        }
    }

    #===========================================================================
    # SECTION 2: Photo Organization
    #===========================================================================

    Write-Log -Message 'SECTION 2: Organizing photos...'

    if ((New-FolderIfMissing -Path $PhotosDir -Description 'Photos root' -Report $Report) -and
        (New-FolderIfMissing -Path $OthersDir -Description 'Others folder' -Report $Report)) {

        try {
            $photoFiles = @(Get-FileSafely -Path $PhotosDir |
                Where-Object { $PhotoExtensions -contains $_.Extension.ToLowerInvariant() } |
                Where-Object { -not (Test-PathEqual -Path $_.Directory.FullName -OtherPath $PhotosDir) })

            Write-Log -Message "Found $($photoFiles.Count) photo(s) to consolidate"

            foreach ($file in $photoFiles) {
                Assert-NotTimedOut
                $null = Move-FileSafely -File $file -DestinationPath (Join-Path $PhotosDir $file.Name) -Report $Report -Section 'Photos' -Description 'Photos root'
            }
        }
        catch {
            Update-Report -Report $Report -ErrorMessage "Photo consolidation error: $_"
        }

        try {
            $nonPhotoFiles = @(Get-FileSafely -Path $PhotosDir |
                Where-Object { $PhotoExtensions -notcontains $_.Extension.ToLowerInvariant() } |
                Where-Object { -not (Test-PathUnder -Path $_.FullName -Root $OthersDir) })

            Write-Log -Message "Found $($nonPhotoFiles.Count) non-photo file(s)"

            foreach ($file in $nonPhotoFiles) {
                Assert-NotTimedOut
                $null = Move-FileSafely -File $file -DestinationPath (Join-Path $OthersDir $file.Name) -Report $Report -Section 'Photos' -Description 'Others folder'
            }
        }
        catch {
            Update-Report -Report $Report -ErrorMessage "Non-photo move error: $_"
        }

        try {
            # Deepest first, and emptiness is checked at removal time so parents that only
            # contained empty folders are removed as well.
            $candidateDirs = @(Get-ChildItem -LiteralPath $PhotosDir -Recurse -Directory -ErrorAction Stop |
                Where-Object { -not (Test-PathUnder -Path $_.FullName -Root $OthersDir) } |
                Sort-Object { $_.FullName.Length } -Descending)

            foreach ($dir in $candidateDirs) {
                if (-not (Test-PathSafe -Path $dir.FullName)) { continue }
                if (@(Get-ChildItem -LiteralPath $dir.FullName -Force -ErrorAction SilentlyContinue).Count -ne 0) { continue }

                if (Remove-ItemSafely -Path $dir.FullName -Recurse -Description 'empty directory') {
                    Update-Report -Report $Report -Section 'Cleanup' -Metric 'EmptyDirsRemoved'
                }
            }
        }
        catch {
            Update-Report -Report $Report -ErrorMessage "Empty directory cleanup error: $_"
        }
    }

    #===========================================================================
    # SECTION 3: Music Organization
    #===========================================================================

    Write-Log -Message 'SECTION 3: Organizing music library...'

    if ((New-FolderIfMissing -Path $MusicDir -Description 'Music root' -Report $Report) -and
        (New-FolderIfMissing -Path $DuplicatesDir -Description 'Duplicates folder' -Report $Report)) {

        try {
            $macosItems = @(Get-ChildItem -LiteralPath $MusicDir -Recurse -Force -ErrorAction Stop |
                Where-Object {
                    ($_.PSIsContainer -and ($_.Name -eq '__MACOSX' -or $_.Name -eq '_MACOSX')) -or
                    (-not $_.PSIsContainer -and $_.Name -eq '.DS_Store')
                })

            foreach ($item in $macosItems) {
                if (Remove-ItemSafely -Path $item.FullName -Recurse:$item.PSIsContainer -Description 'macOS metadata') {
                    Update-Report -Report $Report -Section 'Cleanup' -Metric 'MacOSItemsRemoved'
                }
            }
        }
        catch {
            Update-Report -Report $Report -ErrorMessage "macOS cleanup error: $_"
        }

        try {
            $musicDirs = @(Get-ChildItem -LiteralPath $MusicDir -Recurse -Directory -ErrorAction Stop |
                Where-Object { -not (Test-PathUnder -Path $_.FullName -Root $DuplicatesDir) } |
                Sort-Object { $_.FullName.Length } -Descending)

            Write-Log -Message "Processing $($musicDirs.Count) music directorie(s)..."

            foreach ($dir in $musicDirs) {
                Assert-NotTimedOut

                if (-not (Test-PathSafe -Path $dir.FullName)) { continue }

                $hasMusic = Get-ChildItem -LiteralPath $dir.FullName -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object { $MusicExtensions -contains $_.Extension.ToLowerInvariant() } |
                    Select-Object -First 1

                if (-not $hasMusic) {
                    Write-Log -Message "Empty/non-music folder: $($dir.FullName)" -Level 'Warning'
                    if (Remove-ItemSafely -Path $dir.FullName -Recurse -Description 'non-music folder') {
                        Update-Report -Report $Report -Section 'Cleanup' -Metric 'EmptyDirsRemoved'
                    }
                    continue
                }

                Update-Report -Report $Report -Section 'Music' -Metric 'Processed'
                $isAtRoot = Test-PathEqual -Path $dir.Parent.FullName -OtherPath $MusicDir
                if ($isAtRoot) { continue }

                $targetPath = Join-Path $MusicDir $dir.Name
                if (Test-PathSafe -Path $targetPath) {
                    Write-Log -Message "Duplicate album: $($dir.Name)" -Level 'Warning'
                    $dupDest = Join-Path $DuplicatesDir $dir.Name
                    if (Test-PathSafe -Path $dupDest) {
                        $dupDest = Join-Path $DuplicatesDir "$($dir.Name)_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                    }

                    if ($PSCmdlet.ShouldProcess($dir.FullName, "Move duplicate album to $dupDest")) {
                        try {
                            Move-Item -LiteralPath $dir.FullName -Destination $dupDest -ErrorAction Stop
                            Write-Log -Message "Moved duplicate: $($dir.FullName) -> $dupDest" -Level 'Success'
                            Update-Report -Report $Report -Section 'Music' -Metric 'Duplicates'
                        }
                        catch {
                            Update-Report -Report $Report -Section 'Music' -Metric 'Failed' -ErrorMessage "Failed to move duplicate $($dir.FullName): $_"
                        }
                    }
                }
                elseif ($PSCmdlet.ShouldProcess($dir.FullName, "Move album to $MusicDir")) {
                    try {
                        Move-Item -LiteralPath $dir.FullName -Destination $targetPath -ErrorAction Stop
                        Write-Log -Message "Moved: $($dir.FullName) -> $targetPath" -Level 'Success'
                        Update-Report -Report $Report -Section 'Music' -Metric 'Moved'
                    }
                    catch {
                        Update-Report -Report $Report -Section 'Music' -Metric 'Failed' -ErrorMessage "Failed to move $($dir.FullName): $_"
                    }
                }
            }
        }
        catch {
            Update-Report -Report $Report -ErrorMessage "Music processing error: $_"
        }
    }

    #===========================================================================
    # Completion
    #===========================================================================

    $stopwatch.Stop()
    $sections = $Report.Sections

    Write-Log -Message '==============================================' -Level 'Success'
    Write-Log -Message 'DIRECTORY REORDER COMPLETED' -Level 'Success'
    Write-Log -Message "Duration: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
    Write-Log -Message 'Summary:'
    Write-Log -Message "   Documents: $($sections.Documents.Moved) moved, $($sections.Documents.Skipped) skipped, $($sections.Documents.Failed) failed"
    Write-Log -Message "   Photos: $($sections.Photos.Moved) moved, $($sections.Photos.Skipped) skipped, $($sections.Photos.Failed) failed"
    Write-Log -Message "   Music: $($sections.Music.Moved) moved, $($sections.Music.Duplicates) duplicates, $($sections.Music.Failed) failed"
    Write-Log -Message "   Cleanup: $($sections.Cleanup.EmptyDirsRemoved) empty dirs, $($sections.Cleanup.MacOSItemsRemoved) macOS items"

    if ($Report.Errors.Count -gt 0) {
        Write-Log -Message "Encountered $($Report.Errors.Count) error(s) - check the report" -Level 'Warning'
    }

    Write-Log -Message "Log file: $LogPath"
    Write-Log -Message "Report: $ReportPath"
    Write-Log -Message '==============================================' -Level 'Success'
}
catch {
    $Report.Status = 'Failed'
    Update-Report -Report $Report -ErrorMessage "Critical error: $_" -NoLog
    Write-Log -Message "CRITICAL ERROR: $_" -Level 'Error'
    throw
}
finally {
    Save-Report -Report $Report -Path $ReportPath
}

#endregion
