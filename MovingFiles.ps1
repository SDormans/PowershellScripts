<#
.SYNOPSIS
    Organizes downloaded files into Personal (documents), Music and Photos folders.

.DESCRIPTION
    - Scans a source directory for files (recursively with -PreserveFolderStructure)
    - Moves documents, music and images to dedicated folders, creating them when needed
    - Skips files that already exist at the destination unless -Force is used
    - Same-volume moves are plain renames; cross-volume moves copy, verify and then delete
    - Writes a log file and a JSON execution report

    Folder settings are resolved in this order: explicit parameter > config file > error.
    There are no hardcoded default folders: pass them as parameters or via -ConfigPath
    (see config.example.json, section "MovingFiles").

.PARAMETER SourceDir
    Directory to scan for files.

.PARAMETER PersonalDir
    Destination for documents.

.PARAMETER MusicDir
    Destination for music files.

.PARAMETER PhotosDir
    Destination for images. Alias: FotosDir.

.PARAMETER ConfigPath
    Path to a JSON configuration file. Explicit parameters override values from the file.

.PARAMETER LogPath
    Path of the log file (default: script directory).

.PARAMETER ReportPath
    Path of the JSON execution report (default: script directory).

.PARAMETER TimeoutSeconds
    Maximum seconds to spend moving files (default: 300).

.PARAMETER PreserveFolderStructure
    Scan recursively and keep the relative subfolder structure at the destination.

.PARAMETER Force
    Overwrite files that already exist at the destination.

.EXAMPLE
    .\MovingFiles.ps1 -SourceDir "D:\Downloads" -PersonalDir "D:\Personal" -MusicDir "D:\Music" -PhotosDir "D:\Photos" -WhatIf

.EXAMPLE
    .\MovingFiles.ps1 -ConfigPath ".\config.json"

.EXAMPLE
    .\MovingFiles.ps1 -ConfigPath ".\config.json" -SourceDir "C:\Users\Public\Downloads" -PreserveFolderStructure
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDir,

    [Parameter(Mandatory = $false)]
    [string]$PersonalDir,

    [Parameter(Mandatory = $false)]
    [string]$MusicDir,

    [Parameter(Mandatory = $false)]
    [Alias('FotosDir')]
    [string]$PhotosDir,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path $PSScriptRoot "MovingFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"),

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = (Join-Path $PSScriptRoot "MovingFiles_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"),

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 3600)]
    [int]$TimeoutSeconds = 300,

    [Parameter(Mandatory = $false)]
    [switch]$PreserveFolderStructure,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not $PSBoundParameters.ContainsKey('InformationAction')) {
    $InformationPreference = 'Continue'
}

#region Configuration

$config = Import-ScriptConfig -Path $ConfigPath -Section 'MovingFiles'
$SourceDir = Resolve-Setting -Name 'SourceDir' -BoundParameters $PSBoundParameters -Config $config -Required
$PersonalDir = Resolve-Setting -Name 'PersonalDir' -BoundParameters $PSBoundParameters -Config $config -Required
$MusicDir = Resolve-Setting -Name 'MusicDir' -BoundParameters $PSBoundParameters -Config $config -Required
$PhotosDir = Resolve-Setting -Name 'PhotosDir' -Alias 'FotosDir' -BoundParameters $PSBoundParameters -Config $config -Required
$TimeoutSeconds = Resolve-Setting -Name 'TimeoutSeconds' -BoundParameters $PSBoundParameters -Config $config -Default $TimeoutSeconds
$PreserveFolderStructure = [bool](Resolve-Setting -Name 'PreserveFolderStructure' -BoundParameters $PSBoundParameters -Config $config -Default $false)

Set-LogFile -Path $LogPath

#endregion

#region Extension Definitions

$ExtensionMap = @{
    # Documents
    '.pdf' = 'Document'; '.doc' = 'Document'; '.docx' = 'Document'
    '.ppt' = 'Document'; '.pptx' = 'Document'; '.xls' = 'Document'
    '.xlsx' = 'Document'; '.txt' = 'Document'; '.rtf' = 'Document'
    '.odt' = 'Document'; '.ods' = 'Document'; '.odp' = 'Document'
    '.csv' = 'Document'; '.md' = 'Document'

    # Music
    '.mp3' = 'Music'; '.wav' = 'Music'; '.flac' = 'Music'; '.aac' = 'Music'
    '.m4a' = 'Music'; '.ogg' = 'Music'; '.wma' = 'Music'; '.opus' = 'Music'

    # Photos/Images
    '.jpg' = 'Photo'; '.jpeg' = 'Photo'; '.png' = 'Photo'; '.gif' = 'Photo'
    '.bmp' = 'Photo'; '.tiff' = 'Photo'; '.webp' = 'Photo'; '.heic' = 'Photo'
    '.svg' = 'Photo'; '.ico' = 'Photo'; '.raw' = 'Photo'; '.cr2' = 'Photo'
    '.nef' = 'Photo'
}

$Destinations = @{
    Document = $PersonalDir
    Music    = $MusicDir
    Photo    = $PhotosDir
}

#endregion

#region Helper Functions

function Get-DestinationPath {
    param(
        [System.IO.FileInfo]$File,
        [string]$DestinationFolder
    )

    if ($PreserveFolderStructure) {
        $sourceRoot = [System.IO.Path]::GetFullPath($SourceDir).TrimEnd('\', '/')
        $fileDir = $File.Directory.FullName
        if ($fileDir.Length -gt $sourceRoot.Length) {
            $relativePath = $fileDir.Substring($sourceRoot.Length).TrimStart('\', '/')
            return Join-Path (Join-Path $DestinationFolder $relativePath) $File.Name
        }
    }

    return Join-Path $DestinationFolder $File.Name
}

function Test-SourceDirectory {
    param([string]$Path)

    if (-not (Test-PathSafe -Path $Path)) {
        Write-Log -Message "Source directory does not exist: $Path" -Level 'Error'
        return $false
    }

    try {
        $null = Get-ChildItem -LiteralPath $Path -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log -Message "Cannot access source directory: $Path - $_" -Level 'Error'
        return $false
    }
}

#endregion

#region Main Execution

$Report = Initialize-Report -Sections 'Document', 'Music', 'Photo', 'Unknown' -Metadata ([ordered]@{
    Script          = 'MovingFiles'
    SourceDirectory = $SourceDir
})

try {
    Write-Log -Message '========== MOVING FILES STARTED =========='
    Write-Log -Message "Source directory: $SourceDir"
    Write-Log -Message "Personal directory: $PersonalDir"
    Write-Log -Message "Music directory: $MusicDir"
    Write-Log -Message "Photos directory: $PhotosDir"
    Write-Log -Message "Log file: $LogPath"
    Write-Log -Message "Report file: $ReportPath"
    Write-Log -Message "Timeout: $TimeoutSeconds seconds"
    Write-Log -Message "Preserve folder structure: $PreserveFolderStructure"

    if ($WhatIfPreference) {
        Write-Log -Message 'WHAT IF MODE - No files will be moved' -Level 'Warning'
    }

    if (-not (Test-SourceDirectory -Path $SourceDir)) {
        throw "Source directory validation failed: $SourceDir"
    }

    $allDestinationsValid = $true
    foreach ($category in 'Document', 'Music', 'Photo') {
        if (-not (New-FolderIfMissing -Path $Destinations[$category] -Description "$category folder" -Report $Report)) {
            $allDestinationsValid = $false
        }
    }
    if (-not $allDestinationsValid) {
        throw 'Some destination folders could not be created. Please check permissions.'
    }

    Write-Log -Message "Scanning for files in: $SourceDir"
    $files = @(Get-ChildItem -LiteralPath $SourceDir -File -Recurse:$PreserveFolderStructure -ErrorAction Stop)

    if ($files.Count -eq 0) {
        Write-Log -Message 'No files found in source directory.' -Level 'Warning'
    }
    else {
        Write-Log -Message "Found $($files.Count) file(s) to process" -Level 'Success'
        $timeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()

        foreach ($file in $files) {
            if ($timeoutTimer.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
                Update-Report -Report $Report -WarningMessage "Timeout reached ($TimeoutSeconds seconds). Processing stopped."
                break
            }

            $category = $null
            if ($file.Extension) {
                $category = $ExtensionMap[$file.Extension.ToLowerInvariant()]
            }

            if (-not $category) {
                Write-Log -Message "Skipping unsupported file type: $($file.Name)" -Level 'Debug'
                Update-Report -Report $Report -Section 'Unknown' -Metric 'Processed'
                Update-Report -Report $Report -Section 'Unknown' -Metric 'Skipped'
                continue
            }

            $destinationPath = Get-DestinationPath -File $file -DestinationFolder $Destinations[$category]
            $null = Move-FileSafely -File $file -DestinationPath $destinationPath -Report $Report -Section $category `
                -Description "$category folder" -Force:$Force
        }

        $summary = Get-ReportSummary -Report $Report
        Write-Log -Message '---------- SUMMARY ----------'
        Write-Log -Message "Files processed: $($summary.TotalFilesProcessed)"
        Write-Log -Message "Files moved: $($summary.TotalFilesMoved)" -Level 'Success'
        Write-Log -Message "Files skipped: $($summary.TotalFilesSkipped)"
        Write-Log -Message "Files failed: $($summary.TotalFilesFailed)"
        Write-Log -Message "Total size moved: $(Format-FileSize -Bytes $summary.TotalSizeBytes)"
        Write-Log -Message '-----------------------------'

        foreach ($category in $Report.Sections.Keys) {
            $section = $Report.Sections[$category]
            if ($section.Moved -gt 0 -or $section.Skipped -gt 0 -or $section.Failed -gt 0) {
                Write-Log -Message "$category : $($section.Moved) moved, $($section.Skipped) skipped, $($section.Failed) failed ($(Format-FileSize -Bytes $section.Size))" -Level 'Verbose'
            }
        }
    }
}
catch {
    $Report.Status = 'Failed'
    Update-Report -Report $Report -ErrorMessage "Critical error: $_" -NoLog
    Write-Log -Message "CRITICAL ERROR: $_" -Level 'Error'
    Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'Debug'
    throw
}
finally {
    Save-Report -Report $Report -Path $ReportPath

    Write-Log -Message '========== MOVING FILES COMPLETED ==========' -Level 'Success'
    Write-Log -Message "Log file saved to: $LogPath"

    if ($WhatIfPreference) {
        Write-Log -Message 'WHAT IF MODE was active - No actual changes were made' -Level 'Warning'
    }

    if ($Report.Sections.Values | Where-Object { $_.Failed -gt 0 }) {
        Write-Log -Message 'Some files failed to process. Check the report for details.' -Level 'Warning'
    }
}

#endregion
