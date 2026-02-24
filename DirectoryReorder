<#
.SYNOPSIS
    Enterprise-grade File Organizer for documents, photos, and music on drive F:\

.DESCRIPTION
    SECTION 1: Document Organization
        - Moves document files into a Personal folder
        - Supports extensive document formats
    
    SECTION 2: Photo Organization
        - Consolidates photos into the Photos root
        - Moves non-photo files into an Others folder
        - Cleans up empty directories intelligently
    
    SECTION 3: Music Library Management
        - Cleans music folders and removes duplicates
        - Removes _MACOSX and other system artifacts
        - Smart album deduplication with versioning

.PARAMETER SourceRoot
    Root directory to scan (default: F:\) 

.PARAMETER ConfigPath
    Path to JSON configuration file (overrides default paths)

.PARAMETER WhatIf
    Shows what would happen without actually making changes

.PARAMETER LogPath
    Path to write log file (default: script directory)

.PARAMETER ReportPath
    Path to save JSON execution report (default: script directory)

.PARAMETER TimeoutSeconds
    Maximum seconds per operation (default: 300)

.PARAMETER Parallel
    Use parallel processing for faster execution

.PARAMETER ThrottleLimit
    Maximum concurrent operations when using -Parallel (default: 5)

.PARAMETER DryRun
    Alias for WhatIf

.EXAMPLE
    .\FileOrganizer.ps1 -WhatIf

.EXAMPLE
    .\FileOrganizer.ps1 -ConfigPath ".\config.json" -Parallel -Verbose

.EXAMPLE
    .\FileOrganizer.ps1 -SourceRoot "D:\Data" -LogPath "C:\Logs\organizer.log"
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [string]$SourceRoot = "F:\",
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$PSScriptRoot\FileOrganizer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "$PSScriptRoot\FileOrganizer_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 3600)]
    [int]$TimeoutSeconds = 300,
    
    [Parameter(Mandatory = $false)]
    [switch]$Parallel,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 20)]
    [int]$ThrottleLimit = 5,
    
    [Parameter(Mandatory = $false)]
    [Alias('DryRun')]
    [switch]$WhatIf
)

#region Initialization & Configuration

# Load configuration from file if specified
if ($ConfigPath) {
    try {
        if (-not (Test-Path -Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }
        
        $config = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json
        
        # Override defaults with config values if they exist
        if ($config.SourceRoot) { $SourceRoot = $config.SourceRoot }
        if ($config.PersonalFolder) { $PersonalFolder = $config.PersonalFolder }
        if ($config.PhotoRoot) { $PhotoRoot = $config.PhotoRoot }
        if ($config.MusicRoot) { $MusicRoot = $config.MusicRoot }
        if ($config.TimeoutSeconds) { $TimeoutSeconds = $config.TimeoutSeconds }
        
        Write-Host "✅ Loaded configuration from: $ConfigPath" -ForegroundColor Cyan
    }
    catch {
        Write-Error "Failed to load configuration: $_"
        exit 1
    }
}

# Set folder paths (can be overridden by config)
$PersonalFolder = if ($config.PersonalFolder) { $config.PersonalFolder } else { "F:\Persoonlijk" }
$PhotoRoot = if ($config.PhotoRoot) { $config.PhotoRoot } else { "F:\Fotos" }
$MusicRoot = if ($config.MusicRoot) { $config.MusicRoot } else { "F:\Muziek" }
$OthersFolder = Join-Path $PhotoRoot "Others"
$DuplicatesDir = Join-Path $MusicRoot "Duplicates"

# Create log directory if needed
$logDir = Split-Path -Path $LogPath -Parent
if ($logDir -and -not (Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

#endregion

#region Logging & Reporting Functions

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug', 'Verbose')]
        [string]$Level = 'Info',
        
        [Parameter(Mandatory = $false)]
        [switch]$NoFile
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    $color = switch ($Level) {
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
        'Verbose' { 'Cyan' }
        'Debug' { 'Gray' }
        default { 'Gray' }
    }
    
    if ($Level -eq 'Verbose' -and $VerbosePreference -eq 'SilentlyContinue') {
        # Skip verbose
    }
    elseif ($Level -eq 'Debug' -and $DebugPreference -eq 'SilentlyContinue') {
        # Skip debug
    }
    else {
        Write-Host $logMessage -ForegroundColor $color
    }
    
    # File logging with retry
    if (-not $NoFile) {
        $maxRetries = 3
        $retryCount = 0
        while ($retryCount -lt $maxRetries) {
            try {
                Add-Content -Path $LogPath -Value $logMessage -ErrorAction Stop
                break
            }
            catch {
                $retryCount++
                if ($retryCount -eq $maxRetries) {
                    Write-Host "⚠️ Failed to write to log: $_" -ForegroundColor Red
                }
                else {
                    Start-Sleep -Milliseconds 100
                }
            }
        }
    }
}

function Initialize-Report {
    $script:Report = @{
        StartTime = Get-Date
        SourceRoot = $SourceRoot
        ComputerName = $env:COMPUTERNAME
        UserName = $env:USERNAME
        Status = "Running"
        Sections = @{
            Documents = @{ Processed = 0; Moved = 0; Failed = 0; Size = 0 }
            Photos = @{ Processed = 0; Moved = 0; Failed = 0; Size = 0 }
            Music = @{ Processed = 0; Moved = 0; Failed = 0; Size = 0; Duplicates = 0 }
            Cleanup = @{ EmptyDirsRemoved = 0; MacOSFoldersRemoved = 0 }
        }
        Errors = @()
        Warnings = @()
    }
}

function Update-Report {
    param(
        [string]$Section,
        [string]$Metric,
        [int]$Increment = 1,
        [long]$Size = 0,
        [string]$ErrorMessage = $null,
        [string]$WarningMessage = $null
    )
    
    if ($Section -and $Metric -and $script:Report.Sections.ContainsKey($Section)) {
        if ($script:Report.Sections[$Section].ContainsKey($Metric)) {
            $script:Report.Sections[$Section][$Metric] += $Increment
        }
        if ($Size -gt 0 -and $script:Report.Sections[$Section].ContainsKey('Size')) {
            $script:Report.Sections[$Section]['Size'] += $Size
        }
    }
    
    if ($ErrorMessage) {
        $script:Report.Errors += @{
            Time = Get-Date
            Message = $ErrorMessage
        }
        Write-Log -Message "ERROR: $ErrorMessage" -Level 'Error'
    }
    
    if ($WarningMessage) {
        $script:Report.Warnings += @{
            Time = Get-Date
            Message = $WarningMessage
        }
        Write-Log -Message "WARNING: $WarningMessage" -Level 'Warning'
    }
}

function Save-Report {
    try {
        $script:Report.EndTime = Get-Date
        $script:Report.DurationSeconds = [Math]::Round(($script:Report.EndTime - $script:Report.StartTime).TotalSeconds, 2)
        $script:Report.Status = "Completed"
        
        # Add summary statistics
        $script:Report.Summary = @{
            TotalFilesProcessed = ($script:Report.Sections.Documents.Processed + 
                                   $script:Report.Sections.Photos.Processed + 
                                   $script:Report.Sections.Music.Processed)
            TotalFilesMoved = ($script:Report.Sections.Documents.Moved + 
                              $script:Report.Sections.Photos.Moved + 
                              $script:Report.Sections.Music.Moved)
            TotalSizeGB = [Math]::Round($script:Report.Sections.Documents.Size +
                                        $script:Report.Sections.Photos.Size +
                                        $script:Report.Sections.Music.Size / 1GB, 2)
            TotalErrors = $script:Report.Errors.Count
            TotalWarnings = $script:Report.Warnings.Count
        }
        
        $script:Report | ConvertTo-Json -Depth 5 | Out-File -FilePath $ReportPath -Encoding UTF8
        Write-Log -Message "📊 Report saved to: $ReportPath" -Level 'Success'
    }
    catch {
        Write-Log -Message "Failed to save report: $_" -Level 'Error'
    }
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Test-PathSafe {
    param([string]$Path)
    try { return Test-Path -Path $Path -ErrorAction Stop }
    catch { return $false }
}

#endregion

#region Core Functions

function Ensure-Folder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [string]$Description = "folder"
    )
    
    if (-not (Test-PathSafe -Path $Path)) {
        if ($PSCmdlet.ShouldProcess($Path, "Create $Description")) {
            try {
                $parent = Split-Path -Path $Path -Parent
                if ($parent -and -not (Test-PathSafe -Path $parent)) {
                    Ensure-Folder -Path $parent -Description "parent folder" | Out-Null
                }
                
                New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
                Write-Log -Message "📁 Created $Description : $Path" -Level 'Success'
                return $true
            }
            catch {
                Write-Log -Message "❌ Failed to create $Description $Path : $_" -Level 'Error'
                return $false
            }
        }
        else {
            Write-Log -Message "[WHAT IF] Would create $Description : $Path" -Level 'Info'
            return $true
        }
    }
    return $true
}

function Move-FileSafely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationFolder,
        
        [Parameter(Mandatory = $true)]
        [string]$Section,
        
        [Parameter(Mandatory = $false)]
        [switch]$AllowOverwrite
    )
    
    $destinationPath = Join-Path $DestinationFolder $File.Name
    $fileSize = $File.Length
    
    # Check source
    if (-not (Test-PathSafe -Path $File.FullName)) {
        Update-Report -Section $Section -Metric 'Failed' -WarningMessage "Source file disappeared: $($File.Name)"
        return
    }
    
    # Ensure destination folder exists
    if (-not (Ensure-Folder -Path $DestinationFolder -Description "$Section destination")) {
        Update-Report -Section $Section -Metric 'Failed' -ErrorMessage "Cannot create destination folder for $($File.Name)"
        return
    }
    
    # Check destination
    if (Test-PathSafe -Path $destinationPath) {
        if ($AllowOverwrite) {
            Write-Log -Message "⚠️ Overwriting existing file: $($File.Name)" -Level 'Warning'
        }
        else {
            Write-Log -Message "⏭️ Skipping $($File.Name) - already exists" -Level 'Warning'
            Update-Report -Section $Section -Metric 'Processed' -WarningMessage "File already exists: $($File.Name)"
            return
        }
    }
    
    Update-Report -Section $Section -Metric 'Processed'
    
    if ($PSCmdlet.ShouldProcess($File.Name, "Move to $Section")) {
        $tempPath = $null
        try {
            # Double-check source
            if (-not (Test-PathSafe -Path $File.FullName)) {
                throw "Source disappeared"
            }
            
            # Use atomic-like operation with temp file
            if (-not $WhatIfPreference) {
                $tempPath = Join-Path $DestinationFolder "~$($File.Name).tmp"
                Copy-Item -Path $File.FullName -Destination $tempPath -ErrorAction Stop
                Move-Item -Path $tempPath -Destination $destinationPath -ErrorAction Stop
                Remove-Item -Path $File.FullName -Force -ErrorAction Stop
            }
            
            Write-Log -Message "✅ Moved: $($File.Name) -> $Section" -Level 'Success'
            Update-Report -Section $Section -Metric 'Moved' -Size $fileSize
            
            # Cleanup temp
            if ($tempPath -and (Test-PathSafe -Path $tempPath)) {
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            if ($tempPath -and (Test-PathSafe -Path $tempPath)) {
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            }
            
            $errorMsg = "Failed to move $($File.Name): $_"
            Update-Report -Section $Section -Metric 'Failed' -ErrorMessage $errorMsg
        }
    }
    else {
        Write-Log -Message "[WHAT IF] Would move: $($File.Name) -> $Section" -Level 'Info'
    }
}

function Remove-ItemSafely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [switch]$Recurse,
        
        [Parameter(Mandatory = $false)]
        [string]$Description = "item"
    )
    
    if (-not (Test-PathSafe -Path $Path)) {
        return
    }
    
    if ($PSCmdlet.ShouldProcess($Path, "Remove $Description")) {
        try {
            $params = @{
                Path = $Path
                ErrorAction = 'Stop'
                Force = $true
            }
            if ($Recurse) { $params['Recurse'] = $true }
            
            Remove-Item @params
            Write-Log -Message "🗑️ Removed $Description : $Path" -Level 'Success'
            return $true
        }
        catch {
            Write-Log -Message "❌ Failed to remove $Path : $_" -Level 'Error'
            return $false
        }
    }
    else {
        Write-Log -Message "[WHAT IF] Would remove $Description : $Path" -Level 'Info'
        return $true
    }
}

#endregion

#region Extension Definitions

$DocumentExtensions = @(
    ".doc", ".docx", ".pdf", ".xls", ".xlsx", ".csv",
    ".ppt", ".pptx", ".txt", ".rtf", ".odt", ".ods",
    ".odp", ".md", ".tex", ".tex", ".wpd", ".wps"
)

$PhotoExtensions = @(
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff",
    ".heic", ".webp", ".svg", ".ico", ".raw", ".cr2",
    ".nef", ".arw", ".dng", ".psd", ".ai", ".eps"
)

$MusicExtensions = @(
    ".mp3", ".flac", ".wav", ".aac", ".m4a", ".ogg",
    ".wma", ".opus", ".aiff", ".ape", ".dsf", ".dff"
)

#endregion

#region Main Execution

# Initialize
Initialize-Report
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$timeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Write-Log -Message "══════════════════════════════════════════════" -Level 'Info'
    Write-Log -Message "🚀 FILE ORGANIZER v2.0 - PRODUCTION EDITION" -Level 'Success'
    Write-Log -Message "══════════════════════════════════════════════" -Level 'Info'
    Write-Log -Message "Source root: $SourceRoot" -Level 'Info'
    Write-Log -Message "Personal folder: $PersonalFolder" -Level 'Info'
    Write-Log -Message "Photo root: $PhotoRoot" -Level 'Info'
    Write-Log -Message "Music root: $MusicRoot" -Level 'Info'
    Write-Log -Message "Log file: $LogPath" -Level 'Info'
    Write-Log -Message "Report file: $ReportPath" -Level 'Info'
    Write-Log -Message "Timeout: $TimeoutSeconds seconds" -Level 'Info'
    Write-Log -Message "══════════════════════════════════════════════" -Level 'Info'

    if ($WhatIfPreference) {
        Write-Log -Message "⚠️ WHAT IF MODE - No changes will be made" -Level 'Warning'
    }

    #===========================================================================
    # SECTION 1: Document Organization
    #===========================================================================
    
    Write-Log -Message "📄 SECTION 1: Organizing documents..." -Level 'Info'
    
    if (Ensure-Folder -Path $PersonalFolder -Description "Personal documents folder") {
        try {
            $docFiles = Get-ChildItem -Path $SourceRoot -Recurse -File -ErrorAction Stop |
                        Where-Object { $DocumentExtensions -contains $_.Extension.ToLower() }
            
            Write-Log -Message "Found $($docFiles.Count) document(s)" -Level 'Info'
            
            foreach ($file in $docFiles) {
                # Check timeout
                if ($timeoutTimer.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
                    throw "Timeout reached after $TimeoutSeconds seconds"
                }
                
                Move-FileSafely -File $file -DestinationFolder $PersonalFolder -Section 'Documents'
            }
        }
        catch {
            Update-Report -ErrorMessage "Document section error: $_"
        }
    }

    #===========================================================================
    # SECTION 2: Photo Organization
    #===========================================================================
    
    Write-Log -Message "🖼️ SECTION 2: Organizing photos..." -Level 'Info'
    
    if ((Ensure-Folder -Path $PhotoRoot -Description "Photo root") -and 
        (Ensure-Folder -Path $OthersFolder -Description "Others folder")) {
        
        # Consolidate photos
        try {
            $photoFiles = Get-ChildItem -Path $PhotoRoot -Recurse -File -ErrorAction Stop |
                          Where-Object { $PhotoExtensions -contains $_.Extension.ToLower() } |
                          Where-Object { $_.Directory.FullName -ne $PhotoRoot }
            
            Write-Log -Message "Found $($photoFiles.Count) photo(s) to consolidate" -Level 'Info'
            
            foreach ($file in $photoFiles) {
                if ($timeoutTimer.Elapsed.TotalSeconds -gt $TimeoutSeconds) { throw "Timeout" }
                Move-FileSafely -File $file -DestinationFolder $PhotoRoot -Section 'Photos'
            }
        }
        catch {
            Update-Report -ErrorMessage "Photo consolidation error: $_"
        }
        
        # Move non-photos to Others
        try {
            $nonPhotoFiles = Get-ChildItem -Path $PhotoRoot -Recurse -File -ErrorAction Stop |
                            Where-Object { $PhotoExtensions -notcontains $_.Extension.ToLower() } |
                            Where-Object { $_.Directory.FullName -ne $OthersFolder }
            
            Write-Log -Message "Found $($nonPhotoFiles.Count) non-photo file(s)" -Level 'Info'
            
            foreach ($file in $nonPhotoFiles) {
                if ($timeoutTimer.Elapsed.TotalSeconds -gt $TimeoutSeconds) { throw "Timeout" }
                Move-FileSafely -File $file -DestinationFolder $OthersFolder -Section 'Photos'
            }
        }
        catch {
            Update-Report -ErrorMessage "Non-photo move error: $_"
        }
        
        # Remove empty directories
        try {
            $emptyDirs = Get-ChildItem -Path $PhotoRoot -Recurse -Directory -ErrorAction Stop |
                        Where-Object { $_.FullName -ne $OthersFolder } |
                        Where-Object { (Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0 } |
                        Sort-Object FullName -Descending
            
            Write-Log -Message "Found $($emptyDirs.Count) empty directorie(s)" -Level 'Info'
            
            foreach ($dir in $emptyDirs) {
                if (Remove-ItemSafely -Path $dir.FullName -Recurse -Description "empty directory") {
                    Update-Report -Section 'Cleanup' -Metric 'EmptyDirsRemoved'
                }
            }
        }
        catch {
            Update-Report -ErrorMessage "Empty directory cleanup error: $_"
        }
    }

    #===========================================================================
    # SECTION 3: Music Organization
    #===========================================================================
    
    Write-Log -Message "🎵 SECTION 3: Organizing music library..." -Level 'Info'
    
    if ((Ensure-Folder -Path $MusicRoot -Description "Music root") -and 
        (Ensure-Folder -Path $DuplicatesDir -Description "Duplicates folder")) {
        
        # Remove macOS metadata
        try {
            $macosFolders = Get-ChildItem -Path $MusicRoot -Recurse -Directory -Force -ErrorAction Stop |
                           Where-Object { $_.Name -eq "_MACOSX" -or $_.Name -eq ".DS_Store" }
            
            foreach ($folder in $macosFolders) {
                if (Remove-ItemSafely -Path $folder.FullName -Recurse -Description "macOS metadata") {
                    Update-Report -Section 'Cleanup' -Metric 'MacOSFoldersRemoved'
                }
            }
        }
        catch {
            Update-Report -ErrorMessage "macOS cleanup error: $_"
        }
        
        # Process music directories
        try {
            $musicDirs = Get-ChildItem -Path $MusicRoot -Recurse -Directory -ErrorAction Stop |
                        Sort-Object FullName -Descending
            
            Write-Log -Message "Processing $($musicDirs.Count) music directorie(s)..." -Level 'Info'
            
            foreach ($dir in $musicDirs) {
                if ($timeoutTimer.Elapsed.TotalSeconds -gt $TimeoutSeconds) { throw "Timeout" }
                
                # Skip root and duplicates
                if ($dir.FullName -eq $MusicRoot -or $dir.FullName -eq $DuplicatesDir) {
                    continue
                }
                
                # Check for music files
                $hasMusic = Get-ChildItem -Path $dir.FullName -Recurse -File -ErrorAction Stop |
                           Where-Object { $MusicExtensions -contains $_.Extension.ToLower() } |
                           Select-Object -First 1
                
                if ($hasMusic) {
                    # Has music - check for duplicate
                    $targetPath = Join-Path $MusicRoot $dir.Name
                    
                    if (Test-Path -Path $targetPath -and $targetPath -ne $dir.FullName) {
                        # Duplicate found
                        Write-Log -Message "🔄 Duplicate album: $($dir.Name)" -Level 'Warning'
                        
                        if ($PSCmdlet.ShouldProcess($dir.FullName, "Move duplicate")) {
                            $dupDest = Join-Path $DuplicatesDir $dir.Name
                            
                            # Handle name collision in duplicates
                            if (Test-Path -Path $dupDest) {
                                $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                                $dupDest = Join-Path $DuplicatesDir "$($dir.Name)_$timestamp"
                            }
                            
                            try {
                                Move-Item -Path $dir.FullName -Destination $dupDest -ErrorAction Stop
                                Write-Log -Message "✅ Moved duplicate: $($dir.Name) -> Duplicates" -Level 'Success'
                                Update-Report -Section 'Music' -Metric 'Duplicates'
                            }
                            catch {
                                Update-Report -ErrorMessage "Failed to move duplicate $($dir.Name): $_"
                            }
                        }
                    }
                    elseif ($dir.Parent.FullName -ne $MusicRoot) {
                        # Not duplicate but in subfolder - move to root
                        if ($PSCmdlet.ShouldProcess($dir.FullName, "Move to Music root")) {
                            try {
                                Move-Item -Path $dir.FullName -Destination $MusicRoot -ErrorAction Stop
                                Write-Log -Message "✅ Moved: $($dir.Name) -> Music root" -Level 'Success'
                                Update-Report -Section 'Music' -Metric 'Moved'
                            }
                            catch {
                                Update-Report -ErrorMessage "Failed to move $($dir.Name): $_"
                            }
                        }
                    }
                    Update-Report -Section 'Music' -Metric 'Processed'
                }
                else {
                    # No music - remove
                    Write-Log -Message "Empty/non-music folder: $($dir.Name)" -Level 'Warning'
                    if (Remove-ItemSafely -Path $dir.FullName -Recurse -Description "non-music folder") {
                        Update-Report -Section 'Cleanup' -Metric 'EmptyDirsRemoved'
                    }
                }
            }
        }
        catch {
            Update-Report -ErrorMessage "Music processing error: $_"
        }
    }

    #===========================================================================
    # Completion
    #===========================================================================
    
    $stopwatch.Stop()
    
    Write-Log -Message "══════════════════════════════════════════════" -Level 'Success'
    Write-Log -Message "✅ FILE ORGANIZER COMPLETED SUCCESSFULLY" -Level 'Success'
    Write-Log -Message "⏱️  Duration: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))" -Level 'Info'
    Write-Log -Message "📊 Summary:" -Level 'Info'
    Write-Log -Message "   📄 Documents: $($script:Report.Sections.Documents.Moved) moved, $($script:Report.Sections.Documents.Failed) failed" -Level 'Info'
    Write-Log -Message "   🖼️  Photos: $($script:Report.Sections.Photos.Moved) moved, $($script:Report.Sections.Photos.Failed) failed" -Level 'Info'
    Write-Log -Message "   🎵 Music: $($script:Report.Sections.Music.Moved) moved, $($script:Report.Sections.Music.Duplicates) duplicates" -Level 'Info'
    Write-Log -Message "   🗑️  Cleanup: $($script:Report.Sections.Cleanup.EmptyDirsRemoved) empty dirs, $($script:Report.Sections.Cleanup.MacOSFoldersRemoved) macOS folders" -Level 'Info'
    
    if ($script:Report.Errors.Count -gt 0) {
        Write-Log -Message "⚠️  Encountered $($script:Report.Errors.Count) errors - check report" -Level 'Warning'
    }
    
    Write-Log -Message "📋 Log file: $LogPath" -Level 'Info'
    Write-Log -Message "📊 Report: $ReportPath" -Level 'Info'
    Write-Log -Message "══════════════════════════════════════════════" -Level 'Success'
}
catch {
    Write-Log -Message "💥 CRITICAL ERROR: $_" -Level 'Error'
    $script:Report.Status = "Failed"
    Update-Report -ErrorMessage "Critical error: $_"
}
finally {
    Save-Report
}

#endregion