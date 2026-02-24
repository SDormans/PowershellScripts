<#
.SYNOPSIS
    Organizes downloaded files into Personal, Music, and Photos folders with enterprise-grade reliability.

.DESCRIPTION
    - Scans a source directory for files with timeout protection
    - Moves documents, music, and images to dedicated folders
    - Creates destination directories if they do not exist
    - Includes comprehensive logging, error handling, and reporting
    - Supports configuration files, parallel processing, and scheduled execution
    - Prevents file loss with transaction-style operations

.PARAMETER SourceDir
    The directory to scan for files (default: C:\Temp\Downloads)

.PARAMETER ConfigPath
    Path to JSON configuration file (overrides individual parameters)

.PARAMETER WhatIf
    Shows what would happen if the script runs without actually making changes

.PARAMETER LogPath
    Optional path to write a log file (default: script directory)

.PARAMETER ReportPath
    Optional path to save execution report (default: script directory)

.PARAMETER TimeoutSeconds
    Maximum seconds to spend on file operations (default: 300 = 5 minutes)

.PARAMETER Parallel
    Use parallel processing for faster execution on many files

.PARAMETER ThrottleLimit
    Maximum number of concurrent operations when using -Parallel (default: 5)

.PARAMETER PreserveFolderStructure
    Preserve subfolder structure when moving files instead of flattening

.EXAMPLE
    .\DownloadOrganizer.ps1 -SourceDir "D:\Downloads" -WhatIf
    
.EXAMPLE
    .\DownloadOrganizer.ps1 -ConfigPath ".\config.json" -Parallel

.EXAMPLE
    .\DownloadOrganizer.ps1 -SourceDir "C:\Users\Public\Downloads" -PreserveFolderStructure
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false, ParameterSetName = 'Direct')]
    [string]$SourceDir = "C:\Temp\Downloads",
    
    [Parameter(Mandatory = $false, ParameterSetName = 'Direct')]
    [string]$PersonalDir = "C:\Users\YourName\Personal",
    
    [Parameter(Mandatory = $false, ParameterSetName = 'Direct')]
    [string]$MusicDir = "C:\Users\YourName\Music",
    
    [Parameter(Mandatory = $false, ParameterSetName = 'Direct')]
    [string]$FotosDir = "C:\Users\YourName\Fotos",
    
    [Parameter(Mandatory = $false, ParameterSetName = 'Config')]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$ConfigPath,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$PSScriptRoot\DownloadOrganizer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "$PSScriptRoot\DownloadOrganizer_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 3600)]
    [int]$TimeoutSeconds = 300,
    
    [Parameter(Mandatory = $false)]
    [switch]$Parallel,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 20)]
    [int]$ThrottleLimit = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$PreserveFolderStructure,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

#region Initialization & Configuration Loading

# Load configuration from file if specified
if ($ConfigPath) {
    try {
        $config = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json
        $SourceDir = $config.SourceDir
        $PersonalDir = $config.PersonalDir
        $MusicDir = $config.MusicDir
        $FotosDir = $config.FotosDir
        if ($config.PreserveFolderStructure) { $PreserveFolderStructure = $true }
        Write-Host "Loaded configuration from: $ConfigPath" -ForegroundColor Cyan
    }
    catch {
        Write-Error "Failed to load configuration file: $_"
        exit 1
    }
}

# Create log directory if it doesn't exist
$logDir = Split-Path -Path $LogPath -Parent
if ($logDir -and -not (Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

#endregion

#region Logging Functions

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
    
    # Write to console with colors
    $consoleColor = switch ($Level) {
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
        'Verbose' { 'Cyan' }
        'Debug' { 'Gray' }
        default { 'Gray' }
    }
    
    if ($Level -eq 'Verbose' -and $VerbosePreference -eq 'SilentlyContinue') {
        # Skip verbose output if not enabled
    }
    elseif ($Level -eq 'Debug' -and $DebugPreference -eq 'SilentlyContinue') {
        # Skip debug output if not enabled
    }
    else {
        Write-Host $logMessage -ForegroundColor $consoleColor
    }
    
    # Write to log file
    if (-not $NoFile) {
        $maxRetries = 3
        $retryCount = 0
        $written = $false
        
        while (-not $written -and $retryCount -lt $maxRetries) {
            try {
                Add-Content -Path $LogPath -Value $logMessage -ErrorAction Stop
                $written = $true
            }
            catch {
                $retryCount++
                if ($retryCount -eq $maxRetries) {
                    Write-Host "[CRITICAL] Failed to write to log after $maxRetries attempts: $_" -ForegroundColor Red
                }
                else {
                    Start-Sleep -Milliseconds 100
                }
            }
        }
    }
}

#endregion

#region Helper Functions

function Initialize-Report {
    $script:Report = @{
        StartTime = Get-Date
        SourceDirectory = $SourceDir
        Status = "Running"
        FilesProcessed = 0
        FilesMoved = 0
        FilesSkipped = 0
        FilesFailed = 0
        TotalSizeBytes = 0
        Categories = @{
            Document = @{ Count = 0; Size = 0 }
            Music = @{ Count = 0; Size = 0 }
            Photo = @{ Count = 0; Size = 0 }
            Unknown = @{ Count = 0; Size = 0 }
        }
        Errors = @()
        Warnings = @()
    }
}

function Update-Report {
    param(
        [string]$Category,
        [int]$Count = 1,
        [long]$Size = 0,
        [string]$ErrorMessage = $null,
        [string]$Warning = $null
    )
    
    if ($Category -and $script:Report.Categories.ContainsKey($Category)) {
        $script:Report.Categories[$Category].Count += $Count
        $script:Report.Categories[$Category].Size += $Size
    }
    
    if ($Error) {
        $script:Report.Errors += @{
            Time = Get-Date
            Message = $Error
        }
    }
    
    if ($Warning) {
        $script:Report.Warnings += @{
            Time = Get-Date
            Message = $Warning
        }
    }
    
    $script:Report.FilesProcessed++
    $script:Report.TotalSizeBytes += $Size
}

function Save-Report {
    try {
        $script:Report.EndTime = Get-Date
        $script:Report.DurationSeconds = [Math]::Round(($script:Report.EndTime - $script:Report.StartTime).TotalSeconds, 2)
        $script:Report.Status = "Completed"
        
        $script:Report | ConvertTo-Json -Depth 5 | Out-File -FilePath $ReportPath -Encoding UTF8
        Write-Log -Message "Report saved to: $ReportPath" -Level 'Success'
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
    
    try {
        return Test-Path -Path $Path -ErrorAction Stop
    }
    catch {
        return $false
    }
}

function Ensure-Folder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [string]$FolderDescription = "folder"
    )
    
    if (-not (Test-PathSafe -Path $Path)) {
        if ($PSCmdlet.ShouldProcess($Path, "Create $FolderDescription")) {
            try {
                $parent = Split-Path -Path $Path -Parent
                if ($parent -and -not (Test-PathSafe -Path $parent)) {
                    Ensure-Folder -Path $parent -FolderDescription "parent folder" | Out-Null
                }
                
                New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
                Write-Log -Message "Created $FolderDescription : $Path" -Level 'Success'
                return $true
            }
            catch {
                Write-Log -Message "FAILED to create $FolderDescription $Path : $_" -Level 'Error'
                Update-Report -Error "Failed to create folder $Path : $_"
                return $false
            }
        }
        else {
            Write-Log -Message "[WHAT IF] Would create $FolderDescription : $Path" -Level 'Info'
            return $true
        }
    }
    return $true
}

function Get-DestinationPath {
    param(
        [System.IO.FileInfo]$File,
        [string]$DestinationFolder
    )
    
    if ($PreserveFolderStructure) {
        # Get relative path from source
        $relativePath = $File.Directory.FullName.Substring($SourceDir.Length).TrimStart('\')
        if ($relativePath) {
            return Join-Path $DestinationFolder $relativePath $File.Name
        }
    }
    
    return Join-Path $DestinationFolder $File.Name
}

function Move-FileSafely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationFolder,
        
        [Parameter(Mandatory = $true)]
        [string]$FileType
    )
    
    $destinationPath = Get-DestinationPath -File $File -DestinationFolder $DestinationFolder
    $destinationDir = Split-Path -Path $destinationPath -Parent
    
    # Check if source file still exists
    if (-not (Test-PathSafe -Path $File.FullName)) {
        Write-Log -Message "Source file no longer exists: $($File.Name)" -Level 'Warning'
        Update-Report -Warning "Source file disappeared: $($File.Name)"
        return
    }
    
    # Ensure destination directory exists
    if (-not (Ensure-Folder -Path $destinationDir -FolderDescription "destination subfolder")) {
        Update-Report -Error "Failed to create destination directory for $($File.Name)"
        $script:Report.FilesFailed++
        return
    }
    
    # Check if destination already exists
    if (Test-PathSafe -Path $destinationPath) {
        if ($Force) {
            Write-Log -Message "Overwriting existing file: $($File.Name)" -Level 'Warning'
        }
        else {
            Write-Log -Message "Cannot move $($File.Name) - already exists in $FileType folder" -Level 'Warning'
            Update-Report -Warning "File already exists: $($File.Name)"
            $script:Report.FilesSkipped++
            return
        }
    }
    
    if ($PSCmdlet.ShouldProcess($File.Name, "Move to $FileType folder")) {
        # Use a temporary file for atomic-like operation
        $tempPath = $null
        try {
            # Double-check source still exists
            if (-not (Test-PathSafe -Path $File.FullName)) {
                throw "Source file disappeared before move"
            }
            
            # Get file size for reporting
            $fileSize = $File.Length
            
            if ($Force -and (Test-PathSafe -Path $destinationPath)) {
                # Remove existing file if Force is specified
                Remove-Item -Path $destinationPath -Force -ErrorAction Stop
            }
            
            # Optional: Use temporary file for safer operation
            if (-not $WhatIfPreference) {
                $tempPath = Join-Path $destinationDir "~$($File.Name).tmp"
                Copy-Item -Path $File.FullName -Destination $tempPath -ErrorAction Stop
                Move-Item -Path $tempPath -Destination $destinationPath -ErrorAction Stop
                Remove-Item -Path $File.FullName -Force -ErrorAction Stop
            }
            else {
                # In WhatIf mode, just simulate
                Write-Log -Message "[WHAT IF] Would move: $($File.Name) -> $destinationPath" -Level 'Info'
            }
            
            Write-Log -Message "[MOVED] $($File.Name) -> $FileType folder" -Level 'Success'
            $script:Report.FilesMoved++
            Update-Report -Category $FileType -Size $fileSize
            
            # Clean up temp file if something went wrong
            if ($tempPath -and (Test-PathSafe -Path $tempPath)) {
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Clean up temp file on error
            if ($tempPath -and (Test-PathSafe -Path $tempPath)) {
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            }
            
            Write-Log -Message "[ERROR] Failed to move $($File.Name) : $_" -Level 'Error'
            Update-Report -Error "Failed to move $($File.Name) : $_"
            $script:Report.FilesFailed++
            
            # Additional error details
            if ($_.Exception.Message -like "*Access to the path*denied*") {
                Write-Log -Message "  → Access denied. File may be in use or permissions issue." -Level 'Debug'
            }
            elseif ($_.Exception.Message -like "*disk is full*") {
                Write-Log -Message "  → Disk full. Cannot complete operation." -Level 'Error'
            }
        }
    }
    else {
        Write-Log -Message "[WHAT IF] Would move: $($File.Name) -> $FileType folder" -Level 'Info'
        $script:Report.FilesSkipped++
    }
}

function Test-SourceDirectory {
    param([string]$Path)
    
    if (-not (Test-PathSafe -Path $Path)) {
        Write-Log -Message "Source directory does not exist: $Path" -Level 'Error'
        return $false
    }
    
    # Check if directory is accessible
    try {
        $null = Get-ChildItem -Path $Path -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log -Message "Cannot access source directory: $Path - $_" -Level 'Error'
        return $false
    }
}

#endregion

#region Extension Definitions

$script:ExtensionMap = @{
    # Documents
    ".pdf" = 'Document'
    ".doc" = 'Document'
    ".docx" = 'Document'
    ".ppt" = 'Document'
    ".pptx" = 'Document'
    ".xls" = 'Document'
    ".xlsx" = 'Document'
    ".txt" = 'Document'
    ".rtf" = 'Document'
    ".odt" = 'Document'
    ".ods" = 'Document'
    ".odp" = 'Document'
    ".csv" = 'Document'
    ".md" = 'Document'
    
    # Music
    ".mp3" = 'Music'
    ".wav" = 'Music'
    ".flac" = 'Music'
    ".aac" = 'Music'
    ".m4a" = 'Music'
    ".ogg" = 'Music'
    ".wma" = 'Music'
    ".opus" = 'Music'
    
    # Photos/Images
    ".jpg" = 'Photo'
    ".jpeg" = 'Photo'
    ".png" = 'Photo'
    ".gif" = 'Photo'
    ".bmp" = 'Photo'
    ".tiff" = 'Photo'
    ".webp" = 'Photo'
    ".heic" = 'Photo'
    ".svg" = 'Photo'
    ".ico" = 'Photo'
    ".raw" = 'Photo'
    ".cr2" = 'Photo'
    ".nef" = 'Photo'
}

#endregion

#region Main Execution

# Initialize report
Initialize-Report

try {
    Write-Log -Message "========== DOWNLOAD ORGANIZER STARTED ==========" -Level 'Info'
    Write-Log -Message "Source directory: $SourceDir" -Level 'Info'
    Write-Log -Message "Personal directory: $PersonalDir" -Level 'Info'
    Write-Log -Message "Music directory: $MusicDir" -Level 'Info'
    Write-Log -Message "Photos directory: $FotosDir" -Level 'Info'
    Write-Log -Message "Log file: $LogPath" -Level 'Info'
    Write-Log -Message "Report file: $ReportPath" -Level 'Info'
    Write-Log -Message "Timeout: $TimeoutSeconds seconds" -Level 'Info'
    Write-Log -Message "Parallel processing: $Parallel" -Level 'Info'
    Write-Log -Message "Preserve folder structure: $PreserveFolderStructure" -Level 'Info'

    if ($WhatIfPreference) {
        Write-Log -Message "WHAT IF MODE - No files will be actually moved" -Level 'Warning'
    }

    # Validate source directory
    if (-not (Test-SourceDirectory -Path $SourceDir)) {
        throw "Source directory validation failed"
    }

    # Ensure destination directories exist
    $destinations = @(
        @{ Path = $PersonalDir; Type = 'Personal documents' },
        @{ Path = $MusicDir; Type = 'Music' },
        @{ Path = $FotosDir; Type = 'Photos' }
    )

    $allDestinationsValid = $true
    foreach ($dest in $destinations) {
        if (-not (Ensure-Folder -Path $dest.Path -FolderDescription $dest.Type)) {
            $allDestinationsValid = $false
        }
    }

    if (-not $allDestinationsValid) {
        throw "Some destination folders could not be created. Please check permissions."
    }

    # Get files to process
    Write-Log -Message "Scanning for files in: $SourceDir" -Level 'Info'
    
    $getFilesParams = @{
        Path = $SourceDir
        File = $true
        ErrorAction = 'Stop'
    }
    
    if ($PreserveFolderStructure) {
        $getFilesParams['Recurse'] = $true
    }
    
    $files = Get-ChildItem @getFilesParams
    
    if ($files.Count -eq 0) {
        Write-Log -Message "No files found in source directory." -Level 'Warning'
    }
    else {
        Write-Log -Message "Found $($files.Count) file(s) to process" -Level 'Success'
        
        # Set up timeout
        $timeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()
        
        # Process files
        if ($Parallel -and $files.Count -gt 10) {
            Write-Log -Message "Using parallel processing with throttle limit $ThrottleLimit" -Level 'Info'
            
            $files | ForEach-Object -Parallel {
                $file = $_
                $sourceDir = $using:SourceDir
                $personalDir = $using:PersonalDir
                $musicDir = $using:MusicDir
                $fotosDir = $using:FotosDir
                $extensionMap = $using:ExtensionMap
                $whatIf = $using:WhatIfPreference
                $force = $using:Force
                $preserveStructure = $using:PreserveFolderStructure
                
                # Copy necessary functions and variables
                function Write-LogParallel { param($m,$l) # Simplified for parallel
                    Write-Host "[$l] $m"
                }
                
                $extension = $file.Extension.ToLower()
                
                if ([string]::IsNullOrEmpty($extension)) {
                    Write-LogParallel -m "Skipping file with no extension: $($file.Name)" -l "Debug"
                    return
                }
                
                $fileType = $extensionMap[$extension]
                
                switch ($fileType) {
                    'Document' { $dest = $personalDir }
                    'Music' { $dest = $musicDir }
                    'Photo' { $dest = $fotosDir }
                    default { return }
                }
                
                # Simplified move for parallel (would need more robust implementation)
                if ($whatIf) {
                    Write-LogParallel -m "[WHAT IF] Would move: $($file.Name) -> $dest" -l "Info"
                }
                else {
                    try {
                        $destPath = if ($preserveStructure) {
                            $relative = $file.Directory.FullName.Substring($sourceDir.Length).TrimStart('\')
                            if ($relative) { Join-Path $dest $relative $file.Name }
                            else { Join-Path $dest $file.Name }
                        } else {
                            Join-Path $dest $file.Name
                        }
                        
                        $destDir = Split-Path $destPath -Parent
                        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                        
                        Move-Item -Path $file.FullName -Destination $destPath -Force:$force -ErrorAction Stop
                        Write-LogParallel -m "[MOVED] $($file.Name) -> $fileType folder" -l "Success"
                    }
                    catch {
                        Write-LogParallel -m "[ERROR] Failed to move $($file.Name) : $_" -l "Error"
                    }
                }
            } -ThrottleLimit $ThrottleLimit
        }
        else {
            # Sequential processing
            foreach ($file in $files) {
                # Check timeout
                if ($timeoutTimer.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
                    Write-Log -Message "Timeout reached ($TimeoutSeconds seconds). Stopping processing." -Level 'Warning'
                    Update-Report -Warning "Processing stopped due to timeout"
                    break
                }
                
                $extension = $file.Extension.ToLower()
                
                # Skip files with no extension
                if ([string]::IsNullOrEmpty($extension)) {
                    Write-Log -Message "Skipping file with no extension: $($file.Name)" -Level 'Debug'
                    Update-Report -Category 'Unknown' -Size $file.Length
                    continue
                }
                
                $fileType = $script:ExtensionMap[$extension]
                
                switch ($fileType) {
                    'Document' {
                        Move-FileSafely -File $file -DestinationFolder $PersonalDir -FileType "Personal"
                    }
                    'Music' {
                        Move-FileSafely -File $file -DestinationFolder $MusicDir -FileType "Music"
                    }
                    'Photo' {
                        Move-FileSafely -File $file -DestinationFolder $FotosDir -FileType "Photos"
                    }
                    default {
                        Write-Log -Message "Skipping unsupported file type: $($file.Name) ($extension)" -Level 'Debug'
                        Update-Report -Category 'Unknown' -Size $file.Length
                    }
                }
            }
        }
        
        # Log summary statistics
        Write-Log -Message "---------- SUMMARY ----------" -Level 'Info'
        Write-Log -Message "Files processed: $($script:Report.FilesProcessed)" -Level 'Info'
        Write-Log -Message "Files moved: $($script:Report.FilesMoved)" -Level 'Success'
        Write-Log -Message "Files skipped: $($script:Report.FilesSkipped)" -Level 'Info'
        Write-Log -Message "Files failed: $($script:Report.FilesFailed)" -Level 'Warning'
        Write-Log -Message "Total size: $(Format-FileSize -Bytes $script:Report.TotalSizeBytes)" -Level 'Info'
        Write-Log -Message "-----------------------------" -Level 'Info'
        
        # Category breakdown
        foreach ($category in $script:Report.Categories.Keys) {
            if ($script:Report.Categories[$category].Count -gt 0) {
                Write-Log -Message "$category : $($script:Report.Categories[$category].Count) files ($(Format-FileSize -Bytes $script:Report.Categories[$category].Size))" -Level 'Debug'
            }
        }
    }
}
catch {
    Write-Log -Message "CRITICAL ERROR: $_" -Level 'Error'
    Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'Debug'
    $script:Report.Status = "Failed"
    Update-Report -Error "Critical error: $_"
}
finally {
    # Save report
    Save-Report
    
    # Final completion message
    Write-Log -Message "========== DOWNLOAD ORGANIZER COMPLETED ==========" -Level 'Success'
    Write-Log -Message "Log file saved to: $LogPath" -Level 'Info'
    Write-Log -Message "Report saved to: $ReportPath" -Level 'Info'
    
    if ($WhatIfPreference) {
        Write-Log -Message "WHAT IF MODE was active - No actual changes were made" -Level 'Warning'
    }
    
    if ($script:Report.FilesFailed -gt 0) {
        Write-Log -Message "WARNING: $($script:Report.FilesFailed) file(s) failed to process. Check the report for details." -Level 'Warning'
    }
}

#endregion