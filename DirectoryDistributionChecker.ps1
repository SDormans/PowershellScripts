<#
.SYNOPSIS
    Generates a comprehensive size and file-extension report for a directory and all subdirectories.

.DESCRIPTION
    - Validates that the provided path exists and is accessible
    - Recursively scans directories with timeout protection
    - Calculates total size in MB/GB with configurable precision
    - Lists unique file extensions per directory
    - Outputs formatted tables with multiple sorting options
    - Can export reports to CSV/JSON/HTML

.PARAMETER Path
    The root directory to analyze (required)

.PARAMETER MinSizeMB
    Minimum directory size to include in report (default: 0 = all directories)

.PARAMETER SortBy
    Sort results by: 'Size', 'Name', or 'ExtensionCount' (default: 'Size')

.PARAMETER ExportFormat
    Export report to: 'CSV', 'JSON', 'HTML', or 'None' (default: 'None')

.PARAMETER OutputPath
    Path where to save exported report (default: current directory)

.PARAMETER TimeoutSeconds
    Maximum seconds to spend scanning (default: 300 = 5 minutes)

.PARAMETER WhatIf
    Shows what would happen without actually scanning

.PARAMETER LogPath
    Optional path to write a log file (default: script directory)
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({Test-Path -Path $_}, ErrorMessage = "Path does not exist: '{0}'")]
    [string]$Path,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [double]::MaxValue)]
    [double]$MinSizeMB = 0,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Size', 'Name', 'ExtensionCount')]
    [string]$SortBy = 'Size',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('None', 'CSV', 'JSON', 'HTML')]
    [string]$ExportFormat = 'None',
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = $PSScriptRoot,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 3600)]
    [int]$TimeoutSeconds = 300,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$PSScriptRoot\DirectoryReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

#################################################
# LOGGING FUNCTION
#################################################
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug', 'Verbose')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console with colors
    switch ($Level) {
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
        'Verbose' { 
            if ($VerbosePreference -ne 'SilentlyContinue') {
                Write-Host $logMessage -ForegroundColor Cyan
            }
        }
        'Debug' { 
            if ($DebugPreference -ne 'SilentlyContinue') {
                Write-Host $logMessage -ForegroundColor Gray
            }
        }
        default { Write-Host $logMessage -ForegroundColor Gray }
    }
    
    # Write to log file
    try {
        Add-Content -Path $LogPath -Value $logMessage -ErrorAction Stop
    }
    catch {
        Write-Host "[WARNING] Could not write to log file: $_" -ForegroundColor Yellow
    }
}

#################################################
# PROGRESS DISPLAY FUNCTION
#################################################
function Show-Progress {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$Current,
        [int]$Total
    )
    
    if ($Total -gt 0) {
        $percent = [Math]::Round(($Current / $Total) * 100, 0)
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $percent
    }
}

#################################################
# SIZE FORMATTING FUNCTION
#################################################
function Format-FileSize {
    param([double]$SizeInBytes)
    
    if ($SizeInBytes -ge 1TB) {
        return "{0:N2} TB" -f ($SizeInBytes / 1TB)
    }
    elseif ($SizeInBytes -ge 1GB) {
        return "{0:N2} GB" -f ($SizeInBytes / 1GB)
    }
    elseif ($SizeInBytes -ge 1MB) {
        return "{0:N2} MB" -f ($SizeInBytes / 1MB)
    }
    elseif ($SizeInBytes -ge 1KB) {
        return "{0:N2} KB" -f ($SizeInBytes / 1KB)
    }
    else {
        return "$SizeInBytes B"
    }
}

#################################################
# DIRECTORY REPORT FUNCTION (Verbeterd)
#################################################
function Get-DirectoryReport {
    <#
    .SYNOPSIS
        Creates a comprehensive report object for a single directory.
    #>

    param (
        [Parameter(Mandatory = $true)]
        [string]$Directory,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 30,
        
        [Parameter(Mandatory = $false)]
        [int]$CurrentIndex = 0,
        
        [Parameter(Mandatory = $false)]
        [int]$TotalDirectories = 1
    )

    # Update progress
    Show-Progress -Activity "Scanning directories" -Status $Directory -Current $CurrentIndex -Total $TotalDirectories

    # Start timing for timeout
    $startTime = Get-Date
    
    try {
        Write-Log -Message "Scanning: $Directory" -Level 'Verbose'
        
        # Use a timeout mechanism for Get-ChildItem (can hang on network drives)
        $files = $null
        $job = Start-Job -ScriptBlock {
            param($path)
            Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue -Force
        } -ArgumentList $Directory
        
        # Wait for job with timeout
        $job | Wait-Job -Timeout $TimeoutSeconds | Out-Null
        
        if ($job.State -eq 'Running') {
            # Job timed out
            Stop-Job $job
            Remove-Job $job -Force
            throw "Directory scan timed out after $TimeoutSeconds seconds"
        }
        else {
            # Job completed
            $files = Receive-Job $job
            Remove-Job $job
        }

        # Handle empty or inaccessible directories
        if (-not $files) {
            Write-Log -Message "  → No files found or inaccessible" -Level 'Debug'
            return [PSCustomObject]@{
                Directory        = $Directory
                SizeBytes        = 0
                SizeFormatted    = "0 B"
                SizeMB           = 0
                FileCount        = 0
                ExtensionCount   = 0
                Extensions       = "-"
                TopExtensions    = "-"
                ScanDuration     = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
                Status           = "No files"
            }
        }

        # Calculate total size
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
        $sizeMB = [Math]::Round($totalSize / 1MB, 2)
        
        # Collect extension statistics
        $extensionStats = $files |
            Where-Object { $_.Extension } |
            Group-Object { $_.Extension.ToLower() } |
            Select-Object @{N='Extension';E={$_.Name}}, 
                         @{N='Count';E={$_.Count}},
                         @{N='TotalSize';E={($_.Group | Measure-Object Length -Sum).Sum}} |
            Sort-Object Count -Descending
        
        $extensions = $extensionStats | Select-Object -ExpandProperty Extension
        $extensionCount = $extensions.Count
        
        # Get top 5 extensions by count
        $topExtensions = $extensionStats | Select-Object -First 5 | ForEach-Object {
            "$($_.Extension) ($($_.Count))"
        }
        
        # Calculate scan duration
        $duration = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
        Write-Log -Message "  → Found $($files.Count) files, $extensionCount extensions, $sizeMB MB in ${duration}s" -Level 'Verbose'

        # Return structured report object
        [PSCustomObject]@{
            Directory        = $Directory
            SizeBytes        = $totalSize
            SizeFormatted    = Format-FileSize -SizeInBytes $totalSize
            SizeMB           = $sizeMB
            FileCount        = $files.Count
            ExtensionCount   = $extensionCount
            Extensions       = ($extensions -join ", ")
            TopExtensions    = ($topExtensions -join "; ")
            ScanDuration     = $duration
            Status           = "Success"
        }
    }
    catch {
        Write-Log -Message "ERROR scanning $Directory : $_" -Level 'Warning'
        
        return [PSCustomObject]@{
            Directory        = $Directory
            SizeBytes        = 0
            SizeFormatted    = "0 B"
            SizeMB           = 0
            FileCount        = 0
            ExtensionCount   = 0
            Extensions       = "-"
            TopExtensions    = "-"
            ScanDuration     = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
            Status           = "Error: $_"
        }
    }
}

#################################################
# EXPORT FUNCTIONS
#################################################
function Export-Report {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Report,
        
        [Parameter(Mandatory = $true)]
        [string]$Format,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputDir
    )
    
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseName = "DirectoryReport_$timestamp"
    
    switch ($Format) {
        'CSV' {
            $filePath = Join-Path $OutputDir "$baseName.csv"
            try {
                $report | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
                Write-Log -Message "Report exported to CSV: $filePath" -Level 'Success'
            }
            catch {
                Write-Log -Message "Failed to export CSV: $_" -Level 'Error'
            }
        }
        'JSON' {
            $filePath = Join-Path $OutputDir "$baseName.json"
            try {
                $report | ConvertTo-Json -Depth 3 | Out-File -FilePath $filePath -Encoding UTF8
                Write-Log -Message "Report exported to JSON: $filePath" -Level 'Success'
            }
            catch {
                Write-Log -Message "Failed to export JSON: $_" -Level 'Error'
            }
        }
        'HTML' {
            $filePath = Join-Path $OutputDir "$baseName.html"
            try {
                $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Directory Report - $timestamp</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        table { border-collapse: collapse; width: 100%; }
        th { background-color: #4CAF50; color: white; padding: 8px; text-align: left; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .warning { background-color: #ff9800; color: white; padding: 10px; }
        .success { background-color: #4CAF50; color: white; padding: 10px; }
    </style>
</head>
<body>
    <h1>Directory Size Report</h1>
    <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <p>Root Path: $Path</p>
    <hr>
"@
                
                $htmlTable = $report | Select-Object Directory, SizeFormatted, FileCount, ExtensionCount, TopExtensions, Status |
                    ConvertTo-Html -Fragment
                
                $htmlFooter = @"
    <hr>
    <p>Total directories scanned: $($report.Count)</p>
    <p>Total size: $(Format-FileSize -SizeInBytes ($report | Measure-Object SizeBytes -Sum).Sum)</p>
</body>
</html>
"@
                
                $htmlHeader + $htmlTable + $htmlFooter | Out-File -FilePath $filePath -Encoding UTF8
                Write-Log -Message "Report exported to HTML: $filePath" -Level 'Success'
            }
            catch {
                Write-Log -Message "Failed to export HTML: $_" -Level 'Error'
            }
        }
    }
}

#################################################
# MAIN SCRIPT EXECUTION
#################################################

Write-Log -Message "========== DIRECTORY REPORT GENERATOR STARTED ==========" -Level 'Info'
Write-Log -Message "Root path: $Path" -Level 'Info'
Write-Log -Message "Minimum size filter: $MinSizeMB MB" -Level 'Info'
Write-Log -Message "Timeout per directory: $TimeoutSeconds seconds" -Level 'Info'
Write-Log -Message "Log file: $LogPath" -Level 'Info'

if ($WhatIfPreference) {
    Write-Log -Message "WHAT IF MODE - Analysis will be simulated" -Level 'Warning'
}

#################################################
# DIRECTORY COLLECTION
#################################################

Write-Log -Message "Collecting directories..." -Level 'Info'

try {
    # Get all subdirectories with error handling
    $directories = try {
        Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction Stop -Force
    }
    catch {
        Write-Log -Message "Error during directory enumeration: $_" -Level 'Warning'
        Write-Log -Message "Continuing with accessible directories only..." -Level 'Info'
        Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue -Force
    }
    
    $directories = @($Path) + ($directories | Select-Object -ExpandProperty FullName)
    Write-Log -Message "Found $($directories.Count) directories to analyze" -Level 'Success'
}
catch {
    Write-Log -Message "Critical error collecting directories: $_" -Level 'Error'
    exit 1
}

#################################################
# WHAT IF CHECK
#################################################

if ($WhatIfPreference) {
    Write-Log -Message "WHAT IF: Would analyze $($directories.Count) directories" -Level 'Info'
    Write-Log -Message "First 5 directories:" -Level 'Info'
    $directories | Select-Object -First 5 | ForEach-Object {
        Write-Log -Message "  → $_" -Level 'Info'
    }
    exit 0
}

#################################################
# REPORT GENERATION
#################################################

Write-Log -Message "Generating directory reports..." -Level 'Info'

$report = @()
$index = 0
$total = $directories.Count
$failedDirs = 0
$startTime = Get-Date

foreach ($dir in $directories) {
    $index++
    
    $reportEntry = Get-DirectoryReport -Directory $dir -TimeoutSeconds $TimeoutSeconds -CurrentIndex $index -TotalDirectories $total
    $report += $reportEntry
    
    if ($reportEntry.Status -ne 'Success') {
        $failedDirs++
    }
}

$totalDuration = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)

#################################################
# FILTER REPORT
#################################################

if ($MinSizeMB -gt 0) {
    $originalCount = $report.Count
    $report = $report | Where-Object { $_.SizeMB -ge $MinSizeMB }
    Write-Log -Message "Filtered from $originalCount to $($report.Count) directories (size >= $MinSizeMB MB)" -Level 'Info'
}

#################################################
# OUTPUT
#################################################

Write-Log -Message "========== REPORT SUMMARY ==========" -Level 'Success'
Write-Log -Message "Total scan duration: $totalDuration seconds" -Level 'Info'
Write-Log -Message "Directories analyzed: $($directories.Count)" -Level 'Info'
Write-Log -Message "Directories with data: $($report.Count)" -Level 'Info'
Write-Log -Message "Failed/scanned with errors: $failedDirs" -Level 'Info'
Write-Log -Message "Total size: $(Format-FileSize -SizeInBytes ($report | Measure-Object SizeBytes -Sum).Sum)" -Level 'Info'
Write-Log -Message "====================================" -Level 'Success'

# Determine sort order
switch ($SortBy) {
    'Size' { $report = $report | Sort-Object SizeBytes -Descending }
    'Name' { $report = $report | Sort-Object Directory }
    'ExtensionCount' { $report = $report | Sort-Object ExtensionCount -Descending }
}

# Display results in a readable table
$report | Select-Object Directory, 
    @{N='Size';E={$_.SizeFormatted}},
    @{N='Files';E={$_.FileCount}},
    @{N='Exts';E={$_.ExtensionCount}},
    @{N='Top Extensions';E={$_.TopExtensions}},
    @{N='Status';E={$_.Status}} |
Format-Table -AutoSize -Wrap

#################################################
# EXPORT
#################################################

if ($ExportFormat -ne 'None') {
    Export-Report -Report $report -Format $ExportFormat -OutputDir $OutputPath
}

#################################################
# COMPLETION
#################################################

Write-Log -Message "========== DIRECTORY REPORT GENERATOR COMPLETED ==========" -Level 'Success'
Write-Log -Message "Log file saved to: $LogPath" -Level 'Info'

# Additional insights
$largestDir = $report | Sort-Object SizeBytes -Descending | Select-Object -First 1
if ($largestDir) {
    Write-Log -Message "Largest directory: $($largestDir.Directory) ($($largestDir.SizeFormatted))" -Level 'Info'
}

$mostFiles = $report | Sort-Object FileCount -Descending | Select-Object -First 1
if ($mostFiles -and $mostFiles.FileCount -gt 0) {
    Write-Log -Message "Directory with most files: $($mostFiles.Directory) ($($mostFiles.FileCount) files)" -Level 'Info'
}
