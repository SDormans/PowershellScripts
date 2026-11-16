<#
.SYNOPSIS
    Generates a size and file-extension report for a directory and all subdirectories.

.DESCRIPTION
    - Validates that the provided path exists and is accessible
    - Walks the directory tree once, accumulating per-directory size, file count and
      extension statistics (sizes are recursive: a directory includes its subdirectories)
    - Skips symbolic links / junctions to avoid cycles
    - Stops after -TimeoutSeconds and marks unscanned and affected directories
    - Outputs a formatted table (or objects with -PassThru)
    - Can export reports to CSV/JSON/HTML

.PARAMETER Path
    The root directory to analyze (required).

.PARAMETER MinSizeMB
    Minimum directory size to include in the report (default: 0 = all directories).

.PARAMETER SortBy
    Sort results by: 'Size', 'Name', or 'ExtensionCount' (default: 'Size').

.PARAMETER ExportFormat
    Export report to: 'CSV', 'JSON', 'HTML', or 'None' (default: 'None').

.PARAMETER OutputPath
    Directory where exported reports are saved (default: script directory).

.PARAMETER TimeoutSeconds
    Maximum seconds for the whole scan (default: 300 = 5 minutes).

.PARAMETER LogPath
    Path of the log file (default: script directory).

.PARAMETER PassThru
    Write the report objects to the pipeline instead of a formatted table.

.EXAMPLE
    .\DirectoryDistributionChecker.ps1 -Path "D:\Data" -ExportFormat HTML -MinSizeMB 100

.EXAMPLE
    .\DirectoryDistributionChecker.ps1 -Path "\\nas\shared" -PassThru | Where-Object FileCount -gt 1000
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({
        if (-not (Test-Path -LiteralPath $_ -PathType Container)) { throw "Directory does not exist: '$_'" }
        $true
    })]
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
    [string]$LogPath = (Join-Path $PSScriptRoot "DirectoryDistributionChecker_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"),

    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not $PSBoundParameters.ContainsKey('InformationAction')) {
    $InformationPreference = 'Continue'
}

Set-LogFile -Path $LogPath

#################################################
# DIRECTORY SCAN (single pass)
#################################################

function Initialize-DirectoryEntry {
    param([string]$FullName, [string]$Parent, [int]$Depth)

    return @{
        FullName   = $FullName
        Parent     = $Parent
        Depth      = $Depth
        Bytes      = [long]0
        Files      = 0
        Extensions = @{}
        Status     = 'Success'
    }
}

function Add-ExtensionStat {
    param([hashtable]$Target, [string]$Extension, [int]$Count, [long]$Size)

    $stat = $Target[$Extension]
    if (-not $stat) {
        $stat = @{ Count = 0; Size = [long]0 }
        $Target[$Extension] = $stat
    }
    $stat.Count += $Count
    $stat.Size += $Size
}

function Get-DirectoryTreeStat {
    <#
    .SYNOPSIS
        Walks the tree under RootPath once and returns one report object per directory.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutSeconds
    )

    $rootInfo = Get-Item -LiteralPath $RootPath -Force -ErrorAction Stop
    $entries = @{}
    $entries[$rootInfo.FullName] = Initialize-DirectoryEntry -FullName $rootInfo.FullName -Parent $null -Depth 0

    $pending = New-Object 'System.Collections.Generic.Stack[System.IO.DirectoryInfo]'
    $pending.Push($rootInfo)

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $scanned = 0
    $timedOut = $false

    while ($pending.Count -gt 0) {
        if ($timer.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
            $timedOut = $true
            break
        }

        $dir = $pending.Pop()
        $entry = $entries[$dir.FullName]
        $scanned++
        if ($scanned % 200 -eq 0) {
            Write-Progress -Activity 'Scanning directories' -Status "$scanned scanned, $($pending.Count) queued: $($dir.FullName)"
        }

        try {
            foreach ($item in $dir.EnumerateFileSystemInfos()) {
                if ($item -is [System.IO.DirectoryInfo]) {
                    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                        Write-Log -Message "Skipping link/junction: $($item.FullName)" -Level 'Verbose'
                        continue
                    }
                    $entries[$item.FullName] = Initialize-DirectoryEntry -FullName $item.FullName -Parent $dir.FullName -Depth ($entry.Depth + 1)
                    $pending.Push($item)
                }
                else {
                    $entry.Bytes += $item.Length
                    $entry.Files++
                    if ($item.Extension) {
                        Add-ExtensionStat -Target $entry.Extensions -Extension $item.Extension.ToLowerInvariant() -Count 1 -Size $item.Length
                    }
                }
            }
        }
        catch {
            $message = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
            $entry.Status = "Error: $message"
            Write-Log -Message "ERROR scanning $($dir.FullName) : $message" -Level 'Warning'
        }
    }
    Write-Progress -Activity 'Scanning directories' -Completed

    if ($timedOut) {
        Write-Log -Message "Scan timed out after $TimeoutSeconds seconds; $($pending.Count) director(ies) not scanned. Results are partial." -Level 'Warning'
        foreach ($unscanned in $pending) {
            $entries[$unscanned.FullName].Status = 'Not scanned (timeout)'
            $parent = $entries[$unscanned.FullName].Parent
            while ($parent) {
                if ($entries[$parent].Status -eq 'Success') { $entries[$parent].Status = 'Partial (timeout)' }
                $parent = $entries[$parent].Parent
            }
        }
    }

    # Roll up totals from the deepest directories towards the root.
    foreach ($entry in ($entries.Values | Sort-Object { $_.Depth } -Descending)) {
        if (-not $entry.Parent) { continue }
        $parentEntry = $entries[$entry.Parent]
        $parentEntry.Bytes += $entry.Bytes
        $parentEntry.Files += $entry.Files
        foreach ($ext in $entry.Extensions.Keys) {
            Add-ExtensionStat -Target $parentEntry.Extensions -Extension $ext -Count $entry.Extensions[$ext].Count -Size $entry.Extensions[$ext].Size
        }
    }

    foreach ($entry in $entries.Values) {
        $extensionStats = @($entry.Extensions.GetEnumerator() |
            Sort-Object -Property @{ Expression = { $_.Value.Count }; Descending = $true }, @{ Expression = { $_.Key }; Descending = $false })
        $topExtensions = $extensionStats | Select-Object -First 5 | ForEach-Object { "$($_.Key) ($($_.Value.Count))" }

        $status = $entry.Status
        if ($status -eq 'Success' -and $entry.Files -eq 0) { $status = 'No files' }

        [PSCustomObject]@{
            Directory      = $entry.FullName
            SizeBytes      = $entry.Bytes
            SizeFormatted  = Format-FileSize -Bytes $entry.Bytes
            SizeMB         = [Math]::Round($entry.Bytes / 1MB, 2)
            FileCount      = $entry.Files
            ExtensionCount = $extensionStats.Count
            Extensions     = if ($extensionStats.Count) { ($extensionStats | ForEach-Object { $_.Key }) -join ', ' } else { '-' }
            TopExtensions  = if ($extensionStats.Count) { $topExtensions -join '; ' } else { '-' }
            Status         = $status
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
        [ValidateSet('CSV', 'JSON', 'HTML')]
        [string]$Format,

        [Parameter(Mandatory = $true)]
        [string]$OutputDir
    )

    if (-not (Test-PathSafe -Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseName = "DirectoryDistribution_Report_$timestamp"

    switch ($Format) {
        'CSV' {
            $filePath = Join-Path $OutputDir "$baseName.csv"
            try {
                $Report | Export-Csv -LiteralPath $filePath -NoTypeInformation -Encoding UTF8
                Write-Log -Message "Report exported to CSV: $filePath" -Level 'Success'
            }
            catch {
                Write-Log -Message "Failed to export CSV: $_" -Level 'Error'
            }
        }
        'JSON' {
            $filePath = Join-Path $OutputDir "$baseName.json"
            try {
                $Report | ConvertTo-Json -Depth 3 | Out-File -LiteralPath $filePath -Encoding UTF8
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
    <p>Root Path: $([System.Net.WebUtility]::HtmlEncode($Path))</p>
    <hr>
"@

                $htmlTable = $Report | Select-Object Directory, SizeFormatted, FileCount, ExtensionCount, TopExtensions, Status |
                    ConvertTo-Html -Fragment

                $htmlFooter = @"
    <hr>
    <p>Total directories scanned: $($Report.Count)</p>
    <p>Total size: $(Format-FileSize -Bytes ($Report | Measure-Object SizeBytes -Sum).Sum)</p>
</body>
</html>
"@

                $htmlHeader + ($htmlTable -join [Environment]::NewLine) + $htmlFooter | Out-File -LiteralPath $filePath -Encoding UTF8
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

Write-Log -Message '========== DIRECTORY DISTRIBUTION CHECKER STARTED =========='
Write-Log -Message "Root path: $Path"
Write-Log -Message "Minimum size filter: $MinSizeMB MB"
Write-Log -Message "Scan timeout: $TimeoutSeconds seconds"
Write-Log -Message "Log file: $LogPath"

if (-not $PSCmdlet.ShouldProcess($Path, 'Scan directory tree')) {
    $topLevel = @(Get-ChildItem -LiteralPath $Path -Directory -Force -ErrorAction SilentlyContinue)
    Write-Log -Message "WHAT IF: Would analyze $Path ($($topLevel.Count) top-level subdirectories)"
    $topLevel | Select-Object -First 5 | ForEach-Object { Write-Log -Message "  -> $($_.FullName)" }
    return
}

#################################################
# REPORT GENERATION
#################################################

Write-Log -Message 'Scanning directory tree...'
$startTime = Get-Date

try {
    $report = @(Get-DirectoryTreeStat -RootPath $Path -TimeoutSeconds $TimeoutSeconds)
}
catch {
    Write-Log -Message "Critical error scanning $Path : $_" -Level 'Error'
    throw
}

$totalDuration = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
$directoryCount = $report.Count
$failedDirs = @($report | Where-Object { $_.Status -like 'Error*' -or $_.Status -like 'Not scanned*' }).Count
$rootEntry = $report | Where-Object { $_.Directory -eq (Get-Item -LiteralPath $Path -Force).FullName } | Select-Object -First 1

#################################################
# FILTER REPORT
#################################################

if ($MinSizeMB -gt 0) {
    $report = @($report | Where-Object { $_.SizeMB -ge $MinSizeMB })
    Write-Log -Message "Filtered from $directoryCount to $($report.Count) directories (size >= $MinSizeMB MB)"
}

switch ($SortBy) {
    'Size' { $report = @($report | Sort-Object SizeBytes -Descending) }
    'Name' { $report = @($report | Sort-Object Directory) }
    'ExtensionCount' { $report = @($report | Sort-Object ExtensionCount -Descending) }
}

#################################################
# OUTPUT
#################################################

Write-Log -Message '========== REPORT SUMMARY ==========' -Level 'Success'
Write-Log -Message "Total scan duration: $totalDuration seconds"
Write-Log -Message "Directories analyzed: $directoryCount"
Write-Log -Message "Directories in report: $($report.Count)"
Write-Log -Message "Directories with errors / not scanned: $failedDirs"
if ($rootEntry) {
    Write-Log -Message "Total size: $($rootEntry.SizeFormatted) in $($rootEntry.FileCount) files"
}
Write-Log -Message '====================================' -Level 'Success'

if ($PassThru) {
    Write-Output $report
}
else {
    $report | Select-Object Directory,
        @{ N = 'Size'; E = { $_.SizeFormatted } },
        @{ N = 'Files'; E = { $_.FileCount } },
        @{ N = 'Exts'; E = { $_.ExtensionCount } },
        @{ N = 'Top Extensions'; E = { $_.TopExtensions } },
        @{ N = 'Status'; E = { $_.Status } } |
        Format-Table -AutoSize -Wrap
}

#################################################
# EXPORT
#################################################

if ($ExportFormat -ne 'None' -and $report.Count -gt 0) {
    Export-Report -Report $report -Format $ExportFormat -OutputDir $OutputPath
}

#################################################
# COMPLETION
#################################################

$largestDir = $report | Where-Object { -not $rootEntry -or $_.Directory -ne $rootEntry.Directory } | Sort-Object SizeBytes -Descending | Select-Object -First 1
if ($largestDir) {
    Write-Log -Message "Largest subdirectory: $($largestDir.Directory) ($($largestDir.SizeFormatted))"
}

Write-Log -Message '========== DIRECTORY DISTRIBUTION CHECKER COMPLETED ==========' -Level 'Success'
Write-Log -Message "Log file saved to: $LogPath"
