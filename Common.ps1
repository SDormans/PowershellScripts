<#
.SYNOPSIS
    Shared helpers for the PowerShell file organization scripts.

.DESCRIPTION
    Dot-source this file (". $PSScriptRoot\Common.ps1") instead of importing it as a module so
    the functions run in the calling script's scope and honour its -WhatIf, -Confirm, -Verbose
    and -InformationAction preferences. Module functions do not inherit those preferences from
    a calling script, which would make -WhatIf unsafe for the move operations below.

    Compatible with Windows PowerShell 5.1 and PowerShell 7+.
#>

#region Logging

function Set-LogFile {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Only sets the in-memory log file path.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Path
    )

    $script:CommonLogFile = $Path

    if ($Path) {
        $directory = Split-Path -Path $Path -Parent
        if ($directory -and -not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force -WhatIf:$false -Confirm:$false | Out-Null
        }
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped message to the matching PowerShell stream and to the log file
        configured with Set-LogFile.
    .DESCRIPTION
        Info/Success -> Information stream, Warning -> Warning stream, Error -> Error stream
        (non-terminating), Verbose -> Verbose stream, Debug -> Debug stream.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Write-Log is not a built-in cmdlet on the supported platforms; the dot-sourced function takes precedence.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug', 'Verbose')]
        [string]$Level = 'Info',

        [Parameter(Mandatory = $false)]
        [switch]$NoFile
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $logMessage = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'Warning' { Write-Warning -Message $logMessage }
        'Error' {
            # Never let a log call turn into a terminating error inside a catch block.
            $errorAction = if ($ErrorActionPreference -eq 'Stop') { 'Continue' } else { $ErrorActionPreference }
            Write-Error -Message $logMessage -ErrorAction $errorAction
        }
        'Verbose' { Write-Verbose -Message $logMessage }
        'Debug' { Write-Debug -Message $logMessage }
        default { Write-Information -MessageData $logMessage -Tags $Level }
    }

    # Verbose/Debug entries only go to the file when that stream is enabled.
    $skipFile = ($Level -eq 'Verbose' -and $VerbosePreference -eq 'SilentlyContinue') -or
                ($Level -eq 'Debug' -and $DebugPreference -eq 'SilentlyContinue')

    if ($NoFile -or $skipFile -or -not $script:CommonLogFile) {
        return
    }

    $maxRetries = 3
    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            Add-Content -LiteralPath $script:CommonLogFile -Value $logMessage -ErrorAction Stop -WhatIf:$false -Confirm:$false
            return
        }
        catch {
            if ($attempt -eq $maxRetries) {
                Write-Warning -Message "Failed to write to log file '$($script:CommonLogFile)' after $maxRetries attempts: $_"
            }
            else {
                Start-Sleep -Milliseconds 100
            }
        }
    }
}

#endregion

#region Configuration

function Import-ScriptConfig {
    <#
    .SYNOPSIS
        Loads a JSON config file and returns the settings for one script as a hashtable.
    .DESCRIPTION
        If the file contains a top-level property named after -Section, that object is used;
        otherwise the top-level object is used. Returns an empty hashtable when -Path is empty.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$Section
    )

    $settings = @{}
    if (-not $Path) {
        return $settings
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Configuration file not found: $Path"
    }

    try {
        $json = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Failed to parse configuration file '$Path': $_"
    }

    if ($Section -and $json.PSObject.Properties.Name -contains $Section) {
        $json = $json.$Section
    }

    foreach ($property in $json.PSObject.Properties) {
        $settings[$property.Name] = $property.Value
    }

    return $settings
}

function Resolve-Setting {
    <#
    .SYNOPSIS
        Resolves a setting value: explicit parameter > config file > default.
    .DESCRIPTION
        Throws when -Required is set and no non-empty value is found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary]$BoundParameters = @{},

        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary]$Config = @{},

        [Parameter(Mandatory = $false)]
        [object]$Default = $null,

        # Legacy config keys accepted for this setting (parameter aliases do not apply to JSON keys).
        [Parameter(Mandatory = $false)]
        [string[]]$Alias = @(),

        [Parameter(Mandatory = $false)]
        [switch]$Required
    )

    $value = $Default
    if (@($BoundParameters.Keys) -contains $Name) {
        $value = $BoundParameters[$Name]
    }
    else {
        foreach ($key in @($Name) + $Alias) {
            if ((@($Config.Keys) -contains $key) -and $null -ne $Config[$key] -and "$($Config[$key])" -ne '') {
                $value = $Config[$key]
                break
            }
        }
    }

    if ($Required -and ($null -eq $value -or "$value" -eq '')) {
        throw "Setting '$Name' is required. Pass -$Name or set '$Name' in the configuration file (see config.example.json)."
    }

    if ($value -is [switch]) {
        return $value.IsPresent
    }
    return $value
}

#endregion

#region File system helpers

function Format-FileSize {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias('SizeInBytes')]
        [double]$Bytes
    )

    if ($Bytes -ge 1TB) { return '{0:N2} TB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Test-PathSafe {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    try {
        return [bool](Test-Path -LiteralPath $Path -ErrorAction Stop)
    }
    catch {
        return $false
    }
}

function ConvertTo-FullPath {
    <#
    .SYNOPSIS
        Resolves a (possibly relative or not yet existing) path against the current PowerShell
        location. [System.IO.Path]::GetFullPath uses the process directory, which Set-Location
        does not change.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path
    )

    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function New-FolderIfMissing {
    <#
    .SYNOPSIS
        Creates a folder (and missing parents) if it does not exist. Returns $true when the
        folder exists afterwards, or would exist when running with -WhatIf.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$Description = 'folder',

        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary]$Report
    )

    if (Test-PathSafe -Path $Path) {
        return $true
    }

    if (-not $PSCmdlet.ShouldProcess($Path, "Create $Description")) {
        Write-Log -Message "[WHAT IF] Would create $Description : $Path"
        return $true
    }

    try {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop -WhatIf:$false -Confirm:$false | Out-Null
        Write-Log -Message "Created $Description : $Path" -Level 'Success'
        return $true
    }
    catch {
        Write-Log -Message "Failed to create $Description $Path : $_" -Level 'Error'
        if ($Report) {
            Update-Report -Report $Report -ErrorMessage "Failed to create folder $Path : $_" -NoLog
        }
        return $false
    }
}

function Test-SameVolume {
    <#
    .SYNOPSIS
        Returns $true when both paths share the same root (drive letter, UNC share or '/').
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$OtherPath
    )

    $root = [System.IO.Path]::GetPathRoot((ConvertTo-FullPath -Path $Path))
    $otherRoot = [System.IO.Path]::GetPathRoot((ConvertTo-FullPath -Path $OtherPath))
    return [string]::Equals($root, $otherRoot, [System.StringComparison]::OrdinalIgnoreCase)
}

function Move-FileSafely {
    <#
    .SYNOPSIS
        Moves a file to an exact destination path and records the outcome in a report section.
    .DESCRIPTION
        Same-volume moves use Move-Item (a rename, no data copied). Cross-volume moves copy to a
        temporary file next to the destination, verify the size, rename it into place and only
        then delete the source.

        Report counters touched on -Section: Processed, Moved, Skipped, Failed and Size.
        Returns one of: Moved, Skipped, Failed, WhatIf, Missing.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary]$Report,

        [Parameter(Mandatory = $false)]
        [string]$Section,

        [Parameter(Mandatory = $false)]
        [string]$Description,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    if (-not $Description) { $Description = if ($Section) { "$Section folder" } else { 'destination' } }
    $hasReport = $null -ne $Report

    if (-not (Test-PathSafe -Path $File.FullName)) {
        Write-Log -Message "Source file no longer exists: $($File.FullName)" -Level 'Warning'
        if ($hasReport) { Update-Report -Report $Report -Section $Section -Metric 'Skipped' -WarningMessage "Source file disappeared: $($File.FullName)" -NoLog }
        return 'Missing'
    }

    if ($hasReport) { Update-Report -Report $Report -Section $Section -Metric 'Processed' }

    $destinationDir = Split-Path -Path $DestinationPath -Parent
    if (-not (New-FolderIfMissing -Path $destinationDir -Description 'destination folder' -Report $Report)) {
        if ($hasReport) { Update-Report -Report $Report -Section $Section -Metric 'Failed' }
        return 'Failed'
    }

    if (Test-PathSafe -Path $DestinationPath) {
        if (-not $Force) {
            Write-Log -Message "Skipping $($File.Name) - already exists in $Description" -Level 'Warning'
            if ($hasReport) { Update-Report -Report $Report -Section $Section -Metric 'Skipped' -WarningMessage "File already exists: $DestinationPath" -NoLog }
            return 'Skipped'
        }
        Write-Log -Message "Overwriting existing file: $DestinationPath" -Level 'Warning'
    }

    if (-not $PSCmdlet.ShouldProcess($File.FullName, "Move to $DestinationPath")) {
        Write-Log -Message "[WHAT IF] Would move: $($File.FullName) -> $DestinationPath"
        if ($hasReport) { Update-Report -Report $Report -Section $Section -Metric 'Skipped' }
        return 'WhatIf'
    }

    $fileSize = $File.Length
    $tempPath = $null
    try {
        if (Test-SameVolume -Path $File.FullName -OtherPath $DestinationPath) {
            Move-Item -LiteralPath $File.FullName -Destination $DestinationPath -Force:$Force -ErrorAction Stop -WhatIf:$false -Confirm:$false
        }
        else {
            $tempPath = Join-Path $destinationDir "~$($File.Name).$([guid]::NewGuid().ToString('N')).tmp"
            Copy-Item -LiteralPath $File.FullName -Destination $tempPath -ErrorAction Stop -WhatIf:$false -Confirm:$false

            $copiedSize = (Get-Item -LiteralPath $tempPath -ErrorAction Stop).Length
            if ($copiedSize -ne $fileSize) {
                throw "Copy verification failed: expected $fileSize bytes, got $copiedSize bytes"
            }

            Move-Item -LiteralPath $tempPath -Destination $DestinationPath -Force:$Force -ErrorAction Stop -WhatIf:$false -Confirm:$false
            $tempPath = $null
            Remove-Item -LiteralPath $File.FullName -Force -ErrorAction Stop -WhatIf:$false -Confirm:$false
        }

        Write-Log -Message "[MOVED] $($File.FullName) -> $DestinationPath" -Level 'Success'
        if ($hasReport) { Update-Report -Report $Report -Section $Section -Metric 'Moved' -Size $fileSize }
        return 'Moved'
    }
    catch {
        $reason = "$_"
        if ($_.Exception.Message -like '*Access to the path*denied*') {
            $reason += ' (access denied: file may be in use or permissions are missing)'
        }
        elseif ($_.Exception.Message -like '*disk*full*' -or $_.Exception.Message -like '*not enough space*') {
            $reason += ' (destination disk is full)'
        }

        Write-Log -Message "Failed to move $($File.FullName) : $reason" -Level 'Error'
        if ($hasReport) { Update-Report -Report $Report -Section $Section -Metric 'Failed' -ErrorMessage "Failed to move $($File.FullName) : $reason" -NoLog }
        return 'Failed'
    }
    finally {
        if ($tempPath -and (Test-PathSafe -Path $tempPath)) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
        }
    }
}

#endregion

#region Reporting

function Initialize-Report {
    <#
    .SYNOPSIS
        Creates a report hashtable.
    .PARAMETER Sections
        Names of report sections. Each gets Processed/Moved/Skipped/Failed/Size counters.
    .PARAMETER ExtraCounters
        Optional per-section extra counters, e.g. @{ Music = @('Duplicates') }.
    .PARAMETER Metadata
        Extra top-level properties (source directory, script name, ...).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Sections,

        [Parameter(Mandatory = $false)]
        [hashtable]$ExtraCounters = @{},

        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary]$Metadata = @{}
    )

    $report = [ordered]@{
        StartTime    = Get-Date
        ComputerName = [System.Environment]::MachineName
        UserName     = [System.Environment]::UserName
        Status       = 'Running'
    }

    foreach ($key in $Metadata.Keys) {
        $report[$key] = $Metadata[$key]
    }

    $sectionTable = [ordered]@{}
    foreach ($name in $Sections) {
        $counters = [ordered]@{ Processed = 0; Moved = 0; Skipped = 0; Failed = 0; Size = [long]0 }
        if ($ExtraCounters.ContainsKey($name)) {
            foreach ($counter in $ExtraCounters[$name]) { $counters[$counter] = 0 }
        }
        $sectionTable[$name] = $counters
    }

    $report.Sections = $sectionTable
    $report.Errors = New-Object System.Collections.Generic.List[object]
    $report.Warnings = New-Object System.Collections.Generic.List[object]
    return $report
}

function Update-Report {
    <#
    .SYNOPSIS
        Increments a section counter and/or records an error or warning message.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Only updates an in-memory report.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Report,

        [Parameter(Mandatory = $false)]
        [string]$Section,

        [Parameter(Mandatory = $false)]
        [string]$Metric,

        [Parameter(Mandatory = $false)]
        [int]$Increment = 1,

        [Parameter(Mandatory = $false)]
        [long]$Size = 0,

        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage,

        [Parameter(Mandatory = $false)]
        [string]$WarningMessage,

        # Record messages without also writing them to the log (caller already logged them).
        [Parameter(Mandatory = $false)]
        [switch]$NoLog
    )

    if ($Section -and $Report.Sections.Contains($Section)) {
        $counters = $Report.Sections[$Section]
        if ($Metric -and $counters.Contains($Metric)) {
            $counters[$Metric] += $Increment
        }
        if ($Size -gt 0 -and $counters.Contains('Size')) {
            $counters['Size'] += $Size
        }
    }

    if ($ErrorMessage) {
        $Report.Errors.Add([ordered]@{ Time = Get-Date; Message = $ErrorMessage })
        if (-not $NoLog) { Write-Log -Message $ErrorMessage -Level 'Error' }
    }

    if ($WarningMessage) {
        $Report.Warnings.Add([ordered]@{ Time = Get-Date; Message = $WarningMessage })
        if (-not $NoLog) { Write-Log -Message $WarningMessage -Level 'Warning' }
    }
}

function Get-ReportSummary {
    <#
    .SYNOPSIS
        Totals the counters of all report sections.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Report
    )

    $totals = @{ Processed = 0; Moved = 0; Skipped = 0; Failed = 0; Size = [long]0 }
    foreach ($section in $Report.Sections.Values) {
        foreach ($metric in @($totals.Keys)) {
            if ($section.Contains($metric)) { $totals[$metric] += $section[$metric] }
        }
    }

    return [ordered]@{
        TotalFilesProcessed = $totals.Processed
        TotalFilesMoved     = $totals.Moved
        TotalFilesSkipped   = $totals.Skipped
        TotalFilesFailed    = $totals.Failed
        TotalSizeBytes      = $totals.Size
        TotalSizeGB         = [Math]::Round(($totals.Size) / 1GB, 2)
        TotalErrors         = $Report.Errors.Count
        TotalWarnings       = $Report.Warnings.Count
    }
}

function Save-Report {
    <#
    .SYNOPSIS
        Finalizes the report (end time, duration, status, summary) and writes it as JSON.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $Report.EndTime = Get-Date
        $Report.DurationSeconds = [Math]::Round(($Report.EndTime - $Report.StartTime).TotalSeconds, 2)
        if ($Report.Status -eq 'Running') {
            $Report.Status = 'Completed'
        }
        $Report.Summary = Get-ReportSummary -Report $Report

        $directory = Split-Path -Path $Path -Parent
        if ($directory -and -not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force -WhatIf:$false -Confirm:$false | Out-Null
        }

        $Report | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $Path -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Log -Message "Report saved to: $Path" -Level 'Success'
    }
    catch {
        Write-Log -Message "Failed to save report: $_" -Level 'Error'
    }
}

#endregion
