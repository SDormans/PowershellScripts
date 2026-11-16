# PowerShell File Organization Scripts

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PowerShell CI](https://github.com/SDormans/PowershellScripts/actions/workflows/powershell.yml/badge.svg)](https://github.com/SDormans/PowershellScripts/actions/workflows/powershell.yml)

PowerShell scripts for organizing files, cleaning up photo/music libraries and reporting on disk usage, with logging, JSON reports and `-WhatIf` support.

## Table of Contents
- [Scripts](#scripts)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Parameters](#parameters)
- [Configuration](#configuration)
- [Logging and Reports](#logging-and-reports)
- [Safety](#safety)
- [Automation](#automation)
- [Development](#development)
- [License](#license)

---

## Scripts

| Script | Purpose |
|--------|---------|
| [`MovingFiles.ps1`](MovingFiles.ps1) | Sorts a download/inbox folder into Personal (documents), Music and Photos folders. |
| [`DirectoryReorder.ps1`](DirectoryReorder.ps1) | Multi-stage cleanup of a drive or folder: collects documents, flattens the photo library, and tidies the music library. |
| [`DirectoryDistributionChecker.ps1`](DirectoryDistributionChecker.ps1) | Reports the size, file count and extension mix of every directory under a path. |
| [`Common.ps1`](Common.ps1) | Shared helpers (`Write-Log`, `Format-FileSize`, `Test-PathSafe`, `New-FolderIfMissing`, `Move-FileSafely`, `Initialize-Report`, `Update-Report`, `Save-Report`, config loading). Dot-sourced by the scripts above; keep it next to them. |

### MovingFiles.ps1
- Moves documents (PDF, Office, OpenDocument, TXT, ...), music (MP3, FLAC, WAV, ...) and photos (JPG, PNG, HEIC, RAW, ...) to their destination folders.
- Top-level files only by default; `-PreserveFolderStructure` recurses and keeps relative sub-folders.
- Existing destination files are skipped unless `-Force` is passed.
- Processes files sequentially. **Parallel processing (`-Parallel` / `-ThrottleLimit`) has been removed**: the previous implementation required PowerShell 7, skipped duplicate handling and safe moves, and never updated the report. It may return later with thread-safe reporting.

### DirectoryReorder.ps1
1. **Documents** – finds document files anywhere under `-SourceDir` and moves them into `-PersonalDir`.
2. **Photos** – flattens nested folders under `-PhotosDir`, moves non-photo files to `Others`, removes empty folders.
3. **Music** – removes macOS artifacts (`__MACOSX`, `.DS_Store`), moves album folders to the top of `-MusicDir`, moves duplicate albums to `Duplicates`, and removes folders without music.

### DirectoryDistributionChecker.ps1
- Scans the tree **once** and rolls file sizes/counts up into every parent directory (no per-directory background jobs).
- Shows the top 5 extensions per directory, filters by `-MinSizeMB`, sorts by size, name or extension count.
- Exports to CSV, JSON or HTML; `-PassThru` returns the result objects.
- `-TimeoutSeconds` limits the whole scan; directories that were not reached are marked in the report.

Sample output:
```
Directory          Size     Files Exts Top Extensions
---------          ----     ----- ---- --------------
D:\Projects        2.45 GB   1245   12 .cs (456); .json (234); .md (40)
D:\Projects\Photos 856 MB     342    8 .jpg (201); .png (89)
D:\Projects\Docs   234 MB      67   15 .pdf (23); .zip (12)
```

---

## Requirements
- Windows PowerShell 5.1 or PowerShell 7+ (Windows; PowerShell 7 on Linux/macOS also works).
- Read/write permissions on the source and destination folders.

```powershell
git clone https://github.com/SDormans/PowershellScripts.git
cd PowershellScripts

# Optional: unblock downloaded scripts and allow local scripts to run
Get-ChildItem -Path . -Filter *.ps1 -Recurse | Unblock-File
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Quick Start

There are no built-in default paths: pass the folders as parameters or put them in a configuration file.

```powershell
# Preview how a download folder would be sorted
.\MovingFiles.ps1 -SourceDir "$env:USERPROFILE\Downloads" -PersonalDir "D:\Personal" `
    -MusicDir "D:\Music" -PhotosDir "D:\Photos" -WhatIf

# Reorganize a data drive, showing progress messages
.\DirectoryReorder.ps1 -SourceDir "D:\" -PersonalDir "D:\Personal" -PhotosDir "D:\Photos" `
    -MusicDir "D:\Music" -InformationAction Continue

# Disk usage report exported to HTML, only directories of 100 MB or more
.\DirectoryDistributionChecker.ps1 -Path "D:\" -ExportFormat HTML -MinSizeMB 100

# Use a configuration file (see below)
.\MovingFiles.ps1 -ConfigPath ".\config.json" -WhatIf
.\DirectoryReorder.ps1 -ConfigPath ".\config.json" -DryRun

# Storage audit of a share, JSON export for a dashboard
.\DirectoryDistributionChecker.ps1 -Path "\\nas\shared" -ExportFormat JSON -TimeoutSeconds 600
```

---

## Parameters

### MovingFiles.ps1

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-SourceDir` | Folder to sort | **Required** (parameter or config) |
| `-PersonalDir` | Destination for documents | **Required** (parameter or config) |
| `-MusicDir` | Destination for music | **Required** (parameter or config) |
| `-PhotosDir` (alias `-FotosDir`) | Destination for photos | **Required** (parameter or config) |
| `-ConfigPath` | JSON configuration file | none |
| `-PreserveFolderStructure` | Recurse and keep sub-folder structure | `$false` |
| `-Force` | Overwrite existing destination files | `$false` |
| `-TimeoutSeconds` | Stop processing after this many seconds | `300` |
| `-LogPath` | Log file | `MovingFiles_<timestamp>.log` next to the script |
| `-ReportPath` | JSON report | `MovingFiles_Report_<timestamp>.json` next to the script |
| `-WhatIf` / `-Confirm` | Preview / confirm changes | |

### DirectoryReorder.ps1

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-SourceDir` (alias `-SourceRoot`) | Root folder scanned for documents | **Required** (parameter or config) |
| `-PersonalDir` (alias `-PersonalFolder`) | Destination for documents | **Required** (parameter or config) |
| `-PhotosDir` (alias `-PhotoRoot`) | Photo library to flatten | **Required** (parameter or config) |
| `-MusicDir` (alias `-MusicRoot`) | Music library to clean | **Required** (parameter or config) |
| `-ConfigPath` | JSON configuration file | none |
| `-DryRun` | Same as `-WhatIf` | `$false` |
| `-TimeoutSeconds` | Stop the document stage after this many seconds | `300` |
| `-LogPath` | Log file | `DirectoryReorder_<timestamp>.log` next to the script |
| `-ReportPath` | JSON report | `DirectoryReorder_Report_<timestamp>.json` next to the script |
| `-WhatIf` / `-Confirm` | Preview / confirm changes | |

### DirectoryDistributionChecker.ps1

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-Path` | Directory to analyze | **Required** |
| `-MinSizeMB` | Minimum directory size to show | `0` |
| `-SortBy` | `Size`, `Name` or `ExtensionCount` | `Size` |
| `-ExportFormat` | `None`, `CSV`, `JSON` or `HTML` | `None` |
| `-OutputPath` | Export folder | Script folder |
| `-TimeoutSeconds` | Timeout for the whole scan | `300` |
| `-LogPath` | Log file | `DirectoryDistributionChecker_<timestamp>.log` next to the script |
| `-PassThru` | Return the result objects | `$false` |

Parameter names are consistent across scripts (`SourceDir`, `PersonalDir`, `PhotosDir`, `MusicDir`). The old names (`FotosDir`, `SourceRoot`, `PersonalFolder`, `PhotoRoot`, `MusicRoot`) still work as aliases.

---

## Configuration

Copy [`config.example.json`](config.example.json) to `config.json` (ignored by git) and adjust the paths. Each script reads its own section (`MovingFiles`, `DirectoryReorder`); a file without sections is also accepted.

```json
{
    "MovingFiles": {
        "SourceDir": "C:\\Users\\<you>\\Downloads",
        "PersonalDir": "D:\\Personal",
        "PhotosDir": "D:\\Photos",
        "MusicDir": "D:\\Music",
        "PreserveFolderStructure": false,
        "TimeoutSeconds": 300
    },
    "DirectoryReorder": {
        "SourceDir": "D:\\",
        "PersonalDir": "D:\\Personal",
        "PhotosDir": "D:\\Photos",
        "MusicDir": "D:\\Music",
        "TimeoutSeconds": 300
    }
}
```

Precedence: explicit parameter > configuration file > built-in default. The scripts throw a terminating error when a required path is missing.

---

## Logging and Reports

Messages go to the standard PowerShell streams instead of `Write-Host`, so they can be redirected or silenced:

| Level | Stream | Shown by default |
|-------|--------|------------------|
| Info / Success | Information (6) | PowerShell 7 hosts: no – use `-InformationAction Continue`; the log file always has them |
| Warning | Warning (3) | yes |
| Error | Error (2) | yes |

Every message is also appended to the log file:
```
[2026-02-24 14:30:15.123] [Info] ========== MOVING FILES STARTED ==========
[2026-02-24 14:30:15.126] [Info] Found 42 file(s) to process
[2026-02-24 14:30:15.127] [Success] [MOVED] C:\Downloads\report.pdf -> D:\Personal\report.pdf
[2026-02-24 14:30:15.128] [Warning] Skipping song.mp3 - already exists in Music folder
```

`MovingFiles.ps1` and `DirectoryReorder.ps1` write a JSON report:
```json
{
  "StartTime": "2026-02-24T14:30:15",
  "Status": "Completed",
  "Script": "MovingFiles",
  "SourceDirectory": "C:\\Downloads",
  "Sections": {
    "Document": { "Processed": 15, "Moved": 14, "Skipped": 0, "Failed": 1, "Size": 24567891 },
    "Music":    { "Processed": 12, "Moved": 11, "Skipped": 1, "Failed": 0, "Size": 985432167 }
  },
  "Errors": [
    { "Time": "2026-02-24T14:31:40", "Message": "Failed to move C:\\Downloads\\locked.pdf : Access denied" }
  ],
  "Warnings": [],
  "EndTime": "2026-02-24T14:31:42",
  "DurationSeconds": 87.24,
  "Summary": {
    "TotalFilesProcessed": 27, "TotalFilesMoved": 25, "TotalFilesSkipped": 1, "TotalFilesFailed": 1,
    "TotalSizeBytes": 1010000058, "TotalSizeGB": 0.94, "TotalErrors": 1, "TotalWarnings": 0
  }
}
```

Fatal errors are thrown (no `exit 1`), so callers can use `try`/`catch`; when run with `pwsh -File`, a thrown error still produces a non-zero exit code.

---

## Safety
- Run with `-WhatIf` (or `-DryRun` for `DirectoryReorder.ps1`) first and review the log.
- Existing files are never overwritten without `-Force`; duplicate albums are moved to `Duplicates` instead of deleted.
- Same-volume moves use `Move-Item` (a rename). Cross-volume moves copy to a temporary file, verify the size, rename it into place, and only then delete the source.
- Test on a small, non-critical folder first.

---

## Automation

```powershell
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument '-NoProfile -File "C:\Scripts\DirectoryReorder.ps1" -ConfigPath "C:\Scripts\config.json"'
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2am
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount

Register-ScheduledTask -TaskName 'Weekly Directory Reorder' -Action $action -Trigger $trigger -Principal $principal
```

```powershell
$trigger = New-JobTrigger -Weekly -At '3:00 AM' -DaysOfWeek Monday, Wednesday, Friday
Register-ScheduledJob -Name 'MovingFiles' -FilePath 'C:\Scripts\MovingFiles.ps1' `
    -ArgumentList @{ ConfigPath = 'C:\Scripts\config.json' } -Trigger $trigger
```

---

## Development

```powershell
Install-Module PSScriptAnalyzer, Pester -Scope CurrentUser -SkipPublisherCheck

# Static analysis (same settings as CI)
Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1

# Unit and integration tests (Pester 5)
Invoke-Pester -Path .\tests -Output Detailed
```

[GitHub Actions](.github/workflows/powershell.yml) runs PSScriptAnalyzer and the Pester tests on Windows PowerShell 5.1 and PowerShell 7 for every push to `main` and every pull request.

Contributions are welcome: fork the repository, create a feature branch, make sure the analyzer and tests pass, and open a pull request.

---

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.
