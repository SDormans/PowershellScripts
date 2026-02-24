# 🚀 PowerShell File Organization Scripts

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/SDormans/PowershellScripts/pulls)

A collection of PowerShell scripts for enterprise-grade file organization, cleaning, and reporting. These scripts transform chaotic file structures into organized, maintainable hierarchies with comprehensive logging, error handling, and audit capabilities.

---

## 📋 Table of Contents
- [✨ Key Features](#-key-features)
- [📦 Scripts Overview](#-scripts-overview)
- [⚙️ Installation](#️-installation)
- [🚀 Quick Start](#-quick-start)
- [📖 Detailed Usage](#-detailed-usage)
- [🛡️ Safety & Best Practices](#️-safety--best-practices)
- [📊 Reporting & Logging](#-reporting--logging)
- [🔧 Configuration](#-configuration)
- [📅 Automation](#-automation)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)

---

## ✨ Key Features

### 🏢 Enterprise-Grade Reliability
- **Comprehensive error handling** with try/catch blocks and recovery mechanisms
- **Atomic file operations** using temporary files to prevent data loss
- **Timeout protection** to prevent hanging on network drives
- **Retry logic** for transient failures

### 📊 Professional Reporting
- **JSON reports** with detailed statistics per operation
- **Colored console output** for real-time monitoring
- **Structured logging** with multiple verbosity levels
- **Audit trails** of all file movements and changes

### 🛡️ Safety First
- **WhatIf/Dry-run mode** to preview changes before execution
- **Conservative duplicate handling** with versioning
- **No accidental overwrites** without explicit force flags
- **Permission validation** before operations

### ⚡ Performance Optimized
- **Parallel processing** for large file sets
- **Configurable throttling** to manage system load
- **Smart caching** of directory structures
- **Minimal memory footprint**

---

## 📦 Scripts Overview

### 1️⃣ Download Organizer (`DownloadOrganizer.ps1`)
Enterprise-grade download folder manager with advanced categorization.

**Features:**
- 📁 Automatic sorting into Personal, Music, and Photos folders
- 🔧 Configurable extension mapping via JSON
- 📊 Detailed processing reports
- ⚡ Parallel processing support
- 🔍 Duplicate detection and handling

**Supported formats:**
- **Documents:** PDF, DOC/DOCX, XLS/XLSX, PPT/PPTX, TXT, RTF, ODT, and more
- **Music:** MP3, FLAC, WAV, AAC, M4A, OGG, WMA
- **Photos:** JPG, JPEG, PNG, GIF, BMP, TIFF, HEIC, WebP, RAW formats

---

### 2️⃣ File System Cleaner (`FileOrganizer.ps1`)
Advanced multi-stage organizer for deep file system maintenance.

**Three-stage cleaning process:**

#### Stage 1: Document Organization
- Scans entire drive for document files
- Consolidates into centralized Personal folder
- Preserves folder structure (optional)

#### Stage 2: Photo Library Management
- Flattens nested photo structures
- Separates non-photo files into `Others` folder
- Intelligent empty directory cleanup
- Handles RAW and HEIC formats

#### Stage 3: Music Library Optimization
- Removes macOS artifacts (`_MACOSX`, `.DS_Store`)
- Smart duplicate album detection
- Versioned duplicate storage
- Non-music folder cleanup

---

### 3️⃣ Directory Analyzer (`DirectoryReport.ps1`)
Professional storage audit and reporting tool.

**Capabilities:**
- 📏 Recursive size calculation per directory
- 🔤 Unique extension inventory
- 📈 Top 5 extensions per folder
- ⏱️ Scan duration tracking
- 📤 Multiple export formats (CSV, JSON, HTML)
- 🎯 Filter by minimum size
- 🔍 Timeout protection for deep scans

**Sample output:**
```
Directory                    Size     Files  Exts  Top Extensions
---------------------------  -------  -----  ----  ------------------------
F:\Projects                  2.45 GB  1245   12    .cs (456); .json (234)
F:\Photos                    856 MB   342    8     .jpg (201); .png (89)
F:\Downloads                 234 MB   67     15    .pdf (23); .zip (12)
```

---

## ⚙️ Installation

### Prerequisites
- Windows 7/8/10/11 or Windows Server 2012+
- PowerShell 5.1 or higher
- Appropriate file system permissions

### Quick Install
```powershell
# Clone the repository
git clone https://github.com/SDormans/PowershellScripts.git
cd PowershellScripts

# (Optional) Unblock scripts
Get-ChildItem -Path . -Recurse *.ps1 | Unblock-File

# Set execution policy (if needed)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## 🚀 Quick Start

### Basic Usage Examples

```powershell
# 1. Test the Download Organizer (dry run)
.\DownloadOrganizer.ps1 -SourceDir "C:\Users\$env:USERNAME\Downloads" -WhatIf

# 2. Run File Organizer with verbose output
.\FileOrganizer.ps1 -SourceRoot "D:\Data" -Verbose

# 3. Generate directory report and export to HTML
.\DirectoryReport.ps1 -Path "F:\" -ExportFormat HTML -MinSizeMB 100
```

### Production Scenarios

```powershell
# Automated nightly cleanup with configuration
.\FileOrganizer.ps1 -ConfigPath ".\config\production.json" -Parallel -LogPath "D:\Logs\cleanup.log"

# Storage audit with JSON export for dashboard
.\DirectoryReport.ps1 -Path "\\nas\shared" -ExportFormat JSON -TimeoutSeconds 600

# Quick download organization with reporting
.\DownloadOrganizer.ps1 -SourceDir "C:\Downloads" -ReportPath ".\reports\downloads.json"
```

---

## 📖 Detailed Usage

### Download Organizer Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-SourceDir` | Directory to scan | `C:\Temp\Downloads` |
| `-PersonalDir` | Destination for documents | `C:\Users\YourName\Personal` |
| `-MusicDir` | Destination for music | `C:\Users\YourName\Music` |
| `-FotosDir` | Destination for photos | `C:\Users\YourName\Fotos` |
| `-ConfigPath` | JSON configuration file | `$null` |
| `-Parallel` | Enable parallel processing | `$false` |
| `-ThrottleLimit` | Max concurrent operations | `5` |
| `-PreserveFolderStructure` | Keep subfolder structure | `$false` |
| `-Force` | Overwrite existing files | `$false` |
| `-WhatIf` | Preview changes | `$false` |
| `-LogPath` | Log file location | Auto-generated |
| `-ReportPath` | JSON report location | Auto-generated |
| `-TimeoutSeconds` | Operation timeout | `300` |

### File Organizer Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-SourceRoot` | Root directory to scan | `F:\` |
| `-ConfigPath` | JSON configuration file | `$null` |
| `-Parallel` | Enable parallel processing | `$false` |
| `-ThrottleLimit` | Max concurrent operations | `5` |
| `-WhatIf` | Preview changes | `$false` |
| `-LogPath` | Log file location | Auto-generated |
| `-ReportPath` | JSON report location | Auto-generated |
| `-TimeoutSeconds` | Operation timeout | `300` |

### Directory Report Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-Path` | **Required** Directory to analyze | - |
| `-MinSizeMB` | Minimum directory size to show | `0` |
| `-SortBy` | Sort method (Size/Name/ExtensionCount) | `Size` |
| `-ExportFormat` | Export format (None/CSV/JSON/HTML) | `None` |
| `-OutputPath` | Export destination | Current directory |
| `-TimeoutSeconds` | Scan timeout per directory | `300` |
| `-WhatIf` | Preview scan | `$false` |
| `-LogPath` | Log file location | Auto-generated |

---

## 🛡️ Safety & Best Practices

### Before First Run
1. **Always use `-WhatIf` first** to preview changes
2. Test on a small, non-critical directory
3. Verify destination paths have sufficient space
4. Ensure you have proper permissions

### Production Guidelines
```powershell
# Recommended testing workflow
1. .\script.ps1 -WhatIf -Verbose
2. .\script.ps1 -WhatIf -Verbose -LogPath "test.log"
3. .\script.ps1 -Verbose -LogPath "production.log"
4. Review the JSON report after execution
```

### Safety Features
- ✅ **No automatic overwrites** without `-Force`
- ✅ **Transaction-like moves** with temp files
- ✅ **Timeout protection** for network drives
- ✅ **Permission validation** before operations
- ✅ **Comprehensive logging** for audit trails
- ✅ **Duplicate versioning** instead of deletion

---

## 📊 Reporting & Logging

### Log File Structure
```
[2024-02-24 14:30:15.123] [Info] ========== DOWNLOAD ORGANIZER STARTED ==========
[2024-02-24 14:30:15.124] [Info] Source directory: C:\Downloads
[2024-02-24 14:30:15.125] [Info] Log file: C:\Logs\DownloadOrganizer_20240224_143015.log
[2024-02-24 14:30:15.126] [Success] Found 42 file(s) to process
[2024-02-24 14:30:15.127] [Success] [MOVED] report.pdf -> Personal folder
[2024-02-24 14:30:15.128] [Warning] Skipping song.mp3 - already exists in Music folder
```

### JSON Report Example
```json
{
  "StartTime": "2024-02-24T14:30:15",
  "EndTime": "2024-02-24T14:31:42",
  "DurationSeconds": 87.24,
  "SourceDirectory": "C:\\Downloads",
  "FilesProcessed": 42,
  "FilesMoved": 38,
  "FilesSkipped": 3,
  "FilesFailed": 1,
  "TotalSizeBytes": 1584328704,
  "Categories": {
    "Document": { "Count": 15, "Size": 24567891 },
    "Music": { "Count": 12, "Size": 985432167 },
    "Photo": { "Count": 11, "Size": 574239546 },
    "Unknown": { "Count": 4, "Size": 1024000 }
  },
  "Errors": [
    {
      "Time": "2024-02-24T14:31:40",
      "Message": "Failed to move locked-file.pdf: Access denied"
    }
  ]
}
```

---

## 🔧 Configuration

### JSON Configuration File Example (`config.json`)
```json
{
    "SourceRoot": "F:\\",
    "PersonalFolder": "F:\\Documenten",
    "PhotoRoot": "F:\\FotoBibliotheek",
    "MusicRoot": "F:\\MuziekCollectie",
    "TimeoutSeconds": 600,
    "Parallel": true,
    "ThrottleLimit": 8,
    "PreserveFolderStructure": false,
    "Extensions": {
        "Documents": [".pdf", ".docx", ".xlsx", ".pptx"],
        "Music": [".mp3", ".flac", ".wav"],
        "Photos": [".jpg", ".png", ".heic", ".raw"]
    }
}
```

### Using Configuration Files
```powershell
.\FileOrganizer.ps1 -ConfigPath ".\config\production.json"
.\DownloadOrganizer.ps1 -ConfigPath ".\config\downloads.json" -Parallel
```

---

## 📅 Automation

### Windows Task Scheduler Setup
```powershell
# Create scheduled task for weekly cleanup
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-File `"C:\Scripts\FileOrganizer.ps1`" -ConfigPath `"C:\Scripts\config.json`" -Parallel"

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2am

$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount

Register-ScheduledTask -TaskName "Weekly File Organizer" `
    -Action $action -Trigger $trigger -Principal $principal
```

### PowerShell Scheduled Job
```powershell
$jobOptions = New-ScheduledJobOption -RunElevated
$trigger = New-JobTrigger -Weekly -At "3:00 AM" -DaysOfWeek Monday,Wednesday,Friday

Register-ScheduledJob -Name "DownloadOrganizer" `
    -FilePath "C:\Scripts\DownloadOrganizer.ps1" `
    -ArgumentList "-ConfigPath C:\Scripts\config.json" `
    -Trigger $trigger -ScheduledJobOption $jobOptions
```

---

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### Development Guidelines
- Maintain the existing coding style
- Add comprehensive comments
- Include error handling for new features
- Update documentation for changes
- Test with `-WhatIf` before submitting

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---
