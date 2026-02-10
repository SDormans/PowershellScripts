# File Organization Scripts

A collection of **PowerShell scripts** to automatically organize files such as documents, photos, and music into clean, structured folders.

These scripts are designed to be **safe, readable, and easy to customize**, making them suitable for both personal use and small automation setups.

---

## 📁 Repository Overview

This repository contains PowerShell scripts that help you:

* Organize downloaded files into dedicated folders
* Consolidate photos and clean up photo directories
* Structure music libraries and remove duplicates
* Keep folders tidy by removing empty or unwanted directories

All scripts are heavily commented and structured for clarity.

---

## 🧰 Scripts Included

### 1. Download Organizer

Organizes files from a source folder (e.g. Downloads) into:

* **Personal** → documents (PDF, Word, Excel, etc.)
* **Music** → audio files (MP3, FLAC, WAV, etc.)
* **Fotos** → images (JPG, PNG, GIF, etc.)

Creates destination folders automatically if they do not exist.

---

### 2. Document, Photo & Music Cleaner

A more advanced organizer that:

* Moves documents into a personal folder
* Consolidates photos into a single root directory
* Separates non-photo files into an `Others` folder
* Cleans empty photo directories
* Organizes music folders
* Removes macOS `_MACOSX` artifacts
* Detects and moves duplicate music folders

---

### 3. Directory Size & Extension Report

Generates a report that:

* Scans a directory and all subdirectories
* Calculates total size per folder (in MB)
* Lists unique file extensions
* Outputs a CLI-friendly table sorted by size

Useful for storage audits and cleanup planning.

---

## ⚙️ Requirements

* Windows
* PowerShell 5.1 or newer
* Appropriate file system permissions

---

## 🚀 Usage

1. Clone or download this repository
2. Open PowerShell
3. Edit the script configuration sections to match your paths
4. Run the script:

```powershell
.\script-name.ps1
```

> 💡 Tip: Run PowerShell as **Administrator** if you are organizing protected directories.

---

## 🛡️ Safety Notes

* Scripts avoid overwriting files unless explicitly forced
* Duplicate handling is conservative
* Always test on a small directory first
* Consider adding `-WhatIf` during testing

---

## 🔧 To Do Ideas

* Add logging to a file
* Enable dry-run (`-WhatIf`) mode
* Convert scripts into reusable PowerShell modules
* Write tests for them
* Schedule scripts using Windows Task Scheduler

---

## 📄 License

This project is provided as-is for personal and educational use.

Feel free to modify and adapt it to your needs.

---

## ✨ Why This Repo Exists

These scripts were written to:

* Reduce manual file sorting
* Improve folder hygiene
* Serve as clean, readable PowerShell examples

If you value **clarity over cleverness**, you’re in the right place 🙂


  
