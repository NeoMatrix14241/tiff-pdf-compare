# TIFF-PDF Compare

A PowerShell tool to compare TIFF files count with PDF pages and manage mismatched files.

## Description

This tool is designed to help manage and verify OCR processing results by comparing the number of TIFF files in source folders with the number of pages in corresponding PDF files. It's particularly useful for batch OCR processing workflows where TIFF files are converted to PDFs.

## Features

- Recursive folder scanning
- User-friendly GUI for folder selection
- Detailed logging with timestamps
- Option to move mismatched TIFF files to archive
- Maintains folder structure when moving files
- Color-coded console output
- Summary reporting

## Requirements

- Windows PowerShell 5.1 or later
- PDFtk Server ([Download here](https://www.pdflabs.com/tools/pdftk-server/))
- Folder structure must follow the pattern:
  ```
  Input/
    └── BatchFolder1/
        ├── scan001.tif
        ├── scan002.tif
        └── scan003.tif
  Output/
    └── BatchFolder1.pdf
  Archive/
    └── BatchFolder1/
        ├── scan001.tif
        ├── scan002.tif
        └── scan003.tif
  ```

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/YourUsername/tiff-pdf-compare.git
   ```

2. Install PDFtk Server if you haven't already:
   - Download from [PDFtk Server website](https://www.pdflabs.com/tools/pdftk-server/)
   - Install and ensure it's added to your system's PATH

3. No other dependencies required!

## Usage

### Method 1: Using the Batch File (Recommended)
1. Unblock files by right-clicking > properties > **check unblock** then click **apply** then **ok**
2. Double-click `Compare-TiffPdf.bat`
3. Follow the on-screen menu prompts

### Method 2: Running PowerShell Script Directly
```powershell
.\Compare-TiffPdfPages.ps1
```

## Operation Modes

1. **Count Only Mode**
   - Compares TIFF file counts with PDF pages
   - Reports mismatches without moving files
   - Generates detailed report

2. **Move Mode**
   - Compares TIFF file counts with PDF pages
   - Moves mismatched TIFF files to archive folder
   - Maintains original folder structure
   - Generates detailed report

## Example Output

```
[2025-04-04 23:23:20] Info: Starting comparison process...
[2025-04-04 23:23:20] Info: Input Path: C:\OCR\Input
[2025-04-04 23:23:20] Info: Output Path: C:\OCR\Output
[2025-04-04 23:23:20] Info: Move-To Path: C:\OCR\Archive
[2025-04-04 23:23:20] Info: Processing: BatchFolder1
[2025-04-04 23:23:20] Info:   TIFF files: 3
[2025-04-04 23:23:20] Info:   PDF pages: 3
[2025-04-04 23:23:20] Info:   Status: MATCH
```

## Common Issues

1. **PDFtk Not Found**
   - Ensure PDFtk Server is installed
   - Verify PDFtk is in system PATH
   - Try restarting your terminal

2. **Permission Errors**
   - Run as administrator if needed
   - Check folder permissions

3. **PDF Reading Errors**
   - Verify PDF is not corrupted
   - Ensure PDF is not password protected
   - Check if PDF is locked by another process
