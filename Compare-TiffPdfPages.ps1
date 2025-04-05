<#
.SYNOPSIS
    TIFF-PDF Compare Tool - Compares TIFF files with PDF pages
.DESCRIPTION
    Compares the number of TIFF files in folders with their corresponding PDF page counts.
    Handles nested folder structures and maintains hierarchy when moving files.
.NOTES
    Author: NeoMatrix14241
    Last Updated: 2025-04-04 23:52:26
#>

# Get the script's directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Show-Menu {
    Clear-Host
    Write-Host "==================================================="
    Write-Host "   TIFF Files and PDF Pages Comparison Tool"
    Write-Host "==================================================="
    Write-Host "Current Date/Time (UTC): 2025-04-04 23:52:26"
    Write-Host "User: NeoMatrix14241"
    Write-Host "==================================================="
    Write-Host
}

function Write-Log {
    param(
        $Message,
        [ValidateSet('Info','Warning','Error')]$Type = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch($Type) {
        'Info'    { 'White' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
    }
    
    Write-Host "[$timestamp] $Type`: $Message" -ForegroundColor $color
}

function Get-FolderPath {
    param([string]$Description)
    
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true
    
    if ($dialog.ShowDialog() -eq 'OK') {
        return $dialog.SelectedPath
    }
    return $null
}

function Test-PathAndPermissions {
    param([string]$Path, [string]$PathType)
    
    try {
        Write-Log "Testing $PathType path: $Path" -Type Info
        if (Test-Path $Path) {
            Write-Log "  Path exists: Yes" -Type Info
            
            # Test if we can list files
            $files = Get-ChildItem -Path $Path -ErrorAction Stop
            Write-Log "  Can list files: Yes" -Type Info
            Write-Log "  Total items found: $($files.Count)" -Type Info
            
            # For output folder, list some PDF files
            if ($PathType -eq "Output") {
                Write-Log "Sample PDFs in output folder:" -Type Info
                Get-ChildItem -Path $Path -Filter "*.pdf" -Recurse | 
                    Select-Object -First 5 | ForEach-Object {
                        Write-Log "  - $($_.FullName)" -Type Info
                    }
            }
            return $true
        } else {
            Write-Log "  Path exists: No" -Type Error
            return $false
        }
    }
    catch {
        Write-Log "  Error accessing path: $_" -Type Error
        return $false
    }
}

function Get-MatchingPdfPath {
    param(
        [string]$OutputPath,
        [System.IO.DirectoryInfo]$TiffFolder
    )
    
    try {
        # Get parent folder directly from FullName
        $parentFolder = Split-Path -Path $TiffFolder.FullName -Parent
        $parentName = Split-Path -Path $parentFolder -Leaf
        
        Write-Log "Checking possible PDF locations:" -Type Info
        Write-Log "  Parent folder: $parentName" -Type Info
        
        # First path: Check in parent folder structure
        $nestedPath = Join-Path -Path $OutputPath -ChildPath $parentName
        $nestedPath = Join-Path -Path $nestedPath -ChildPath "$($TiffFolder.Name).pdf"
        
        # Second path: Check directly in output
        $directPath = Join-Path -Path $OutputPath -ChildPath "$($TiffFolder.Name).pdf"
        
        Write-Log "  Checking nested path: $nestedPath" -Type Info
        if (Test-Path -Path $nestedPath -PathType Leaf) {
            Write-Log "  Found PDF at nested path" -Type Info
            return $nestedPath
        }
        
        Write-Log "  Checking direct path: $directPath" -Type Info
        if (Test-Path -Path $directPath -PathType Leaf) {
            Write-Log "  Found PDF at direct path" -Type Info
            return $directPath
        }
        
        Write-Log "  No PDF found in any location" -Type Warning
        return $null
    }
    catch {
        Write-Log "Error in Get-MatchingPdfPath: $_" -Type Error
        return $null
    }
}

function Get-PdfPageCount {
    param([string]$PdfPath)
    
    try {
        if (-not (Test-Path $PdfPath)) {
            Write-Log "PDF file does not exist: $PdfPath" -Type Error
            return -1
        }

        $pdfData = & pdftk $PdfPath dump_data
        if ($LASTEXITCODE -ne 0) {
            Write-Log "PDFtk failed to read the PDF file" -Type Error
            return -1
        }

        $pageCount = ($pdfData | Select-String "NumberOfPages: (\d+)").Matches.Groups[1].Value
        if ([string]::IsNullOrEmpty($pageCount)) {
            Write-Log "Could not determine page count from PDF" -Type Error
            return -1
        }

        return [int]$pageCount
    }
    catch {
        Write-Log "Error reading PDF: $_" -Type Error
        return -1
    }
}

function Move-TiffFiles {
    param(
        [System.IO.DirectoryInfo]$SourceDir,
        [string]$MoveToPath,
        [string]$InputPath
    )
    
    try {
        # Ensure we have valid parameters
        if (-not $SourceDir -or -not $MoveToPath -or -not $InputPath) {
            Write-Log "Invalid parameters provided to Move-TiffFiles" -Type Error
            return 0
        }

        # Calculate relative path preserving folder structure
        $fullSourcePath = $SourceDir.FullName
        $relativePath = $fullSourcePath.Substring($InputPath.TrimEnd('\').Length).TrimStart('\')
        
        # Create destination path
        $destinationPath = Join-Path -Path $MoveToPath -ChildPath $relativePath
        
        Write-Log "Moving files:" -Type Info
        Write-Log "  From: $fullSourcePath" -Type Info
        Write-Log "  To: $destinationPath" -Type Info
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path -Path $destinationPath)) {
            New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
            Write-Log "  Created destination directory" -Type Info
        }
        
        # Move TIFF files
        $movedCount = 0
        $tiffFiles = Get-ChildItem -Path $fullSourcePath -Filter "*.tif" -File
        
        foreach ($file in $tiffFiles) {
            $destinationFile = Join-Path -Path $destinationPath -ChildPath $file.Name
            Move-Item -Path $file.FullName -Destination $destinationFile -Force
            $movedCount++
            Write-Log "  Moved: $($file.Name)" -Type Info
        }
        
        # Clean up empty directories
        if ($movedCount -gt 0) {
            if (-not (Get-ChildItem -Path $fullSourcePath)) {
                Remove-Item -Path $fullSourcePath -Force
                Write-Log "  Removed empty source folder: $fullSourcePath" -Type Info
                
                # Check and remove parent if empty
                $parentPath = Split-Path -Path $fullSourcePath -Parent
                if (-not (Get-ChildItem -Path $parentPath)) {
                    Remove-Item -Path $parentPath -Force
                    Write-Log "  Removed empty parent folder: $parentPath" -Type Info
                }
            }
        }
        
        return $movedCount
    }
    catch {
        Write-Log "Error in Move-TiffFiles: $_" -Type Error
        return 0
    }
}

# Check if PDFtk is installed
try {
    $null = & pdftk --version
}
catch {
    Write-Host "PDFtk Server is not installed or not in PATH!" -ForegroundColor Red
    Write-Host "Please install PDFtk Server from: https://www.pdflabs.com/tools/pdftk-server/"
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 1
}

# Main Script
do {
    Show-Menu
    
    Write-Host "Please select the operation mode:"
    Write-Host "1. Count only (no file moving)"
    Write-Host "2. Count and move mismatched TIFF files"
    Write-Host "3. Exit"
    Write-Host
    $choice = Read-Host "Enter your choice (1-3)"

    switch ($choice) {
        '1' {
            Write-Host "`nPlease select the input folder (containing TIFF files)..."
            $InputPath = Get-FolderPath "Select Input Folder"
            if (-not $InputPath) { continue }

            Write-Host "`nPlease select the output folder (containing PDF files)..."
            $OutputPath = Get-FolderPath "Select Output Folder"
            if (-not $OutputPath) { continue }

            $CountOnly = $true
            $MoveToPath = $null
        }
        '2' {
            Write-Host "`nPlease select the input folder (containing TIFF files)..."
            $InputPath = Get-FolderPath "Select Input Folder"
            if (-not $InputPath) { continue }

            Write-Host "`nPlease select the output folder (containing PDF files)..."
            $OutputPath = Get-FolderPath "Select Output Folder"
            if (-not $OutputPath) { continue }

            Write-Host "`nPlease select the folder where mismatched TIFF files should be moved..."
            $MoveToPath = Get-FolderPath "Select Move-To Folder"
            if (-not $MoveToPath) { continue }

            $CountOnly = $false
        }
        '3' {
            exit
        }
        default {
            Write-Host "`nInvalid choice. Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            continue
        }
    }

    # Path validation and diagnostics
    Write-Log "=== Starting Path Diagnostics ===" -Type Info
    $inputValid = Test-PathAndPermissions -Path $InputPath -PathType "Input"
    $outputValid = Test-PathAndPermissions -Path $OutputPath -PathType "Output"
    $moveToValid = if ($MoveToPath) { 
        Test-PathAndPermissions -Path $MoveToPath -PathType "Move-To" 
    } else { $true }
    Write-Log "=== End Path Diagnostics ===" -Type Info

    if (-not ($inputValid -and $outputValid -and $moveToValid)) {
        Write-Log "One or more paths are invalid or inaccessible" -Type Error
        Write-Host "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        continue
    }

    # Initialize counters
    $script:totalProcessed = 0
    $script:matchingCount = 0
    $script:mismatchCount = 0
    $script:movedFilesCount = 0

    Write-Log "Starting comparison process..." -Type Info
    Write-Log "Input Path: $InputPath" -Type Info
    Write-Log "Output Path: $OutputPath" -Type Info
    if ($MoveToPath) {
        Write-Log "Move-To Path: $MoveToPath" -Type Info
    }

    # Get all directories containing .tif files
    $tiffDirs = Get-ChildItem -Path $InputPath -Directory -Recurse | 
        Where-Object { 
            Get-ChildItem -Path $_.FullName -Filter "*.tif" -File
        }

    foreach ($tiffDir in $tiffDirs) {
        $folderName = $tiffDir.Name
        Write-Log "Processing folder: $folderName" -Type Info
        Write-Log "Full path: $($tiffDir.FullName)" -Type Info
        
        # Use new function to find PDF
        $pdfPath = Get-MatchingPdfPath -OutputPath $OutputPath -TiffFolder $tiffDir
        
        if (-not $pdfPath) {
            Write-Log "No matching PDF found for folder: $folderName" -Type Warning
            
            if ($MoveToPath -and -not $CountOnly) {
                $movedCount = Move-TiffFiles -SourceDir $tiffDir -MoveToPath $MoveToPath -InputPath $InputPath
                $script:movedFilesCount += $movedCount
            }
            continue
        }
        
        $tiffCount = (Get-ChildItem -Path $tiffDir.FullName -Filter "*.tif" -File).Count
        $pdfPages = Get-PdfPageCount -PdfPath $pdfPath
        
        $script:totalProcessed++
        
        Write-Log "  TIFF files: $tiffCount" -Type Info
        Write-Log "  PDF pages: $pdfPages" -Type Info
        
        if ($pdfPages -eq -1) {
            Write-Log "  Status: SKIPPED (Could not read PDF)" -Type Warning
            if ($MoveToPath -and -not $CountOnly) {
                $movedCount = Move-TiffFiles -SourceDir $tiffDir -MoveToPath $MoveToPath -InputPath $InputPath
                $script:movedFilesCount += $movedCount
            }
            continue
        }
        
        if ($pdfPages -eq $tiffCount) {
            Write-Log "  Status: MATCH" -Type Info
            $script:matchingCount++
        }
        else {
            Write-Log "  Status: MISMATCH" -Type Warning
            Write-Log "  Expected $tiffCount pages, found $pdfPages pages" -Type Warning
            $script:mismatchCount++
            
            if ($MoveToPath -and -not $CountOnly) {
                $movedCount = Move-TiffFiles -SourceDir $tiffDir -MoveToPath $MoveToPath -InputPath $InputPath
                $script:movedFilesCount += $movedCount
            }
        }
    }

    # Display summary
    Write-Log "`nSummary Report:" -Type Info
    Write-Log "Total folders processed: $script:totalProcessed" -Type Info
    Write-Log "Matching counts: $script:matchingCount" -Type Info
    Write-Log "Mismatching counts: $script:mismatchCount" -Type Info
    Write-Log "Total files moved: $script:movedFilesCount" -Type Info

    if ($CountOnly) {
        Write-Log "Running in count-only mode - no files were moved" -Type Info
    }

    Write-Host "`nPress any key to return to the menu..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

} while ($true)
