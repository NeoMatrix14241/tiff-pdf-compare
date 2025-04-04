<#
.SYNOPSIS
    Compares the number of TIFF files in folders with the page count of corresponding PDF files using PDFtk Server.
#>

function Show-Menu {
    Clear-Host
    Write-Host "==================================================="
    Write-Host "   TIFF Files and PDF Pages Comparison Tool"
    Write-Host "==================================================="
    Write-Host "Current Date/Time (UTC): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Host "User: $env:USERNAME"
    Write-Host "==================================================="
    Write-Host
}

function Write-Log {
    param($Message, [ValidateSet('Info','Warning','Error')]$Type = 'Info')
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch($Type) {
        'Info'    { 'White' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
    }
    
    Write-Host "[$timestamp] $Type`: $Message" -ForegroundColor $color
}

function Get-PdfPageCount {
    param([string]$PdfPath)
    
    try {
        if (-not (Test-Path $PdfPath)) {
            Write-Log "PDF file does not exist: $PdfPath" -Type Error
            return -1
        }

        # Use PDFtk to dump PDF data and extract page count
        $pdfData = & pdftk $PdfPath dump_data
        if ($LASTEXITCODE -ne 0) {
            Write-Log "PDFtk failed to read the PDF file" -Type Error
            return -1
        }

        # Extract NumberOfPages from PDFtk output
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

    # Validate paths
    if (-not (Test-Path $InputPath)) {
        Write-Log "Input path does not exist: $InputPath" -Type Error
        continue
    }

    if (-not (Test-Path $OutputPath)) {
        Write-Log "Output path does not exist: $OutputPath" -Type Error
        continue
    }

    if ($MoveToPath -and -not (Test-Path $MoveToPath)) {
        try {
            New-Item -Path $MoveToPath -ItemType Directory -Force | Out-Null
            Write-Log "Created move-to directory: $MoveToPath" -Type Info
        }
        catch {
            Write-Log "Failed to create move-to directory: $MoveToPath. Error: $_" -Type Error
            continue
        }
    }

    # Initialize counters
    $script:totalProcessed = 0
    $script:matchingCount = 0
    $script:mismatchCount = 0

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
        $pdfPath = Join-Path $OutputPath "$folderName.pdf"
        
        if (-not (Test-Path $pdfPath)) {
            Write-Log "No matching PDF found for folder: $folderName" -Type Warning
            continue
        }
        
        $tiffCount = (Get-ChildItem -Path $tiffDir.FullName -Filter "*.tif" -File).Count
        $pdfPages = Get-PdfPageCount -PdfPath $pdfPath
        
        $script:totalProcessed++
        
        Write-Log "Processing: $folderName" -Type Info
        Write-Log "  TIFF files: $tiffCount" -Type Info
        Write-Log "  PDF pages: $pdfPages" -Type Info
        
        if ($pdfPages -eq -1) {
            Write-Log "  Status: SKIPPED (Could not read PDF)" -Type Warning
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
                try {
                    # Create the same folder structure in the archive folder
                    $relativePath = $tiffDir.FullName.Substring($InputPath.Length)
                    $archivePath = Join-Path $MoveToPath $relativePath
                    
                    # Create the directory in the archive if it doesn't exist
                    if (-not (Test-Path $archivePath)) {
                        New-Item -Path $archivePath -ItemType Directory -Force | Out-Null
                    }

                    # Move all .tif files from the source to archive, maintaining folder structure
                    Get-ChildItem -Path $tiffDir.FullName -Filter "*.tif" -File | ForEach-Object {
                        $destinationFile = Join-Path $archivePath $_.Name
                        Move-Item -Path $_.FullName -Destination $destinationFile -Force
                        Write-Log "  Moved TIFF file: $($_.Name) to $archivePath" -Type Info
                    }

                    # Remove empty source folder if all files were moved
                    if (-not (Get-ChildItem -Path $tiffDir.FullName)) {
                        Remove-Item -Path $tiffDir.FullName -Force
                        Write-Log "  Removed empty source folder: $($tiffDir.FullName)" -Type Info
                    }
                }
                catch {
                    Write-Log "  Failed to move TIFF files: $_" -Type Error
                }
            }
        }
    }

    # Display summary
    Write-Log "`nSummary Report:" -Type Info
    Write-Log "Total folders processed: $script:totalProcessed" -Type Info
    Write-Log "Matching counts: $script:matchingCount" -Type Info
    Write-Log "Mismatching counts: $script:mismatchCount" -Type Info

    if ($CountOnly) {
        Write-Log "Running in count-only mode - no files were moved" -Type Info
    }

    Write-Host "`nPress any key to return to the menu..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

} while ($true)