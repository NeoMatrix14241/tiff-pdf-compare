<#
.SYNOPSIS
    TIFF-PDF Compare Tool with Parallel Processing
.DESCRIPTION
    Compares TIFF files with PDF pages using parallel processing for faster execution.
    Current Date/Time (UTC): 2025-04-05 00:05:13
    User: NeoMatrix14241
#>

# Get the script's directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Create a runspace pool for parallel processing
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, [Environment]::ProcessorCount)
$RunspacePool.Open()
$Jobs = New-Object System.Collections.ArrayList

# Script constants
$CURRENT_UTC_TIME = "2025-04-05 00:05:13"
$CURRENT_USER = "NeoMatrix14241"

function Show-Menu {
    Clear-Host
    Write-Host "==================================================="
    Write-Host "   TIFF Files and PDF Pages Comparison Tool"
    Write-Host "==================================================="
    Write-Host "Current Date/Time (UTC): $CURRENT_UTC_TIME"
    Write-Host "User: $CURRENT_USER"
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

function Start-ParallelPdfCheck {
    param(
        [string]$PdfPath
    )
    
    $PowerShell = [powershell]::Create().AddScript({
        param($PdfPath)
        
        try {
            if (-not (Test-Path $PdfPath)) {
                return @{
                    Path = $PdfPath
                    Pages = -1
                    Error = "File not found"
                }
            }

            $pdfData = & pdftk $PdfPath dump_data
            if ($LASTEXITCODE -ne 0) {
                return @{
                    Path = $PdfPath
                    Pages = -1
                    Error = "PDFtk failed"
                }
            }

            $pageCount = ($pdfData | Select-String "NumberOfPages: (\d+)").Matches.Groups[1].Value
            if ([string]::IsNullOrEmpty($pageCount)) {
                return @{
                    Path = $PdfPath
                    Pages = -1
                    Error = "Could not determine page count"
                }
            }

            return @{
                Path = $PdfPath
                Pages = [int]$pageCount
                Error = $null
            }
        }
        catch {
            return @{
                Path = $PdfPath
                Pages = -1
                Error = $_.Exception.Message
            }
        }
    }).AddArgument($PdfPath)

    $PowerShell.RunspacePool = $RunspacePool

    [void]$Jobs.Add(@{
        PowerShell = $PowerShell
        Handle = $PowerShell.BeginInvoke()
        Path = $PdfPath
    })
}

function Wait-AllJobs {
    foreach ($Job in $Jobs) {
        $Result = $Job.PowerShell.EndInvoke($Job.Handle)
        $Job.PowerShell.Dispose()
        $Result
    }
    $Jobs.Clear()
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
        # Get the relative path from the input folder
        $folderName = $TiffFolder.Name
        $relativePath = $TiffFolder.FullName -replace [regex]::Escape($InputPath), ""
        $relativePath = $relativePath.TrimStart('\')
        
        Write-Log "Checking possible PDF locations:" -Type Info
        
        # First path: Check with full folder structure
        $fullStructurePath = Join-Path -Path $OutputPath -ChildPath $relativePath
        $fullStructurePath = Join-Path -Path $fullStructurePath -ChildPath "$folderName.pdf"
        
        # Second path: Check direct in output with relative path
        $directPath = Join-Path -Path $OutputPath -ChildPath "$folderName.pdf"
        
        Write-Log "  Checking full structure path: $fullStructurePath" -Type Info
        if (Test-Path -Path $fullStructurePath -PathType Leaf) {
            Write-Log "  Found PDF at full structure path" -Type Info
            return $fullStructurePath
        }
        
        Write-Log "  Checking direct path: $directPath" -Type Info
        if (Test-Path -Path $directPath -PathType Leaf) {
            Write-Log "  Found PDF at direct path" -Type Info
            return $directPath
        }
        
        # Third path: Check with preserved folder structure
        $preservedPath = Join-Path -Path $OutputPath -ChildPath $relativePath
        $files = Get-ChildItem -Path $preservedPath -Filter "*.pdf" -File
        if ($files.Count -gt 0) {
            Write-Log "  Found PDF in preserved structure: $($files[0].FullName)" -Type Info
            return $files[0].FullName
        }
        
        Write-Log "  No PDF found in any location" -Type Warning
        return $null
    }
    catch {
        Write-Log "Error in Get-MatchingPdfPath: $_" -Type Error
        return $null
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
        $tiffFiles = Get-ChildItem -Path $fullSourcePath -Filter "*.tif", "*.tiff" -File
        
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
            Write-Host "`n[Containing TIFF/TIF FILES] Please select TIFF/TIF folder..."
            $InputPath = Get-FolderPath "Select Input Folder"
            if (-not $InputPath) { continue }

            Write-Host "`n[Containing PDF FILES]Please select the PDF folder..."
            $OutputPath = Get-FolderPath "Select Output Folder"
            if (-not $OutputPath) { continue }

            $CountOnly = $true
            $MoveToPath = $null
        }
        '2' {
            Write-Host "`n[Containing TIFF/TIF FILES] Please select TIFF/TIF folder..."
            $InputPath = Get-FolderPath "Select Input Folder"
            if (-not $InputPath) { continue }

            Write-Host "`n[Containing PDF FILES]Please select the PDF folder..."
            $OutputPath = Get-FolderPath "Select Output Folder"
            if (-not $OutputPath) { continue }

            Write-Host "`n[Where To Move] Please select the folder where mismatched TIFF/TIF files should be moved..."
            $MoveToPath = Get-FolderPath "Select Move-To Folder"
            if (-not $MoveToPath) { continue }

            $CountOnly = $false
        }
        '3' {
            if ($RunspacePool) {
                $RunspacePool.Close()
                $RunspacePool.Dispose()
            }
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
    $script:pdfCheckErrors = 0

    Write-Log "Starting comparison process..." -Type Info
    Write-Log "Input Path: $InputPath" -Type Info
    Write-Log "Output Path: $OutputPath" -Type Info
    if ($MoveToPath) {
        Write-Log "Move-To Path: $MoveToPath" -Type Info
    }

    # Get all directories containing .tif or .tiff files
    $tiffDirs = Get-ChildItem -Path $InputPath -Directory -Recurse | 
        Where-Object { 
            Get-ChildItem -Path $_.FullName -Filter "*.tif", "*.tiff" -File
        }

    $totalDirs = $tiffDirs.Count
    Write-Log "Found $totalDirs directories containing TIFF files" -Type Info

    # Process directories in parallel
    foreach ($tiffDir in $tiffDirs) {
        $folderName = $tiffDir.Name
        Write-Log "Processing folder: $folderName" -Type Info
        
        $pdfPath = Get-MatchingPdfPath -OutputPath $OutputPath -TiffFolder $tiffDir
        
        if ($pdfPath) {
            Start-ParallelPdfCheck -PdfPath $pdfPath
        }
        else {
            if ($MoveToPath -and -not $CountOnly) {
                $movedCount = Move-TiffFiles -SourceDir $tiffDir -MoveToPath $MoveToPath -InputPath $InputPath
                $script:movedFilesCount += $movedCount
            }
        }
    }

    # Wait for all PDF checks to complete
    Write-Log "Waiting for parallel PDF checks to complete..." -Type Info
    $Results = Wait-AllJobs

    # Process results
    foreach ($Result in $Results) {
        $pdfPath = $Result.Path
        $folderName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
        $tiffDir = $tiffDirs | Where-Object { $_.Name -eq $folderName } | Select-Object -First 1
        
        if (-not $tiffDir) {
            Write-Log "Could not find matching TIFF folder for PDF: $pdfPath" -Type Warning
            continue
        }
        
        $tiffCount = (Get-ChildItem -Path $tiffDir.FullName -Filter "*.tif", "*.tiff" -File).Count
        $pdfPages = $Result.Pages
        
        $script:totalProcessed++
        
        Write-Log "Processing results for: $folderName" -Type Info
        Write-Log "  TIFF files: $tiffCount" -Type Info
        Write-Log "  PDF pages: $pdfPages" -Type Info
        
        if ($Result.Error) {
            Write-Log "  Error processing PDF: $($Result.Error)" -Type Warning
            $script:pdfCheckErrors++
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
    Write-Log "Total directories found: $totalDirs" -Type Info
    Write-Log "Total folders processed: $script:totalProcessed" -Type Info
    Write-Log "Matching counts: $script:matchingCount" -Type Info
    Write-Log "Mismatching counts: $script:mismatchCount" -Type Info
    Write-Log "PDF check errors: $script:pdfCheckErrors" -Type Info
    Write-Log "Total files moved: $script:movedFilesCount" -Type Info

    if ($CountOnly) {
        Write-Log "Running in count-only mode - no files were moved" -Type Info
    }

    Write-Host "`nPress any key to return to the menu..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

} while ($true)