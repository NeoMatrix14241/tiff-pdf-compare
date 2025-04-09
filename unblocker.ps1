# Get the directory where the script is located
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Path

# Initialize counters
$totalFiles = 0
$unblockedFiles = 0

Write-Host "Starting to unblock files recursively..." -ForegroundColor Green

# Get all files recursively
Get-ChildItem -Path $scriptPath -Recurse -File | ForEach-Object {
    $totalFiles++
    
    try {
        # Check if file is blocked
        $stream = [System.IO.File]::OpenRead($_.FullName)
        $stream.Close()
        
        $zone = Get-Item -Path $_.FullName -Stream "Zone.Identifier" -ErrorAction SilentlyContinue
        if ($zone) {
            # Unblock the file
            Unblock-File -Path $_.FullName
            $unblockedFiles++
            Write-Host "Unblocked:" $_.FullName -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error processing:" $_.FullName -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}

Write-Host "`nOperation completed!" -ForegroundColor Green
Write-Host "Total files scanned: $totalFiles" -ForegroundColor Cyan
Write-Host "Files unblocked: $unblockedFiles" -ForegroundColor Cyan

Read-Host -Prompt "Press Enter to exit"