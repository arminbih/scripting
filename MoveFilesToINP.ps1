#requires -version 5.1
<#
.SYNOPSIS
    Copies all files with the .ras extension from a source network path to a destination network path.
.DESCRIPTION
    This script is designed to run as a scheduled task. It checks for .ras files
    in the source directory, COPIES them (leaving the originals in place) to the
    destination directory, and logs all operations to a dedicated log file.
.NOTES
    Author: Gemini
    Date: 10/06/2025
    Action: Copy files instead of Move.
    Log File: MoveRasFilesToINP.log
#>

# === CONFIGURATION ===
$SourcePath      = "\\fileserver\UPLATNI RACUNI\000"
$DestinationPath = "\\batch04-08-new\inp"
# Log file saved in the source directory (\\fileserver\UPLATNI RACUNI\000)
$LogFileName     = "MoveRasFilesToINP.log"
$LogFilePath     = Join-Path $SourcePath $LogFileName
$FileExtension   = "*.ras"

# === HELPER FUNCTION: Logging ===
Function Write-Log {
    Param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "$Timestamp $Message"
    # Output to console (for immediate review) and the log file
    Write-Host $Line
    Add-Content -Path $LogFilePath -Value $Line -Encoding UTF8
}

# Clear the log file at the start of the run
"" | Out-File -FilePath $LogFilePath -Encoding UTF8 -Force

# === SCRIPT START ===
Write-Log "=== RAS File Copier Started ==="
Write-Log "Source: $SourcePath"
Write-Log "Destination: $DestinationPath"
Write-Log "Log file: $LogFilePath"

# --- 1. Validation ---
if (!(Test-Path -Path $SourcePath -PathType Container)) {
    Write-Log "ERROR: Source path '$SourcePath' not found or is inaccessible. Exiting."
    Exit 1
}

if (!(Test-Path -Path $DestinationPath -PathType Container)) {
    Write-Log "ERROR: Destination path '$DestinationPath' not found or is inaccessible. Attempting to create it."
    try {
        # Attempt to create the destination path if it doesn't exist (assuming necessary permissions)
        New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        Write-Log "SUCCESS: Destination path created."
    }
    catch {
        Write-Log "FATAL ERROR: Could not create destination path '$DestinationPath'. Details: $_. Exiting."
        Exit 1
    }
}

# --- 2. Find and Copy Files ---
try {
    # Get files matching the extension
    $FilesToCopy = Get-ChildItem -Path $SourcePath -Filter $FileExtension -File

    if ($FilesToCopy.Count -eq 0) {
        Write-Log "INFO: No '$FileExtension' files found in '$SourcePath'."
    } else {
        Write-Log "INFO: Found $($FilesToCopy.Count) file(s) to copy."
        
        $CopiedCount = 0
        foreach ($File in $FilesToCopy) {
            $DestFileName = Join-Path -Path $DestinationPath -ChildPath $File.Name
            
            # Use Copy-Item instead of Move-Item
            # Use -Force to overwrite if a file with the same name already exists in the destination
            Copy-Item -Path $File.FullName -Destination $DestFileName -Force -ErrorAction Stop
            
            Write-Log "Copied: $($File.Name) -> $($DestinationPath)"
            $CopiedCount++
        }
        Write-Log "SUCCESS: Successfully copied $CopiedCount out of $($FilesToCopy.Count) files."
    }
}
catch {
    Write-Log "ERROR during file copying: $_"
    # Log the general error
}

# === SCRIPT END ===
Write-Log "=== RAS File Copier Finished ==="
