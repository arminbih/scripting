# Set the main folder path
$mainFolderPath = "PATH"

# Set the subfolder path where files are located
$subFolderPath = Join-Path -Path $mainFolderPath -ChildPath "000"

# Get the list of files in the subfolder "000"
$files = Get-ChildItem -Path $subFolderPath -File

# Set the log file path
$logFilePath = Join-Path -Path $subFolderPath -ChildPath "MoveFilesLog.txt"

# Clear existing log or create a new log file
$null | Out-File -FilePath $logFilePath

foreach ($file in $files) {
    # Extract the first three characters from the file name
    $prefix = $file.BaseName.Substring(0, 3)
    
    # Construct the folder path
    $folderPath = Join-Path -Path $mainFolderPath -ChildPath $prefix
    
    if (Test-Path $folderPath) {
        # Move the file to the corresponding folder
        $destination = Join-Path -Path $folderPath -ChildPath $file.Name
        Move-Item -Path $file.FullName -Destination $destination

        # Log the move operation
        $logEntry = "Moved $($file.Name) to $($folderPath)"
        Add-Content -Path $logFilePath -Value $logEntry
    } else {
        # Log that the file was not moved due to the folder not existing
        $logEntry = "NotMoved $($file.Name) - Folder $($folderPath) does not exist."
        Add-Content -Path $logFilePath -Value $logEntry
    }
}

Write-Host "Files moved successfully."
Write-Host "Log file created at: $logFilePath"
Read-Host -Prompt "Press Enter to exit"
