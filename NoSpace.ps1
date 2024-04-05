# Specify the path to your text file
$filePath = "PATH"

# Read the content of the file
$fileContent = Get-Content $filePath

# Create an array to store modified lines
$modifiedContent = @()

# Loop through each line in the file
foreach ($line in $fileContent) {
    # Check if the line has exactly 35 characters (excluding spaces)
    if (($line.Trim() -replace '\s+$').Length -eq 35) {
        # Trim trailing whitespaces
        $modifiedLine = $line.TrimEnd()
        
        # Add the modified line to the array
        $modifiedContent += $modifiedLine
    }
    else {
        # Add the original line
        $modifiedContent += $line
    }
}

# Output the modified content to the console
$modifiedContent

# Prompt the user to confirm saving the changes to the file
$confirm = Read-Host "Do you want to save the changes to the file? (Y/N)"
if ($confirm -eq "Y" -or $confirm -eq "y") {
    # Save the modified content back to the file
    $modifiedContent | Set-Content $filePath
    Write-Host "Changes saved to $filePath."
}
else {
    Write-Host "Changes not saved."
}
