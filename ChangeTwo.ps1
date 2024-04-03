# Specify the path to your text file
$filePath = "C:\Users\armin.becirspahic\Desktop\ras poredjenje\55508824_ispravka_ps.RAS"



# Read the content of the file
$fileContent = Get-Content $filePath

# Create an array to store modified lines
$modifiedContent = @()

# Loop through each line in the file
foreach ($line in $fileContent) {
    # Check if the line has at least 119 characters
    if ($line.Length -ge 119) {
        # Modify the 119th character to '2'
        $modifiedLine = $line.Substring(0, 118) + "2" + $line.Substring(119)
        
        # Add the modified line to the array
        $modifiedContent += $modifiedLine
    }
    else {
        # Add the original line if it's shorter than 119 characters
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
