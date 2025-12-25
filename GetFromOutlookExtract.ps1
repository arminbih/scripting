#requires -version 5.1
<#
.SYNOPSIS
    Connects to an Outlook folder, saves attachments from unread emails,
    and selectively extracts files from archives (.zip, .7z) and nested .msg attachments.
.DESCRIPTION
    This script automates the processing of emails from a specific Outlook folder.
    It saves allowed attachment types, extracts compressed archives using 7-Zip,
    and uses a recursive function to open .msg attachments, extracting only
    specific file types (.ras, .7z, .zip) from within them.
    *** MODIFIED to extract archive contents directly to the $SavePath folder. ***
.NOTES
    Author: Armin Becirspahic (Updated by Gemini)
    Date: 10/06/2025
    Version: 2.2 - Modified archive extraction to avoid creating new folders.
#>

# === CONFIGURATION ===
$MailboxName             = "armin.becirspahic@fpu.gov.ba"      # Mailbox root
$OutlookFolderName       = "RAS FILES"                         # Target folder under mailbox root
$SavePath                = "\\fileserver\UPLATNI RACUNI\000"   # Where to save attachments & logs
$MarkAsRead              = $true                                # Mark mails as read after processing?
$SkipDuplicates          = $true                                # Skip if file already exists
$AllowedExtensions       = @(".ras", ".7z", ".zip", ".msg")     # Allowed initial attachment types
$AllowedNestedExtensions = @(".ras", ".7z", ".zip")            # Allowed types inside a .msg
$SevenZipExe             = "C:\Program Files\7-Zip\7z.exe"     # Path to 7-Zip executable

# === PREPARE LOG FILE ===
If (!(Test-Path -Path $SavePath)) {
    New-Item -ItemType Directory -Path $SavePath | Out-Null
}
$LogFile = Join-Path $SavePath "GetFromOutlook.log"

# Overwrite the log file at the start of the run
"" | Out-File -FilePath $LogFile -Encoding UTF8 -Force

Function Write-Log {
    Param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "$Timestamp $Message"
    Add-Content -Path $LogFile -Value $Line
    Write-Output $Line
}

# === SCRIPT START ===
Write-Log "=== Run started ==="
Write-Log "Mailbox: $MailboxName"
Write-Log "Target folder: $OutlookFolderName"
Write-Log "Save path: $SavePath"
Write-Log "Allowed extensions: $($AllowedExtensions -join ', ')"
Write-Log "Allowed NESTED extensions: $($AllowedNestedExtensions -join ', ')"

# Connect to Outlook
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# --- Recursive function to process .msg files ---
Function Process-MsgFile {
    Param(
        [string]$MsgFilePath,
        [string]$DestinationPath,
        [int]$Depth = 1
    )

    $Indent = "  " * $Depth
    Write-Log "$($Indent)-> Processing .msg file: $(Split-Path $MsgFilePath -Leaf)"

    try {
        $msg = $Namespace.OpenSharedItem($MsgFilePath)
        if ($msg.Attachments.Count -gt 0) {
            foreach ($att in $msg.Attachments) {
                $nestedFileName = Join-Path $DestinationPath $att.FileName
                $nestedExt = [System.IO.Path]::GetExtension($att.FileName).ToLower()

                # MODIFIED: Check if the extension is allowed for nested files OR if it's another .msg for recursion
                if ($AllowedNestedExtensions -contains $nestedExt -or $nestedExt -eq ".msg") {
                    if ($SkipDuplicates -and (Test-Path $nestedFileName)) {
                        Write-Log "$($Indent)  -> Attachment: $($att.FileName) -> Skipped duplicate"
                        continue
                    }

                    $att.SaveAsFile($nestedFileName)
                    Write-Log "$($Indent)  -> Attachment: $($att.FileName) -> Saved to $nestedFileName"
                    $global:AttachSaved++

                    # If the nested attachment is another .msg, recurse
                    if ($nestedExt -eq ".msg") {
                        Process-MsgFile -MsgFilePath $nestedFileName -DestinationPath $DestinationPath -Depth ($Depth + 1)
                    }
                }
                else {
                    # Log that the file was skipped because its extension is not allowed
                    Write-Log "$($Indent)  -> Attachment: $($att.FileName) -> Skipped (extension not allowed for nested extraction)"
                    $global:AttachFiltered++
                }
            }
        }
        else {
            Write-Log "$($Indent)  -> No attachments found in this .msg file."
        }
    }
    catch {
        Write-Log "$($Indent)ERROR: Could not process .msg file '$MsgFilePath'. Details: $_"
    }
    finally {
        # FIX: Release msg object
        if ($msg) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null }
    }
}


# Get mailbox root
$Mailbox = $Namespace.Folders.Item($MailboxName)
If (-not $Mailbox) {
    Write-Log "ERROR: Mailbox '$MailboxName' not found."
    Exit
}

# Target folder directly under mailbox root
$TargetFolder = $Mailbox.Folders.Item($OutlookFolderName)
If (-not $TargetFolder) {
    Write-Log "ERROR: Folder '$OutlookFolderName' not found under mailbox root '$MailboxName'."
    Exit
}

# FIX: Store items in array to break COM enumeration
$Items = @($TargetFolder.Items)

# Counters
$MailCount = 0
$AttachSaved = 0
$AttachSkipped = 0
$AttachFiltered = 0
$ArchivesProcessed = 0
$ArchivesFailed = 0

# Loop unread mails
foreach ($MailItem in $Items) {
    $Mail = $null
    try {
        $Mail = $MailItem
        if ($Mail.UnRead -eq $true -and $Mail.Attachments.Count -gt 0) {
            $MailCount++
            Write-Log "[Mail] Subject='$($Mail.Subject)' Received=$($Mail.ReceivedTime)"

            foreach ($Att in $Mail.Attachments) {
                $Ext = [System.IO.Path]::GetExtension($Att.FileName).ToLower()

                if ($AllowedExtensions -contains $Ext) {
                    $FileName = Join-Path $SavePath $Att.FileName

                    if ($SkipDuplicates -and (Test-Path $FileName)) {
                        Write-Log "  -> Attachment: $($Att.FileName) -> Skipped duplicate"
                        $AttachSkipped++
                        continue
                    }

                    # Save attachment
                    $Att.SaveAsFile($FileName)
                    Write-Log "  -> Attachment: $($Att.FileName) -> Saved to $FileName"
                    $AttachSaved++

                    # === Archive extraction ===
                    if ($Ext -in @(".zip", ".7z")) {
                        $Password = ""
                        if ($Att.FileName.StartsWith("199")) { $Password = "ATIDEONID" }
                        elseif ($Att.FileName.StartsWith("306")) { $Password = "hymodr" }
                        elseif ($Att.FileName.StartsWith("132")) { $Password = "TUZB99" }
                        
                        $DestFolder = $SavePath
                        $args = @("x", "`"$FileName`"", "-o`"$DestFolder`"", "-y")
                        if ($Password) { $args += "-p$Password" }
                        
                        Write-Log "    -> Extracting archive $($Att.FileName) to $DestFolder ..."
                        $proc = Start-Process -FilePath $SevenZipExe -ArgumentList $args -Wait -NoNewWindow -PassThru
                        if ($proc.ExitCode -eq 0) {
                            Write-Log "    -> Extraction successful: Contents in $DestFolder"
                            $ArchivesProcessed++
                        } else {
                            Write-Log "    -> Extraction FAILED for $($Att.FileName)"
                            $ArchivesFailed++
                        }
                    }
                    # === NESTED .MSG EXTRACTION ===
                    elseif ($Ext -eq ".msg") {
                        Process-MsgFile -MsgFilePath $FileName -DestinationPath $SavePath
                    }
                }
                else {
                    Write-Log "  -> Attachment: $($Att.FileName) -> Skipped (extension not allowed)"
                    $AttachFiltered++
                }
            }

            if ($MarkAsRead) { 
                $Mail.UnRead = $false
                Write-Log "  -> Mail marked as read"
            }
        }
    }
    catch {
        Write-Log "ERROR processing mail '$($Mail.Subject)': $_"
    }
    finally {
        # FIX: Release mail object
        if ($Mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null }
    }
}

# === SCRIPT END ===
Write-Log "Summary: Processed $MailCount mails, $AttachSaved saved, $AttachSkipped skipped duplicates, $AttachFiltered filtered out."
Write-Log "Archives extracted successfully: $ArchivesProcessed, Failed: $ArchivesFailed"
Write-Log "=== Run finished ==="

# FIX: Clear array and release all objects
$Items = $null
if ($TargetFolder) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($TargetFolder) | Out-Null }
if ($Mailbox) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mailbox) | Out-Null }
if ($Namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null }
if ($Outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null }

# FIX: Force cleanup and exit
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Start-Sleep -Seconds 5
Exit