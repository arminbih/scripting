#requires -version 5.1
<#
.SYNOPSIS
    Connects to an Outlook folder, filters for file attachments, processes them (image-to-PDF, PDF splitting),
    and marks processed emails.
.DESCRIPTION
    Version 6.6 addresses the file deletion issue after PDF splitting by implementing a robust retry loop.
    It also introduces a shorter, clearer file naming convention for better file management.
.NOTES
    Author: Armin Becirspahic (Updated by Gemini)
    Date: 10/23/2025
    Version: 6.6 - DELETION FIX & NAMING REVISION:
    1. Robust file deletion with retry logic added to handle file lock releases from qpdf.
    2. New, shorter file naming convention: YYYYMMDD_HHMMSS_SenderName_OriginalFileName.pdf.
    3. Single-page PDFs now include '_0001' for naming consistency.
#>

# === CONFIGURATION ===
$MailboxName            = "armin.becirspahic@fpu.gov.ba"        # Mailbox root
$OutlookFolderName      = "test"                               # Target Outlook folder
$BaseSavePath           = "D:\uredbaTest"                      # Base directory where run folders will be created

$SkipDuplicates         = $true
$AllowedExtensions      = @(".pdf", ".jpg", ".png", ".jpeg")
$PictureExtensions      = @(".jpg", ".png", ".jpeg")
# The known forwarding account to exclude from naming and harvesting
$ForwarderEmail         = "uredba450@fpu.gov.ba"
$ForwarderEmailDomain   = $ForwarderEmail.Split('@')[1]

# --- External Tool Paths (CONFIGURED) ---
$ImageToPdfToolExe      = "C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe"
$PdfSplitterToolExe     = "C:\Program Files\qpdf 12.2.0\bin\qpdf.exe"

# --- Outlook Constants for Filtering ---
$olByValue              = 1 # Attachment is copied to disk
$olEmbeddedItem         = 5 # Attachment is embedded in the body (e.g., signature image)

# === SETUP UNIQUE RUN FOLDER ===
$RunTimestamp = (Get-Date -Format "yyyyMMddHHmmss")
$SavePath = Join-Path $BaseSavePath $RunTimestamp

# Ensure the Base path exists, then create the unique run path
If (!(Test-Path -Path $BaseSavePath)) {
    Write-Host "Creating Base Save Path: $BaseSavePath"
    New-Item -ItemType Directory -Path $BaseSavePath | Out-Null
}
If (!(Test-Path -Path $SavePath)) {
    Write-Host "Creating Unique Run Folder: $SavePath"
    New-Item -ItemType Directory -Path $SavePath | Out-Null
}

$LogFile = Join-Path $SavePath "GetFromOutlook.log"
"" | Out-File -FilePath $LogFile -Encoding UTF8 -Force

Function Write-Log {
    Param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "$Timestamp $Message"
    Add-Content -Path $LogFile -Value $Line
    Write-Output $Line | Out-Host
}

# === SENDER EXTRACTION AND HARVESTING FUNCTIONS (No changes needed here) ===

# Global array to collect all candidate emails during the lookup process
$script:CandidateEmails = @()

# Helper function to check for and reject known non-email formats
Function Is-ExchangeDN {
    Param([string]$Value)
    if ($Value -match '^(?:/o=.*OU=Exchange Administrative Group|cn=Recipients|EX\\)') {
        Write-Log "  -> [DN REJECTED] Value matches Exchange DN pattern: $($Value.Substring(0, [System.Math]::Min(30, $Value.Length)))..." | Out-Null
        return $true
    }
    return $false
}

# Helper function to add emails to the global list if they are unique and not the forwarder
Function Add-CandidateEmail {
    Param([string]$Email)
    $CleanEmail = $Email.Trim().ToLower()
    
    if (Is-ExchangeDN -Value $CleanEmail) { return }
    if ($CleanEmail -eq $script:ForwarderEmail -or $CleanEmail -like "*$script:ForwarderEmailDomain") { return }
    
    if ($CleanEmail -ne "" -and ($CleanEmail -match '@') -and -not ($script:CandidateEmails -contains $CleanEmail)) {
        $script:CandidateEmails += $CleanEmail
        Write-Log "  -> [HARVEST] Added candidate: $($CleanEmail)" | Out-Null
    }
}

# Tier 1: MAPI Property (Primary)
Function Get-OriginalSenderFromMAPI {
    Param([Microsoft.Office.Interop.Outlook.MailItem]$MailItem)
    $PR_SENT_REPRESENTING_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
    try {
        $OriginalValue = $MailItem.PropertyAccessor.GetProperty($PR_SENT_REPRESENTING_EMAIL_ADDRESS)
        if ($OriginalValue -ne $null -and $OriginalValue -ne "") {
            $OriginalEmail = $OriginalValue.Trim()
            Add-CandidateEmail -Email $OriginalEmail
            if ($OriginalEmail -match '@' -and -not (Is-ExchangeDN -Value $OriginalEmail)) {
                Write-Log "  -> [TIER 1 SUCCESS] Extracted original sender from MAPI: $($OriginalEmail)" | Out-Null
                return $OriginalEmail
            }
        }
    }
    catch {
        Write-Log "  -> [TIER 1 FAIL] MAPI property access failed." | Out-Null
    }
    return $null
}

# Tier 2: Mail Body/HTML Parsing (Structured)
Function Get-OriginalSenderFromBody {
    Param([Microsoft.Office.Interop.Outlook.MailItem]$MailItem)
    $EmailRegex = 'From:\s+.*<([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})>'
    $Content = "$($MailItem.Body)$($MailItem.HTMLBody)"
    if ($Content -match $EmailRegex) {
        $OriginalEmail = $Matches[1].Trim()
        Add-CandidateEmail -Email $OriginalEmail
        if ($OriginalEmail -match '@') {
            Write-Log "  -> [TIER 2 SUCCESS] Extracted original sender from Body: $($OriginalEmail)" | Out-Null
            return $OriginalEmail
        }
    }
    Write-Log "  -> [TIER 2 FAIL] Mail Body parsing failed." | Out-Null
    return $null
}

# Tier 3: Raw Transport Headers Parsing
Function Get-OriginalSenderFromHeaders {
    Param([Microsoft.Office.Interop.Outlook.MailItem]$MailItem)
    $PR_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
    $EmailHeaderRegex = '^(?:Return-Path|Original-Recipient):\s*(?:<)?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})(?:>)?'
    
    try {
        $Headers = $MailItem.PropertyAccessor.GetProperty($PR_TRANSPORT_HEADERS)
        if ($Headers -match $EmailHeaderRegex) {
            $OriginalEmail = $Matches[1].Trim()
            Add-CandidateEmail -Email $OriginalEmail
            if ($OriginalEmail -match '@') {
                Write-Log "  -> [TIER 3 SUCCESS] Extracted original sender from Headers: $($OriginalEmail)" | Out-Null
                return $OriginalEmail
            }
        }
    }
    catch {
        Write-Log "  -> [TIER 3 FAIL] Raw Headers property access failed." | Out-Null
    }
    return $null
}

# Tier 4: Comprehensive Email Harvesting from Body
Function Harvest-AllEmailsFromBody {
    Param([Microsoft.Office.Interop.Outlook.MailItem]$MailItem)
    $FullEmailRegex = '([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6})'
    $Content = "$($MailItem.Body)$($MailItem.HTMLBody)"
    
    $AllMatches = [regex]::Matches($Content, $FullEmailRegex)
    foreach ($Match in $AllMatches) {
        $Email = $Match.Groups[1].Value.Trim()
        Add-CandidateEmail -Email $Email
    }
    
    if ($script:CandidateEmails.Count -gt 0) {
        Write-Log "  -> [TIER 4 HARVEST] Total unique non-forwarder candidates harvested: $($script:CandidateEmails.Count)." | Out-Null
    }
    return $null 
}


# === PDF/IMAGE UTILITY FUNCTIONS ===

Function Get-PdfPageCount {
    Param([string]$PdfPath)
    if (-not (Test-Path $PdfPath)) { Write-Log "  -> [qpdf] PDF file not found for page count check: $PdfPath"; return 0 }
    try {
        # Use -ErrorAction SilentlyContinue to suppress qpdf warnings that go to stderr
        $qpdfOutput = & $PdfSplitterToolExe --json $PdfPath 2>$null | Out-String
        $json = $qpdfOutput | ConvertFrom-Json
        $pageCount = $json.pages.Count
        Write-Log "  -> [qpdf] Detected $pageCount page(s) in $PdfPath" | Out-Null
        return $pageCount
    }
    catch {
        Write-Log "  -> [ERROR] Failed to get page count using qpdf for $PdfPath. Details: $($_.Exception.Message)" | Out-Null
        return 1 # Assume 1 page to prevent splitting if count fails
    }
}

Function Split-PdfFile {
    Param([string]$InputPdfPath)
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPdfPath)
    $SaveDir = [System.IO.Path]::GetDirectoryName($InputPdfPath)
    
    # The base name that appears in the split file names, before the _%04d.pdf
    $OutputTemplateBase = "$($BaseName)_"
    $OutputTemplate = Join-Path $SaveDir "$($OutputTemplateBase)%04d.pdf"
    # Arguments passed as an array to handle paths with spaces/special characters safely
    $args = @( "`"$InputPdfPath`"", '--split-pages', "`"$OutputTemplate`"" )
    
    Write-Log "  -> [ACTION] Splitting multi-page PDF: $($InputPdfPath)" | Out-Null
    
    try {
        # Execute qpdf
        $proc = Start-Process -FilePath $PdfSplitterToolExe -ArgumentList $args -Wait -NoNewWindow -PassThru -ErrorAction Stop
        
        $isSuccessful = $false
        
        if ($proc.ExitCode -eq 0) {
            $isSuccessful = $true
        } else {
            # --- WILDCARD FIX (V6.5) ---
            # 1. Escape brackets in the base name so PowerShell treats them literally, not as a character set.
            $EscapedBaseName = $OutputTemplateBase.Replace('[','`[').Replace(']','`]')
            
            # 2. Check for output files using the escaped name for the filter.
            $OutputFiles = Get-ChildItem -Path $SaveDir -Filter "$($EscapedBaseName)*.pdf" | Where-Object { 
                # This filter checks that the name starts with the escaped base and is not the original unsplit file.
                $_.Name -like "$($EscapedBaseName)*.pdf" -and $_.Name -ne "$($BaseName).pdf" 
            }
            
            if ($OutputFiles.Count -gt 0) {
                # QPDF reported error (e.g., Code 3 for file warning), but files were written.
                Write-Log "  -> [INFO] QPDF exited with code $($proc.ExitCode), but $($OutputFiles.Count) split files were found. Assuming non-fatal error." | Out-Null
                $isSuccessful = $true
            } else {
                Write-Log "  -> [FAILED] PDF splitting failed for: $($InputPdfPath). qpdf Exit Code: $($proc.ExitCode). No output files found." | Out-Null
            }
            # --- END WILDCARD FIX ---
        }

        if ($isSuccessful) {
            Write-Log "  -> [SUCCESS] PDF split done. Files saved as: $($BaseName)_XXXX.pdf" | Out-Null
            # Simple cleanup: Try to remove the input file immediately. The main loop has the robust retry.
            Remove-Item $InputPdfPath -Force -ErrorAction SilentlyContinue 
            return $true
        } else {
            return $false
        } 
    }
    catch {
        # This catches errors *before* qpdf executes (e.g., path error) or if Start-Process itself fails.
        Write-Log "  -> [FATAL ERROR] Exception during qpdf execution for $($InputPdfPath). Details: $($_.Exception.Message)" | Out-Null
        return $false
    }
}

Function Convert-ImageToPDF {
    Param([string]$ImagePath)
    $PdfPath = $ImagePath.Replace([System.IO.Path]::GetExtension($ImagePath), ".pdf")
    $args = @("-density", "300", "`"$ImagePath`"", "`"$PdfPath`"")
    
    Write-Log "  -> [ACTION] Converting image to PDF: $($ImagePath) -> $($PdfPath)" | Out-Null
    
    try {
        $proc = Start-Process -FilePath $ImageToPdfToolExe -ArgumentList $args -Wait -NoNewWindow -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0) {
            Write-Log "  -> [SUCCESS] Converted image to PDF: $($PdfPath)" | Out-Null
            Remove-Item $ImagePath -Force -ErrorAction SilentlyContinue
            return $PdfPath
        } else {
            Write-Log "  -> [FAILED] Image to PDF conversion failed for: $($ImagePath). Magick Exit Code: $($proc.ExitCode)" | Out-Null
            Remove-Item $PdfPath -Force -ErrorAction SilentlyContinue
            return $null
        }
    }
    catch {
        Write-Log "  -> [FATAL ERROR] Exception during Magick execution for $($ImagePath). Details: $($_.Exception.Message)" | Out-Null
        return $null
    }
}

# === RETRY DELETION FUNCTION (NEW) ===
Function Remove-ItemWithRetry {
    Param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$MaxRetries = 5,
        [int]$DelayMS = 200
    )
    
    $Attempt = 0
    do {
        $Attempt++
        if ($Attempt -gt 1) {
            Write-Log "    -> [RETRY DELETE] Waiting $($DelayMS)ms to retry deletion of: $Path (Attempt $Attempt/$MaxRetries)" | Out-Null
            Start-Sleep -Milliseconds $DelayMS
        }
        
        try {
            Remove-Item $Path -Force -ErrorAction Stop
            Write-Log "    -> [SUCCESS DELETE] Removed file: $Path" | Out-Null
            return $true
        }
        catch {
            Write-Log "    -> [FAILED DELETE] Error deleting file: $Path. Details: $($_.Exception.Message)" | Out-Null
        }
    } while ($Attempt -lt $MaxRetries)
    
    Write-Log "    -> [FATAL] Failed to delete file $Path after $MaxRetries attempts. File may be locked." | Out-Null
    return $false
}
# === END RETRY DELETION FUNCTION ===

# === SCRIPT START ===
Write-Log "=== Run started ===" | Out-Null
Write-Log "Target folder: $OutlookFolderName | Base path: $BaseSavePath | Run Folder: $SavePath" | Out-Null

# Connect to Outlook
$Outlook    = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Get target folder
$Mailbox = $Namespace.Folders.Item($MailboxName)
If (-not $Mailbox) { Write-Log "ERROR: Mailbox '$MailboxName' not found."; Exit }
$TargetFolder = $Mailbox.Folders.Item($OutlookFolderName)
If (-not $TargetFolder) { Write-Log "ERROR: Folder '$OutlookFolderName' not found."; Exit }

# Counters
$MailCount = 0
$AttachSaved = 0

# Loop unread mails
foreach ($Mail in $TargetFolder.Items) {
    $SavedAttachmentsInMail = $false 
    $script:CandidateEmails = @() # RESET candidate list for each new mail
    
    try { 
        if ($Mail -is [Microsoft.Office.Interop.Outlook.MailItem] -and $Mail.UnRead -eq $true -and $Mail.Attachments.Count -gt 0) {
            $MailCount++
            Write-Log "[Mail] Subject='$($Mail.Subject)' Received=$($Mail.ReceivedTime)" | Out-Null

            $SenderEmail = $Mail.SenderEmailAddress
            $BestCandidateEmail = $null 
            $TempRenamedPath = $null # Initialize here for cleanup

            Add-CandidateEmail -Email $SenderEmail
            
            # --- SENDER IDENTIFICATION LOGIC (Unchanged) ---
            if ($SenderEmail -ceq $ForwarderEmail) {
                Write-Log "  -> [INFO] Detected Forwarder ($ForwarderEmail). Starting 4-Tier lookup..." | Out-Null
                
                $BestCandidateEmail = Get-OriginalSenderFromMAPI -MailItem $Mail
                
                if ($BestCandidateEmail -eq $null) {
                    $BestCandidateEmail = Get-OriginalSenderFromBody -MailItem $Mail
                }
                
                if ($BestCandidateEmail -eq $null) {
                    $BestCandidateEmail = Get-OriginalSenderFromHeaders -MailItem $Mail
                }
                
                Harvest-AllEmailsFromBody -MailItem $Mail
            }

            # --- FILENAME PREFIX CREATION (REVISED FOR SHORT NAMING) ---

            $NameSource = $null

            if ($script:CandidateEmails.Count -gt 0) {
                # Use all unique, clean candidate names for the filename
                $CleanedCandidates = $script:CandidateEmails | ForEach-Object {
                    if ($_ -match '^([a-zA-Z0-9._%+-]+)@') {
                        $Matches[1]
                    } else {
                        $_
                    }
                } | Select-Object -Unique

                $ConcatenatedNames = $CleanedCandidates -join '__'
                $NameSource = $ConcatenatedNames
                Write-Log "  -> [FINAL PREFIX] Using HARVESTED candidates: $($NameSource)" | Out-Null
                $AllEmailsFound = $script:CandidateEmails -join ", "
                Write-Log "  -> [HARVESTED EMAILS]: $($AllEmailsFound)" | Out-Null

            } elseif ($Mail.SenderName -ne $null -and -not (Is-ExchangeDN -Value $Mail.SenderName)) {
                # Fallback to the clean, human-readable Sender Name
                $NameSource = $Mail.SenderName
                Write-Log "  -> [FINAL PREFIX] Using CLEAN SENDER NAME: $($NameSource)" | Out-Null
                $AllEmailsFound = $script:CandidateEmails -join ", "
                if ($AllEmailsFound -ne "") {
                    Write-Log "  -> [HARVESTED EMAILS]: $($AllEmailsFound) (Sender Name used as prefix fallback)" | Out-Null
                }

            } else {
                # FINAL FALLBACK: Subject
                $NameSource = $Mail.Subject -replace '^(FW|RE|Fwd|Odgovor):\s*' -replace '\s+', ' '
                Write-Log "  -> [WARNING] All lookups failed. Falling back to sanitized Subject." | Out-Null
                $AllEmailsFound = $script:CandidateEmails -join ", "
                if ($AllEmailsFound -ne "") {
                    Write-Log "  -> [HARVESTED EMAILS]: $($AllEmailsFound) (Subject used as prefix fallback)" | Out-Null
                }
            }
            
            # 1. Create the received timestamp prefix (YYYYMMDD_HHMMSS)
            $ReceivedTimestamp = $Mail.ReceivedTime.ToString("yyyyMMdd_HHmmss")
            Write-Log "  -> [TIMESTAMP] Using Received Time: $($ReceivedTimestamp)" | Out-Null

            # 2. Sanitize Sender/NameSource
            $SanitizedSenderName = ($NameSource -replace '[^a-zA-Z0-9\._-]' , '_')
            
            # 3. Handle very short/empty sender name
            if ($SanitizedSenderName.Length -lt 3) {
                $SanitizedSenderName = "UNKNOWN"
                Write-Log "  -> [WARNING] Sanitized sender name was too short; defaulting to UNKNOWN." | Out-Null
            }
            
            # 4. Final file prefix creation (YYYYMMDD_HHMMSS_SenderName_)
            $SanitizedPrefix = "$($ReceivedTimestamp)_$($SanitizedSenderName)_"
            # --- END FILENAME PREFIX CREATION ---

            
            foreach ($Att in $Mail.Attachments) {
                # Attachment Filtering (Retained)
                if ($Att.Type -eq $olEmbeddedItem) { Write-Log "  -> Attachment: $($Att.FileName) -> Skipped (Type: Embedded item)"; continue }
                if ($Att.FileName -match "^(image|clip_image)\d+\.(jpg|jpeg|png|gif|bmp)$") { Write-Log "  -> Attachment: $($Att.FileName) -> Skipped (Generic embedded image name filter)"; continue }
                $Ext = [System.IO.Path]::GetExtension($Att.FileName).ToLower()

                if ($AllowedExtensions -contains $Ext) {
                    $OriginalSafeFileName = $Att.FileName -replace '[^a-zA-Z0-9\._-]' , '_' -replace '\s+', '_'
                    $FileName = Join-Path $SavePath $OriginalSafeFileName 
                    
                    if ($SkipDuplicates -and (Test-Path $FileName)) { Write-Log "  -> Attachment: $($Att.FileName) -> Skipped duplicate"; continue }
                    
                    try {
                        $CurrentPdfPath = $null
                        # Save attachment
                        $Att.SaveAsFile($FileName)
                        Write-Log "  -> Attachment: $($Att.FileName) -> Saved to $FileName" | Out-Null
                        $AttachSaved++

                        # Conversion and Processing Logic
                        if ($PictureExtensions -contains $Ext) {
                            $CurrentPdfPath = Convert-ImageToPDF -ImagePath $FileName
                        } elseif ($Ext -eq ".pdf") {
                            $CurrentPdfPath = $FileName
                        }

                        if ($CurrentPdfPath -ne $null -and (Test-Path $CurrentPdfPath)) {
                            $PageCount = Get-PdfPageCount -PdfPath $CurrentPdfPath
                            
                            if ($PageCount -gt 1) {
                                $SplitBaseName = [System.IO.Path]::GetFileNameWithoutExtension($CurrentPdfPath)
                                
                                # This creates the temporary, unsplit file that Split-PdfFile will attempt to delete
                                $TempRenamedPath = Join-Path $SavePath ($SanitizedPrefix + $SplitBaseName + ".pdf")
                                
                                Rename-Item -Path $CurrentPdfPath -NewName $TempRenamedPath -Force
                                
                                $SplitSuccess = Split-PdfFile -InputPdfPath $TempRenamedPath 
                                
                                # --- DELETION FIX: Delete temporary renamed file with retry ---
                                if ($SplitSuccess -and (Test-Path $TempRenamedPath)) {
                                    Remove-ItemWithRetry -Path $TempRenamedPath
                                }
                                # --- END DELETION FIX ---
                                
                                $SavedAttachmentsInMail = $true

                            } elseif ($PageCount -eq 1) {
                                Write-Log "  -> PDF is single-page. Renaming with unique datetime stamp and sender info." | Out-Null
                                
                                $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($CurrentPdfPath)
                                $i = 0
                                
                                do {
                                    $i++
                                    # New file name structure: YYYYMMDD_HHMMSS_SenderName_ORIGINALFILENAME_0001[_01].pdf
                                    # The $i count is only for collision, the _0001 indicates it's the first and only page
                                    $Suffix = if ($i -gt 1) { "_$($i.ToString('02'))" } else { "" } # Only add suffix for collision avoidance
                                    $NewPdfName = $SanitizedPrefix + $BaseName + "_0001" + $Suffix + ".pdf"
                                    $NewPdfPath = Join-Path $SavePath $NewPdfName
                                    if ($i -gt 99) { throw "File naming conflict: Could not find unique name for $($BaseName)." }
                                } while (Test-Path $NewPdfPath)

                                Rename-Item -Path $CurrentPdfPath -NewName $NewPdfName -Force
                                $CurrentPdfPath = $NewPdfPath 

                                $SavedAttachmentsInMail = $true
                            }
                        }
                        
                    }
                    catch {
                        Write-Log "  -> [FATAL ERROR: ATTACHMENT PROCESSING] Failed to process attachment $($Att.FileName). Details: $($_.Exception.Message)" | Out-Null
                        # Enhanced cleanup in case of failure
                        if (Test-Path $FileName) { Write-Log "  -> [CLEANUP] Removing incomplete original file: $FileName"; Remove-ItemWithRetry -Path $FileName }
                        if ($CurrentPdfPath -ne $null -and (Test-Path $CurrentPdfPath)) { Write-Log "  -> [CLEANUP] Removing incomplete PDF file: $CurrentPdfPath"; Remove-ItemWithRetry -Path $CurrentPdfPath }
                        if ($TempRenamedPath -ne $null -and (Test-Path $TempRenamedPath)) { Write-Log "  -> [CLEANUP] Removing temporary renamed file: $TempRenamedPath"; Remove-ItemWithRetry -Path $TempRenamedPath }
                        continue
                    }
                }
                else {
                    Write-Log "  -> Attachment: $($Att.FileName) -> Skipped (extension not allowed)" | Out-Null
                }
            }

            # Mark mail read if attachments were saved/processed
            if ($SavedAttachmentsInMail) {
                $Mail.UnRead = $false
                Write-Log "  -> Mail marked as read (processed)." | Out-Null
            }
        }
    }
    catch {
        Write-Log "ERROR processing mail '$($Mail.Subject)': $($_.Exception.Message)" | Out-Null
    }
}

# === SCRIPT END ===
Write-Log "Summary: Processed $MailCount mails, $AttachSaved saved/processed." | Out-Null

# Clean up COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()