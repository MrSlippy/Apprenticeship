# ============================================================================
# PHISHING EMAIL PROCESSOR 
# ============================================================================

#Requires -Version 5.1

# ============================================================================
# OUTLOOK CONSTANTS
# ============================================================================
$OlConstants = @{
    # Item classes
    olMailItem              = 43
    
    # Inspector close modes
    olDiscard               = 1
    olSave                  = 0
    olPromptForSave         = 2
    
    # Excel row navigation
    xlUp                    = -4162
    xlDown                  = -4121
    
    # Folder types
    olFolderInbox           = 6
    olFolderSentMail        = 5
    olFolderDeletedItems    = 3
}

# ============================================================================
# CONFIGURATION
# ============================================================================
$Config = @{
    SafelistDomains = @(
        "microsoft.com",
        "office365.com",
        "xtrac.com",
        "apple.com",
        "gov.uk",
        "hmrc.gov.uk"
    )
    FilesPath = "C:\Scripts\Phishing\files"
    TemplatePath = "C:\Scripts\Phishing\phish_template.oft"
    SignaturePath = "C:\Users\kevin_curtis\AppData\Roaming\Microsoft\Signature\Reply (Kevin_Curtis@xtrac.com).htm"
    ChangeLogUrl = "https://xtractransmissions.sharepoint.com/sites/ITIncidents/Shared Documents/general/live Registers/RE-0004 IT Daily Control Log.xlsx"
    FileRetentionDays = 4
    WindowWaitTimeoutSeconds = 10
    MailboxName = "Xtrac Phish Watch"

    # LLM Configuration
    LLM = @{
        ModelPath   = "C:\models\qwen2.5-1.5b-instruct-q8_0.gguf"
        MaxTokens   = 300
        Temperature = 0.3
        GpuLayers   = 99
    }
}

# ============================================================================
# EMAIL TYPE DEFINITIONS
# ============================================================================
$EmailTypes = @{
    Real = @{
        "1" = @{ Name = "Cold call/Sales"; Description = "Unsolicited sales or marketing email" }
        "2" = @{ Name = "Newsletter/Subscription"; Description = "Newsletter or mailing list" }
        "3" = @{ Name = "Internal email"; Description = "Legitimate internal company email" }
        "4" = @{ Name = "External partner"; Description = "Legitimate email from known partner/vendor" }
    }
    Phish = @{
        "1" = @{ Name = "Credential harvesting"; Description = "Attempts to steal login credentials" }
        "2" = @{ Name = "Malware/Attachment"; Description = "Contains malicious attachments or links" }
        "3" = @{ Name = "BEC/Impersonation"; Description = "Business email compromise or impersonation" }
        "4" = @{ Name = "Generic scam"; Description = "General scam or fraud attempt" }
    }
}

# ============================================================================
# SCRIPT-LEVEL VARIABLES
# ============================================================================
$script:Outlook = $null
$script:Namespace = $null
$script:ExcelApp = $null

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $colors = @{
        'Info' = 'Cyan'
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Error' = 'Red'
    }
    
    $prefixes = @{
        'Info' = '►'
        'Success' = '✓'
        'Warning' = '⚠'
        'Error' = '✗'
    }
    
    Write-Host "$($prefixes[$Level]) $Message" -ForegroundColor $colors[$Level]
}

# ============================================================================
# Improvement #1 - Dynamic user identity instead of hardcoded name
# ============================================================================
function Get-CurrentUserIdentity {
    <#
    .SYNOPSIS
    Gets the current user's display name and email for changelog entries
    .OUTPUTS
    PSCustomObject with Name and Email properties
    #>
    try {
        # Try to get from Active Directory first
        $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
        $result = $searcher.FindOne()
        
        if ($result) {
            $displayName = $result.Properties["displayname"][0]
            $email = $result.Properties["mail"][0]
            
            return [PSCustomObject]@{
                Name  = $displayName
                Email = $email
            }
        }
    } catch {
        # AD lookup failed, fall back to alternatives
    }
    
    # Fallback: Try to get from Outlook profile
    try {
        if ($script:Namespace) {
            $currentUser = $script:Namespace.CurrentUser
            if ($currentUser) {
                $exchangeUser = $currentUser.GetExchangeUser()
                if ($exchangeUser) {
                    return [PSCustomObject]@{
                        Name  = $exchangeUser.Name
                        Email = $exchangeUser.PrimarySmtpAddress
                    }
                }
            }
        }
    } catch {
        # Outlook lookup failed
    }
    
    # Final fallback: Use environment variables
    return [PSCustomObject]@{
        Name  = $env:USERNAME
        Email = "$env:USERNAME@$env:USERDNSDOMAIN"
    }
}

# ============================================================================
# Improvement #5 - Single unified function for resolving sender email
# ============================================================================
function Resolve-SenderEmail {
    <#
    .SYNOPSIS
    Resolves the sender email address from a mail item using multiple methods
    .PARAMETER MailItem
    The Outlook mail item to extract sender from
    .OUTPUTS
    String containing the SMTP email address, or $null if unable to resolve
    #>
    param(
        [Parameter(Mandatory = $true)]
        $MailItem
    )
    
    if ($MailItem -eq $null) {
        return $null
    }
    
    $senderEmail = $null
    
    # Method 1: Try Exchange User (most reliable for internal/Exchange addresses)
    try {
        if ($MailItem.Sender -ne $null) {
            $exchangeUser = $MailItem.Sender.GetExchangeUser()
            if ($exchangeUser -ne $null -and -not [string]::IsNullOrWhiteSpace($exchangeUser.PrimarySmtpAddress)) {
                $senderEmail = $exchangeUser.PrimarySmtpAddress
                Write-Log "Resolved sender via Exchange User: $senderEmail" -Level Info
                return $senderEmail
            }
        }
    } catch {
        # Exchange user method failed, continue to next method
    }
    
    # Method 2: Try AddressEntry for Exchange format addresses
    try {
        if ($MailItem.Sender -ne $null) {
            $addressEntry = $MailItem.Sender
            if ($addressEntry.AddressEntryUserType -eq 0) {
                # olExchangeUserAddressEntry
                $exchangeUser = $addressEntry.GetExchangeUser()
                if ($exchangeUser -ne $null) {
                    $senderEmail = $exchangeUser.PrimarySmtpAddress
                    Write-Log "Resolved sender via AddressEntry: $senderEmail" -Level Info
                    return $senderEmail
                }
            }
        }
    } catch {
        # AddressEntry method failed
    }
    
    # Method 3: Direct SenderEmailAddress property
    try {
        $directAddress = $MailItem.SenderEmailAddress
        if (-not [string]::IsNullOrWhiteSpace($directAddress)) {
            # Check if it's an Exchange format address (/O=...)
            if ($directAddress -like "/O=*" -or $directAddress -like "/o=*") {
                # Try to resolve Exchange format to SMTP
                try {
                    $recipient = $script:Namespace.CreateRecipient($directAddress)
                    $recipient.Resolve()
                    if ($recipient.Resolved) {
                        $exchangeUser = $recipient.AddressEntry.GetExchangeUser()
                        if ($exchangeUser -ne $null) {
                            $senderEmail = $exchangeUser.PrimarySmtpAddress
                            Write-Log "Resolved Exchange format sender: $senderEmail" -Level Info
                            return $senderEmail
                        }
                    }
                } catch {
                    Write-Log "Could not resolve Exchange format address" -Level Warning
                }
            } else {
                # Already an SMTP address
                $senderEmail = $directAddress
                Write-Log "Resolved sender via SenderEmailAddress: $senderEmail" -Level Info
                return $senderEmail
            }
        }
    } catch {
        # Direct address method failed
    }
    
    # Method 4: PropertyAccessor for SMTP address
    try {
        $PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
        $senderEmail = $MailItem.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS)
        if (-not [string]::IsNullOrWhiteSpace($senderEmail)) {
            Write-Log "Resolved sender via PropertyAccessor: $senderEmail" -Level Info
            return $senderEmail
        }
    } catch {
        # PropertyAccessor method failed
    }
    
    Write-Log "Unable to resolve sender email address" -Level Warning
    return $null
}

function Get-SignatureHtml {
    <#
    .SYNOPSIS
    Loads the HTML signature file and extracts the body content with styles
    #>
    param([string]$SignaturePath)

    if (-not (Test-Path -LiteralPath $SignaturePath)) {
        Write-Log "Signature file not found: $SignaturePath" -Level Warning
        return ""
    }

    try {
        $signatureContent = Get-Content -LiteralPath $SignaturePath -Raw -Encoding Default

        # Extract content between <body> and </body> tags
        if ($signatureContent -match '(?s)<body[^>]*>(.*)</body>') {
            $bodyContent = $matches[1]

            # Add inline margin:0 to <p> tags without existing style attribute
            $bodyContent = $bodyContent -replace '<p class=MsoNormal>', '<p style="margin:0">'
            $bodyContent = $bodyContent -replace '<p class=MsoAutoSig>', '<p style="margin:0">'

            # For <p> tags that already have a style attribute, prepend margin:0 to existing styles
            $bodyContent = $bodyContent -replace "<p class=MsoNormal style='", "<p style='margin:0;"

            Write-Log "Signature loaded successfully" -Level Success
            return $bodyContent
        } else {
            Write-Log "Could not extract body from signature file" -Level Warning
            return ""
        }
    } catch {
        Write-Log "Error loading signature: $_" -Level Error
        return ""
    }
}

function Test-EmailSafeListed {
    param([string]$EmailAddress)

    if ([string]::IsNullOrWhiteSpace($EmailAddress)) { return $false }

    $domain = $EmailAddress -replace '^.*@', ''

    foreach ($protectedDomain in $Config.SafelistDomains) {
        if ($domain -like "*$protectedDomain") {
            Write-Log "PROTECTED: $EmailAddress matches safelist domain: $protectedDomain" -Level Warning
            return $true
        }
    }
    return $false
}

# ============================================================================
# LLM FUNCTIONS
# ============================================================================
function Get-LLMPromptTemplates {
    <#
    .SYNOPSIS
    Returns prompt templates for different email classification scenarios
    #>
    return @{
        SystemPrompt = @"
You are a member of the IT Security team responding to employees who have reported potential phishing emails. Be warm, professional and concise. Always thank them for their vigilance. Do not include a subject line. Do not include a sign-off or signature as one will be added automatically. Do not use bullet points or numbered lists.

IMPORTANT: Structure your response as 3-4 separate paragraphs. Put a blank line between each paragraph. Each paragraph should be 2-3 sentences maximum.
"@

        Real = @{
            "Cold call/Sales" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: NOT PHISHING - this was a legitimate cold call/sales email.
Explain that while unsolicited sales emails can seem suspicious, this one appears to be from a legitimate marketing source. Suggest they can safely unsubscribe if they don't want future emails. Thank them for their caution and reassure them that reporting suspicious emails is always the right thing to do. Tell them that the original email has been reattached if they need it.
"@
            "Newsletter/Subscription" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: NOT PHISHING - this was a legitimate newsletter or subscription email.
Explain that this appears to be from a mailing list they may have subscribed to previously. Suggest they can unsubscribe using the link at the bottom of the email if they no longer wish to receive it. Thank them for their vigilance and reassure them that it's always better to check. Tell them that the original email has been reattached if they need it.
"@
            "Internal email" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: NOT PHISHING - this was a legitimate internal company email.
Explain that this email came from a verified internal source and is genuine. Thank them for their caution as it shows good security awareness. Reassure them that checking internal emails that seem unusual is always the right approach. Tell them that the original email has been reattached if they need it.
"@
            "External partner" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: NOT PHISHING - this was a legitimate email from a known partner or vendor.
Explain that we have verified this email came from an established business partner/vendor. Thank them for their vigilance as external emails should always be treated with care. Reassure them that reporting suspicious emails helps keep everyone safe. Tell them that the original email has been reattached if they need it.
"@
        }

        Phish = @{
            "Credential harvesting" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: CONFIRMED PHISHING - this was a credential harvesting attempt.
Thank them for catching this phishing email that was designed to steal login credentials. Explain that we have taken action to block the sender. Remind them never to enter credentials via email links and to always navigate directly to websites. Praise their security awareness.
"@
            "Malware/Attachment" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: CONFIRMED PHISHING - this email contained malicious content.
Thank them for catching this dangerous phishing email that contained malware or malicious links. Explain that we have blocked the sender and are monitoring for any impact. Remind them to never open unexpected attachments or click suspicious links. Their quick action may have prevented a security incident.
"@
            "BEC/Impersonation" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: CONFIRMED PHISHING - this was a business email compromise attempt.
Thank them for identifying this impersonation attack. Explain that the attacker was pretending to be someone they're not, possibly requesting money or sensitive information. We have blocked the sender. Remind them to always verify unusual requests through a separate communication channel. Their vigilance is exactly what protects the company.
"@
            "Generic scam" = @"
Write a response to {FirstName} who reported a suspected phishing email.
Outcome: CONFIRMED PHISHING - this was a scam email.
Thank them for reporting this phishing/scam email. Explain that we have blocked the sender to prevent others from receiving similar messages. Remind them that legitimate organisations will never pressure them or ask for sensitive information via email. Their report helps protect the entire organisation.
"@
        }
    }
}

function Invoke-LLMResponse {
    <#
    .SYNOPSIS
    Generates an email response using the local LLM
    .PARAMETER FirstName
    The first name of the person to respond to
    .PARAMETER IsPhishing
    Whether the email was classified as phishing
    .PARAMETER EmailType
    The specific type/category of email
    .OUTPUTS
    String containing the generated response
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FirstName,

        [Parameter(Mandatory = $true)]
        [bool]$IsPhishing,

        [Parameter(Mandatory = $true)]
        [string]$EmailType
    )

    $templates = Get-LLMPromptTemplates
    $category = if ($IsPhishing) { "Phish" } else { "Real" }

    if (-not $templates[$category].ContainsKey($EmailType)) {
        Write-Log "Unknown email type: $EmailType for category: $category" -Level Error
        return $null
    }

    $userPrompt = $templates[$category][$EmailType] -replace '\{FirstName\}', $FirstName
    $systemPrompt = $templates.SystemPrompt

    $fullPrompt = @"
<|im_start|>system
$systemPrompt<|im_end|>
<|im_start|>user
$userPrompt<|im_end|>
<|im_start|>assistant
Hi $FirstName,
"@

    Write-Log "Generating LLM response..." -Level Info

    try {
        $llmArgs = @(
            "-m", $Config.LLM.ModelPath,
            "-p", $fullPrompt,
            "-n", $Config.LLM.MaxTokens,
            "--temp", $Config.LLM.Temperature,
            "-ngl", $Config.LLM.GpuLayers,
            "--single-turn"
        )
        $rawOutput = & llama-cli @llmArgs 2>$null

        $cleanOutput = $rawOutput | Out-String

        # Clean up the output
        if ($cleanOutput -match '(?s)\(truncated\)\s*(.+)$') {
            $cleanOutput = $matches[1].Trim()
        } elseif ($cleanOutput -match '(?s)<\|im_start\|>assistant\s*Hi [^,]+,\s*(.+)$') {
            $cleanOutput = "Hi $FirstName,`n" + $matches[1].Trim()
        }

        # Remove trailing junk
        $cleanOutput = $cleanOutput -replace '\[\s*Prompt:.*', ''
        $cleanOutput = $cleanOutput -replace 'Exiting\.\.\.', ''
        $cleanOutput = $cleanOutput -replace 'PS C:\\.*$', ''
        $cleanOutput = $cleanOutput.Trim()

        # Add greeting back if missing
        if ($cleanOutput -notmatch '^Hi ') {
            $cleanOutput = "Hi $FirstName,`n`n$cleanOutput"
        }

        Write-Log "LLM response generated successfully" -Level Success
        return $cleanOutput

    } catch {
        Write-Log "Error generating LLM response: $_" -Level Error
        return $null
    }
}

function Show-EmailTypeMenu {
    <#
    .SYNOPSIS
    Displays the email classification menu and returns user selection
    .OUTPUTS
    Hashtable with IsPhishing (bool), EmailType (string), Action (string)
    #>

    Write-Host ""
    Write-Host "=" * 50 -ForegroundColor Cyan
    Write-Host "  EMAIL CLASSIFICATION" -ForegroundColor Cyan
    Write-Host "=" * 50 -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[R] Real/Legitimate Email" -ForegroundColor Green
    Write-Host "[P] Phishing Email" -ForegroundColor Red
    Write-Host "[S] Skip this email" -ForegroundColor Yellow
    Write-Host "[Q] Quit processing" -ForegroundColor Yellow
    Write-Host ""

    $mainChoice = Read-Host "Select classification (R/P/S/Q)"
    $mainChoice = $mainChoice.ToUpper()

    switch ($mainChoice) {
        'S' { return @{ Action = 'Skip' } }
        'Q' { return @{ Action = 'Quit' } }
        'R' {
            Write-Host ""
            Write-Host "-" * 40 -ForegroundColor Green
            Write-Host "  LEGITIMATE EMAIL TYPE" -ForegroundColor Green
            Write-Host "-" * 40 -ForegroundColor Green
            foreach ($key in ($EmailTypes.Real.Keys | Sort-Object)) {
                $type = $EmailTypes.Real[$key]
                Write-Host "[$key] $($type.Name)" -ForegroundColor White
                Write-Host "    $($type.Description)" -ForegroundColor DarkGray
            }
            Write-Host ""

            $subChoice = Read-Host "Select type (1-4)"
            if ($EmailTypes.Real.ContainsKey($subChoice)) {
                return @{
                    Action = 'Classify'
                    IsPhishing = $false
                    EmailType = $EmailTypes.Real[$subChoice].Name
                    FolderName = 'Rejected'
                }
            } else {
                Write-Log "Invalid selection" -Level Warning
                return @{ Action = 'Skip' }
            }
        }
        'P' {
            Write-Host ""
            Write-Host "-" * 40 -ForegroundColor Red
            Write-Host "  PHISHING EMAIL TYPE" -ForegroundColor Red
            Write-Host "-" * 40 -ForegroundColor Red
            foreach ($key in ($EmailTypes.Phish.Keys | Sort-Object)) {
                $type = $EmailTypes.Phish[$key]
                Write-Host "[$key] $($type.Name)" -ForegroundColor White
                Write-Host "    $($type.Description)" -ForegroundColor DarkGray
            }
            Write-Host ""

            $subChoice = Read-Host "Select type (1-4)"
            if ($EmailTypes.Phish.ContainsKey($subChoice)) {
                return @{
                    Action = 'Classify'
                    IsPhishing = $true
                    EmailType = $EmailTypes.Phish[$subChoice].Name
                    FolderName = 'Genuine'
                }
            } else {
                Write-Log "Invalid selection" -Level Warning
                return @{ Action = 'Skip' }
            }
        }
        default {
            Write-Log "Invalid selection" -Level Warning
            return @{ Action = 'Skip' }
        }
    }
}

function Initialize-Outlook {
    try {
        $script:Outlook = New-Object -ComObject Outlook.Application
        $script:Namespace = $script:Outlook.GetNameSpace("MAPI")
        Write-Log "Connected to Outlook" -Level Success
        return $true
    } catch {
        Write-Log "Error connecting to Outlook: $_" -Level Error
        return $false
    }
}

function Get-PhishingMailbox {
    try {
        $mailbox = $script:Namespace.Folders | Where-Object { $_.Name -eq $Config.MailboxName }
        if (-not $mailbox) { throw "Mailbox not found" }
        
        $inbox = $mailbox.Folders | Where-Object { $_.Name -eq "Inbox" }
        if (-not $inbox) { throw "Inbox not found" }
        
        Write-Log "Accessed '$($Config.MailboxName)' inbox" -Level Success
        return $inbox
    } catch {
        Write-Log "Error accessing mailbox: $_" -Level Error
        return $null
    }
}

function Wait-ForWindow {
    param(
        [string]$ProcessName = "OUTLOOK",
        [int]$TimeoutSeconds = 10
    )
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        Start-Sleep -Milliseconds 500
        
        try {
            $inspector = $script:Outlook.ActiveInspector()
            if ($inspector -ne $null) {
                Start-Sleep -Milliseconds 500
                return $true
            }
        } catch {
            # Continue waiting
        }
    }
    
    Write-Log "Timeout waiting for window to open" -Level Warning
    return $false
}

function Get-AttachedEmailSender {
    <#
    .SYNOPSIS
    Gets the sender email from the currently open email attachment
    #>
    try {
        $inspector = $script:Outlook.ActiveInspector()
        if ($inspector -eq $null) {
            Write-Log "No active inspector window" -Level Warning
            return $null
        }
        
        $mail = $inspector.CurrentItem
        if ($mail -eq $null -or $mail.Class -ne $OlConstants.olMailItem) {
            Write-Log "No valid email in inspector" -Level Warning
            return $null
        }
        
        # Use unified sender resolution function (Improvement #5)
        return Resolve-SenderEmail -MailItem $mail
        
    } catch {
        Write-Log "Error getting sender email: $_" -Level Error
        return $null
    }
}

function Capture-WindowScreenshot {
    param([string]$OutputPath)
    
    try {
        $inspector = $script:Outlook.ActiveInspector()
        if ($inspector -eq $null) {
            Write-Log "No active inspector for screenshot" -Level Warning
            return $null
        }
        
        $inspector.Activate()
        Start-Sleep -Milliseconds 300
        
        $hWnd = [User32]::GetForegroundWindow()
        $rect = New-Object User32+RECT
        $null = [User32]::GetWindowRect($hWnd, [ref]$rect)
        
        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top
        
        if ($width -le 0 -or $height -le 0) {
            Write-Log "Invalid window dimensions: ${width}x${height}" -Level Error
            return $null
        }
        
        $image = New-Object System.Drawing.Bitmap($width, $height)
        $graphic = [System.Drawing.Graphics]::FromImage($image)
        $point = New-Object System.Drawing.Point($rect.Left, $rect.Top)
        $graphic.CopyFromScreen($point, [System.Drawing.Point]::Empty, $image.Size)
        
        $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
        $fileName = "phish_$timestamp.png"
        $fullPath = "$OutputPath\$fileName"
        
        $image.Save($fullPath, [System.Drawing.Imaging.ImageFormat]::Png)
        
        $graphic.Dispose()
        $image.Dispose()
        
        Write-Log "Screenshot saved: $fullPath" -Level Success
        return $fullPath
    } catch {
        Write-Log "Error capturing screenshot: $_" -Level Error
        return $null
    }
}

function Close-ActiveInspector {
    try {
        $inspector = $script:Outlook.ActiveInspector()
        if ($inspector -ne $null) {
            $inspector.Close($OlConstants.olDiscard)
        }
    } catch {
        # Ignore errors when closing
    }
}

# Function to connect to Exchange Online if not already connected
function Connect-ToExchangeOnline {
    try {
        # Test if already connected by trying to run a simple command
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "✓ Already connected to Exchange Online" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        try {
            Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
            Write-Host "✓ Successfully connected to Exchange Online" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "✗ Error connecting to Exchange Online: $_" -ForegroundColor Red
            return $false
        }
    }
}

function Add-EmailsToTenantBlockList {
    param([string[]]$EmailAddresses)
    
    if ($EmailAddresses.Count -eq 0) {
        Write-Log "No email addresses to block" -Level Warning
        return
    }
    
    if (-not (Connect-ToExchangeOnline)) {
        Write-Log "Cannot block emails - Exchange Online connection failed" -Level Error
        return
    }
    
    Write-Log "Adding $($EmailAddresses.Count) addresses to tenant block list..." -Level Info
    
    $results = @{ Success = 0; Failed = 0; Skipped = 0 }
    
    foreach ($email in $EmailAddresses) {
        try {
            $existing = Get-TenantAllowBlockListItems -ListType Sender -Block -Entry $email -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Log "Already blocked: $email" -Level Warning
                $results.Skipped++
                continue
            }
            
            New-TenantAllowBlockListItems -ListType Sender -Block -Entries $email -NoExpiration -ErrorAction Stop
            Write-Log "Blocked: $email" -Level Success
            $results.Success++
        } catch {
            Write-Log "Failed to block $email : $($_.Exception.Message)" -Level Error
            $results.Failed++
        }
        
        Start-Sleep -Milliseconds 100
    }
    
    Write-Log "Block operation: $($results.Success) successful, $($results.Failed) failed, $($results.Skipped) skipped" -Level Info
}

# ============================================================================
# Improvement #1 - Updated to use dynamic user identity
# ============================================================================
function Update-ChangeLog {
    param([string[]]$BlockedEmails)
    
    if ($BlockedEmails.Count -eq 0) { return }
    
    Write-Log "Updating Excel change log..." -Level Info
    
    # Get current user identity dynamically
    $currentUser = Get-CurrentUserIdentity
    Write-Log "Logging changes as: $($currentUser.Name)" -Level Info
    
    try {
        $script:ExcelApp = New-Object -ComObject Excel.Application
        $script:ExcelApp.Visible = $true
        $script:ExcelApp.DisplayAlerts = $false
        
        $workbook = $script:ExcelApp.Workbooks.Open($Config.ChangeLogUrl)
        $sheet = $workbook.Sheets.Item("Phishing Block Log")
        $lastRow = $sheet.Cells($sheet.Rows.Count, 1).End($OlConstants.xlUp).Row + 1
        
        foreach ($email in $BlockedEmails) {
            $sheet.Cells.Item($lastRow, 1) = Get-Date -Format "dd/MMM"
            $sheet.Cells.Item($lastRow, 2) = $currentUser.Name  # Dynamic user name
            $sheet.Cells.Item($lastRow, 3) = "Normal"
            $sheet.Cells.Item($lastRow, 4) = "GS"
            $sheet.Cells.Item($lastRow, 5) = "Implemented"
            $sheet.Cells.Item($lastRow, 6) = "Blocked email: $email"
            $sheet.Cells.Item($lastRow, 7) = "Added $email to the Tenant Block List (Sender)"
            $sheet.Cells.Item($lastRow, 8) = "Remove From Tenant Block List"
            $sheet.Cells.Item($lastRow, 9) = "Email"
            $sheet.Cells.Item($lastRow, 10) = ""
            $sheet.Cells.Item($lastRow, 11) = "Exchange Online"
            $sheet.Cells.Item($lastRow, 12) = "1"
            $sheet.Cells.Item($lastRow, 13) = "4"
            $sheet.Cells.Item($lastRow, 14) = "4"
            $sheet.Cells.Item($lastRow, 15) = "Low"
            $sheet.Cells.Item($lastRow, 16) = "Could Block Incorrect Address"
            $lastRow++
        }
        
        $workbook.Save()
        Write-Log "Change log updated" -Level Success
    } catch {
        Write-Log "Error updating change log: $_" -Level Error
    }
}

# ============================================================================
# Improvement #8 - COM object cleanup function
# ============================================================================
function Release-ComObject {
    <#
    .SYNOPSIS
    Properly releases a COM object to prevent memory leaks
    .PARAMETER ComObject
    The COM object to release
    .PARAMETER Name
    Optional name for logging purposes
    #>
    param(
        [Parameter(Mandatory = $true)]
        $ComObject,
        [string]$Name = "COM Object"
    )
    
    if ($ComObject -eq $null) { return }
    
    try {
        $refCount = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
        while ($refCount -gt 0) {
            $refCount = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
        }
    } catch {
        # Ignore release errors
    }
}

function Cleanup-AllComObjects {
    <#
    .SYNOPSIS
    Releases all COM objects used by the script
    #>
    Write-Log "Cleaning up COM objects..." -Level Info
    
    # Release Excel
    if ($script:ExcelApp -ne $null) {
        try {
            $script:ExcelApp.Quit()
        } catch { }
        Release-ComObject -ComObject $script:ExcelApp -Name "Excel Application"
        $script:ExcelApp = $null
    }
    
    # Release Outlook objects
    if ($script:Namespace -ne $null) {
        Release-ComObject -ComObject $script:Namespace -Name "Outlook Namespace"
        $script:Namespace = $null
    }
    
    if ($script:Outlook -ne $null) {
        Release-ComObject -ComObject $script:Outlook -Name "Outlook Application"
        $script:Outlook = $null
    }
    
    # Force garbage collection to clean up any remaining references
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    Write-Log "COM objects released" -Level Success
}

# ============================================================================
# TYPE DEFINITIONS
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

# ============================================================================
# MAIN SCRIPT - Improvement #7 - Structured error handling with try/catch/finally
# ============================================================================

try {
    # Ensure output directory exists
    if (-not (Test-Path $Config.FilesPath)) {
        New-Item -ItemType Directory -Force -Path $Config.FilesPath | Out-Null
    }

    # Initialize Outlook
    if (-not (Initialize-Outlook)) { 
        throw "Failed to initialize Outlook"
    }

    $inbox = Get-PhishingMailbox
    if ($inbox -eq $null) { 
        throw "Failed to access phishing mailbox"
    }

    # Get all emails upfront
    Write-Log "Loading emails from inbox..." -Level Info
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)

    # Build a list of emails to process
    $emailsToProcess = @()
    $item = $items.GetFirst()
    while ($item -ne $null) {
        if ($item.Subject -match "^3\|") {
            $emailsToProcess += $item
        }
        $item = $items.GetNext()
    }

    Write-Log "Found $($emailsToProcess.Count) emails to process (matching '3|' pattern)" -Level Info

    if ($emailsToProcess.Count -eq 0) {
        Write-Log "No matching emails found" -Level Warning
        # Still need to run finally block, so we'll exit gracefully
    } else {
        # Tracking
        $emailsToBlock = [System.Collections.ArrayList]::new()
        $processedCount = 0
        $shouldQuit = $false

        # Process each email
        foreach ($currentEmail in $emailsToProcess) {
            if ($shouldQuit) { break }
            
            $processedCount++
            
            # Reset variables for this iteration
            $screenFile = $null
            $blockEmail = $null
            $attachmentPath = $null
            $screenshotPath = $null
            
            Write-Host "`n$("=" * 60)" -ForegroundColor DarkGray
            Write-Log "Processing $processedCount of $($emailsToProcess.Count): '$($currentEmail.Subject)'" -Level Info
            
            try {
                # Get report sender info using unified function (Improvement #5)
                $reporterEmail = Resolve-SenderEmail -MailItem $currentEmail
                $reporterName = $currentEmail.SenderName
                Write-Log "Reporter: $reporterName <$reporterEmail>" -Level Info
                
                # Find and save attachment
                foreach ($att in $currentEmail.Attachments) {
                    if ($att.FileName -match "\.(msg|eml)$") {
                        $sanitizedFileName = ($att.FileName -replace "[^\w\.]", "_")
                        $attachmentPath = Join-Path $Config.FilesPath $sanitizedFileName
                        $att.SaveAsFile($attachmentPath)
                        Write-Log "Saved attachment: $sanitizedFileName" -Level Success
                        break
                    }
                }
                
                if (-not $attachmentPath -or -not (Test-Path -LiteralPath $attachmentPath)) {
                    Write-Log "No valid email attachment found - skipping" -Level Warning
                    continue
                }
                
                # Open the attached email
                Start-Process -FilePath $attachmentPath
                
                # Wait for window to open
                if (-not (Wait-ForWindow -TimeoutSeconds $Config.WindowWaitTimeoutSeconds)) {
                    Write-Log "Could not open attachment window" -Level Error
                    continue
                }
                
                # Get the phishing email sender
                $blockEmail = Get-AttachedEmailSender
                if ($blockEmail) {
                    Write-Log "Phishing sender to block: $blockEmail" -Level Info
                } else {
                    Write-Log "Could not determine sender to block" -Level Warning
                }
                
                # Capture screenshot
                $screenFile = Capture-WindowScreenshot -OutputPath $Config.FilesPath

                if (-not $screenFile) {
                    Write-Log "Screenshot failed - continuing anyway" -Level Warning
                }

                # Load and prepare template
                if (-not (Test-Path -LiteralPath $Config.TemplatePath)) {
                    Write-Log "Template not found: $($Config.TemplatePath)" -Level Error
                    continue
                }
                
                $template = $script:Outlook.CreateItemFromTemplate($Config.TemplatePath)
                
                $firstName = ($reporterName -split ' ')[0]
                
                # Build the image tag
                $imageHtml = ""
                $screenshotPath = [string]$screenFile
                
                if ($screenshotPath -and ($screenshotPath.EndsWith(".png"))) {
                    if (Test-Path -LiteralPath $screenshotPath) {
                        try {
                            $bytes = Get-Content -LiteralPath $screenshotPath -Encoding Byte -Raw
                            $base64String = [Convert]::ToBase64String($bytes)
                            $imageHtml = "<p><img src='data:image/png;base64,$base64String' alt='Screenshot' /></p>"
                            Write-Log "Screenshot embedded successfully" -Level Success
                        } catch {
                            Write-Log "Failed to embed screenshot: $_" -Level Warning
                        }
                    }
                }
                
                # Show classification menu (phishing email stays open for review)
                $classification = Show-EmailTypeMenu

                # Now close the attachment window after classification decision is made
                Close-ActiveInspector

                switch ($classification.Action) {
                    'Quit' {
                        Write-Log "User requested quit" -Level Warning
                        $shouldQuit = $true
                    }
                    'Skip' {
                        Write-Log "Skipped" -Level Info
                    }
                    'Classify' {
                        # Generate LLM response
                        $llmResponse = Invoke-LLMResponse -FirstName $firstName -IsPhishing $classification.IsPhishing -EmailType $classification.EmailType

                        if ($llmResponse) {
                            # Convert plain text to HTML paragraphs
                            # Normalize line endings using actual character codes
                            $normalizedResponse = $llmResponse -replace "\r\n", "`n" -replace "\r", "`n"

                            # Check for newlines to determine paragraph splitting mode
                            $hasDoubleNewline = $normalizedResponse.Contains([Environment]::NewLine + [Environment]::NewLine) -or $normalizedResponse.Contains("`n`n")
                            $hasSingleNewline = $normalizedResponse.Contains("`n")

                            # Split into paragraphs
                            $paragraphs = @()
                            if ($hasDoubleNewline) {
                                $paragraphs = $normalizedResponse -split "(\r?\n){2,}"
                            } elseif ($hasSingleNewline) {
                                $paragraphs = $normalizedResponse -split "\r?\n"
                            } else {
                                # No newlines - split by sentences and group into paragraphs of 2-3 sentences
                                $sentences = [regex]::Split($normalizedResponse, '(?<=[.!?])\s+')
                                $currentPara = ""
                                $sentenceCount = 0
                                foreach ($sentence in $sentences) {
                                    if ([string]::IsNullOrWhiteSpace($sentence)) { continue }
                                    $currentPara += "$sentence "
                                    $sentenceCount++
                                    if ($sentenceCount -ge 2) {
                                        $paragraphs += $currentPara.Trim()
                                        $currentPara = ""
                                        $sentenceCount = 0
                                    }
                                }
                                if ($currentPara.Trim()) {
                                    $paragraphs += $currentPara.Trim()
                                }
                            }

                            # Convert paragraphs to HTML
                            $llmHtml = ($paragraphs | Where-Object { $_.Trim() } | ForEach-Object {
                                $paragraphText = $_.Trim() -replace "`n", "<br>"
                                "<p>$paragraphText</p>"
                            }) -join "`n"

                            # Load signature
                            $signatureHtml = Get-SignatureHtml -SignaturePath $Config.SignaturePath

                            # Build the email body with LLM response
                            $newBody = "<html><body>"
                            $newBody += $llmHtml
                            $newBody += $imageHtml
                            $newBody += $signatureHtml
                            $newBody += "</body></html>"
                        } else {
                            # Load signature
                            $signatureHtml = Get-SignatureHtml -SignaturePath $Config.SignaturePath

                            # Fallback if LLM fails
                            $newBody = "<html><body>"
                            $newBody += "<p>Hello $firstName,</p>"
                            $newBody += "<p>[LLM response failed - please write manually]</p>"
                            $newBody += $imageHtml
                            $newBody += $signatureHtml
                            $newBody += "</body></html>"
                        }

                        $template.HTMLBody = $newBody
                        $template.Subject = "Report A Phish Feedback"
                        $template.SentOnBehalfOfName = "phishwatch@xtrac.com"
                        $null = $template.Recipients.Add($reporterEmail)

                        # Attach original email if it was legitimate (so user gets it back)
                        # Phishing emails just get the embedded screenshot in the body, no attachment
                        if (-not $classification.IsPhishing -and $attachmentPath -and (Test-Path -LiteralPath $attachmentPath)) {
                            $null = $template.Attachments.Add($attachmentPath)
                            Write-Log "Attached original email to response" -Level Success
                        }

                        $template.Display()
                        Write-Log "Template prepared with LLM response" -Level Success

                        # Handle blocking for phishing emails
                        if ($classification.IsPhishing -and $blockEmail) {
                            if (Test-EmailSafeListed -EmailAddress $blockEmail) {
                                Write-Host ("=" * 60) -ForegroundColor Red
                                Write-Log "SAFELIST PROTECTION: $blockEmail will NOT be blocked" -Level Error
                                Write-Host ("=" * 60) -ForegroundColor Red
                            } else {
                                $blockDecision = Read-Host "Block sender '$blockEmail'? (Y/N)"
                                if ($blockDecision.ToUpper() -eq 'Y') {
                                    if ($emailsToBlock -notcontains $blockEmail) {
                                        [void]$emailsToBlock.Add($blockEmail)
                                        Write-Log "Added to block list: $blockEmail" -Level Success
                                    } else {
                                        Write-Log "Already in block list: $blockEmail" -Level Warning
                                    }
                                }
                            }
                        }

                        # Move email to appropriate folder
                        $targetFolder = $inbox.Folders | Where-Object { $_.Name -eq $classification.FolderName }

                        if ($targetFolder) {
                            $currentEmail.UnRead = $false
                            $currentEmail.Save()
                            $currentEmail.Move($targetFolder)
                            Write-Log "Email moved to '$($classification.FolderName)'" -Level Success
                        } else {
                            Write-Log "Target folder '$($classification.FolderName)' not found" -Level Error
                        }
                    }
                }
            } catch {
                Write-Log "Error processing email: $_" -Level Error
                # Continue to next email rather than stopping entirely
                continue
            }
        }

        # Cleanup old files
        Write-Log "Cleaning up files older than $($Config.FileRetentionDays) days..." -Level Info
        $cutoffDate = (Get-Date).AddDays(-$Config.FileRetentionDays)
        Get-ChildItem -Path $Config.FilesPath -File | 
            Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
            Remove-Item -Force -ErrorAction SilentlyContinue

        # Process block list
        if ($emailsToBlock.Count -gt 0) {
            $uniqueEmails = $emailsToBlock | Sort-Object -Unique
            Write-Log "Processing $($uniqueEmails.Count) unique addresses for blocking..." -Level Info
            
            Add-EmailsToTenantBlockList -EmailAddresses $uniqueEmails
            Update-ChangeLog -BlockedEmails $uniqueEmails
        }

        Write-Host "`n$("=" * 60)" -ForegroundColor Green
        Write-Log "Processing complete. Processed $processedCount emails." -Level Success
        Write-Host ("=" * 60) -ForegroundColor Green
    }
}
catch {
    Write-Log "FATAL ERROR: $_" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
}
finally {
    # Improvement #7 & #8 - Always cleanup, even on error
    Write-Log "Running cleanup..." -Level Info
    
    # Disconnect from Exchange Online
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Exchange Online" -Level Success
    } catch {
        # Ignore disconnect errors
    }
    
    # Release all COM objects (Improvement #8)
    Cleanup-AllComObjects
    
    Write-Log "Cleanup complete" -Level Success
}

Pause