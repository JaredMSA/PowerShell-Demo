# Advanced Modifications & Customization Guide

This guide covers advanced customizations and modifications to the PowerShell email script.

## Table of Contents
1. [Subject Line Customization](#subject-line-customization)
2. [Adding CC/BCC Support](#adding-ccbcc-support)
3. [Dynamic Content in Templates](#dynamic-content-in-templates)
4. [Logging & Auditing](#logging--auditing)
5. [Error Recovery](#error-recovery)
6. [Performance Optimization](#performance-optimization)
7. [Custom Metadata](#custom-metadata)
8. [Integration with Other Tools](#integration-with-other-tools)

---

## Subject Line Customization

### Current Default
```powershell
$subject = "$($record['Email Template']) - $($record['Language'])"
```

### Option 1: Template Name Only
```powershell
$subject = $record['Email Template']
```

### Option 2: Template + Recipient Name
```powershell
$subject = "$($record['Email Template']) for $($record['Name'])"
```

### Option 3: Date-Based Subject
```powershell
$subject = "[$(Get-Date -Format 'yyyy-MM-dd')] $($record['Email Template'])"
```

### Option 4: Custom Prefix with Template
```powershell
$subject = "ACTION REQUIRED: $($record['Email Template']) - $($record['Name'])"
```

### Option 5: Language-Specific Subject
```powershell
if ($record['Language'] -eq 'FR') {
    $subject = "[FR] $($record['Email Template'])"
} else {
    $subject = "[EN] $($record['Email Template'])"
}
```

### Option 6: From Excel Column
Add a "Subject" column to your Excel and use:
```powershell
$subject = $record['Subject']
```

---

## Adding CC/BCC Support

### Step 1: Modify the Function Signature
Find this line in `Create-OutlookDraft`:
```powershell
param(
    [string]$RecipientEmail,
    [string]$RecipientName,
    [string]$Subject,
    [string]$HtmlBody,
    [string[]]$AttachmentPaths,
    [switch]$DryRun
)
```

Replace with:
```powershell
param(
    [string]$RecipientEmail,
    [string]$RecipientName,
    [string]$Subject,
    [string]$HtmlBody,
    [string[]]$AttachmentPaths,
    [string]$CcEmail = "",
    [string]$BccEmail = "",
    [switch]$DryRun
)
```

### Step 2: Add CC/BCC to Email
Find this section in the email creation:
```powershell
$mail.To = $RecipientEmail
$mail.Subject = $Subject
$mail.HTMLBody = $HtmlBody
```

Add after it:
```powershell
if ($CcEmail) {
    $mail.CC = $CcEmail
    Write-Verbose "    CC: $CcEmail"
}

if ($BccEmail) {
    $mail.BCC = $BccEmail
    Write-Verbose "    BCC: $BccEmail"
}
```

### Step 3: Include in Dry Run Output
Find the DryRun output section and add:
```powershell
if ($DryRun) {
    Write-Host "`n=== DRY RUN: Would create draft email ===" -ForegroundColor Cyan
    Write-Host "To: $RecipientEmail ($RecipientName)"
    if ($CcEmail) { Write-Host "CC: $CcEmail" }
    if ($BccEmail) { Write-Host "BCC: $BccEmail" }
    Write-Host "Subject: $Subject"
    Write-Host "Attachments: $($AttachmentPaths.Count) file(s)"
    foreach ($path in $AttachmentPaths) {
        Write-Host "  - $(Split-Path $path -Leaf)"
    }
    return $true
}
```

### Step 4: Update the Function Call
In the main loop, update the call:
```powershell
$success = Create-OutlookDraft `
    -RecipientEmail $record['Email'] `
    -RecipientName $record['Name'] `
    -Subject $subject `
    -HtmlBody $htmlBody `
    -AttachmentPaths $attachments `
    -CcEmail "manager@company.com" `
    -BccEmail "archive@company.com" `
    -DryRun:$DryRun
```

Or make it dynamic from Excel (add CC/BCC columns):
```powershell
$success = Create-OutlookDraft `
    -RecipientEmail $record['Email'] `
    -RecipientName $record['Name'] `
    -Subject $subject `
    -HtmlBody $htmlBody `
    -AttachmentPaths $attachments `
    -CcEmail $record['CC Email'] `
    -BccEmail $record['BCC Email'] `
    -DryRun:$DryRun
```

---

## Dynamic Content in Templates

### Method 1: Simple Find & Replace

Modify `Get-HtmlTemplate`:
```powershell
function Get-HtmlTemplate {
    param(
        [string]$TemplateName,
        [string]$TemplateFolder,
        [hashtable]$ReplaceVariables
    )

    $templatePath = Join-Path $TemplateFolder "$TemplateName.html"
    
    if (-not (Test-Path $templatePath -PathType Leaf)) {
        Write-Warning "Template not found: $templatePath"
        return $null
    }
    
    try {
        $content = Get-Content -Path $templatePath -Raw -Encoding UTF8
        
        # Replace variables
        if ($ReplaceVariables) {
            foreach ($key in $ReplaceVariables.Keys) {
                $content = $content -replace "{{$key}}", $ReplaceVariables[$key]
            }
        }
        
        Write-Verbose "Loaded template: $TemplateName"
        return $content
    }
    catch {
        Write-Error "Failed to read template '$TemplateName': $_"
        return $null
    }
}
```

Then in your template, use placeholders:
```html
<p>Dear {{NAME}},</p>
<p>Your email is {{EMAIL}} and your language preference is {{LANGUAGE}}.</p>
```

And call it like:
```powershell
$replaceVars = @{
    'NAME' = $record['Name']
    'EMAIL' = $record['Email']
    'LANGUAGE' = $record['Language']
}

$htmlBody = Get-HtmlTemplate -TemplateName $record['Email Template'] `
    -TemplateFolder $TemplateFolder `
    -ReplaceVariables $replaceVars
```

### Method 2: Language-Specific Templates

Instead of storing language in template name, load based on language:
```powershell
function Get-LocalizedTemplate {
    param(
        [string]$TemplateName,
        [string]$Language,
        [string]$TemplateFolder
    )
    
    $templatePath = Join-Path $TemplateFolder "$TemplateName`_$Language.html"
    
    if (-not (Test-Path $templatePath)) {
        # Fall back to English
        $templatePath = Join-Path $TemplateFolder "$TemplateName`_EN.html"
    }
    
    return Get-Content -Path $templatePath -Raw -Encoding UTF8
}
```

Usage:
```powershell
$htmlBody = Get-LocalizedTemplate -TemplateName "Welcome" `
    -Language $record['Language'] `
    -TemplateFolder $TemplateFolder
```

### Method 3: Date & Time in Templates

```html
<p>Email sent on: {{SEND_DATE}}</p>
```

```powershell
$replaceVars = @{
    'SEND_DATE' = Get-Date -Format 'MMMM dd, yyyy'
    'SEND_TIME' = Get-Date -Format 'HH:mm:ss'
}
```

---

## Logging & Auditing

### Basic File Logging

Add this function:
```powershell
function Write-LogEntry {
    param(
        [string]$LogPath,
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8
}
```

Use in main script:
```powershell
$logPath = "C:\EmailAutomation\logs\email_run_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"

Write-LogEntry -LogPath $logPath -Message "Script started"
Write-LogEntry -LogPath $logPath -Message "Processing record 1: jared@test.com"
Write-LogEntry -LogPath $logPath -Message "Email created successfully" -Level "SUCCESS"
Write-LogEntry -LogPath $logPath -Message "Script completed"
```

### CSV Audit Log

Create a persistent audit trail:
```powershell
function Write-AuditLog {
    param(
        [string]$AuditPath,
        [hashtable]$Record,
        [string]$Status,
        [string]$Notes
    )
    
    $auditEntry = @{
        'Timestamp' = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        'RecipientEmail' = $Record['Email']
        'RecipientName' = $Record['Name']
        'Template' = $Record['Email Template']
        'Status' = $Status
        'Attachments' = $Record.Values -join '; '
        'Notes' = $Notes
    }
    
    # Convert to CSV and append
    $csv = $auditEntry | ConvertTo-Csv -NoTypeInformation
    if ((Test-Path $AuditPath) -eq $false) {
        Set-Content -Path $AuditPath -Value $csv[0]
    }
    Add-Content -Path $AuditPath -Value $csv[1]
}
```

Usage:
```powershell
Write-AuditLog -AuditPath "C:\EmailAutomation\logs\audit.csv" `
    -Record $record `
    -Status "SUCCESS" `
    -Notes "Email created with 2 attachments"
```

---

## Error Recovery

### Retry Logic

```powershell
function Create-OutlookDraftWithRetry {
    param(
        [string]$RecipientEmail,
        [string]$RecipientName,
        [string]$Subject,
        [string]$HtmlBody,
        [string[]]$AttachmentPaths,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 2,
        [switch]$DryRun
    )
    
    $attempt = 0
    
    while ($attempt -lt $MaxRetries) {
        try {
            $attempt++
            
            if ($attempt -gt 1) {
                Write-Verbose "Retry attempt $attempt of $MaxRetries"
                Start-Sleep -Seconds $RetryDelaySeconds
            }
            
            # Original email creation logic here
            # ... (copy the original function body)
            
            return $true
        }
        catch {
            if ($attempt -ge $MaxRetries) {
                Write-Error "Failed to create email after $MaxRetries attempts: $_"
                return $false
            }
        }
    }
}
```

### Skip Failed Records

Wrap processing in try-catch:
```powershell
foreach ($index in 0..($emailData.Count - 1)) {
    try {
        # ... existing processing code ...
    }
    catch {
        Write-Error "Failed to process record $($index + 1): $_"
        $failureCount++
        continue
    }
}
```

---

## Performance Optimization

### Process in Parallel Batches

For large datasets, process multiple emails concurrently:
```powershell
function New-OutlookInstance {
    # Create a fresh COM instance for parallel processing
    return New-Object -ComObject Outlook.Application
}

# Process in batches of 5
$batchSize = 5
for ($i = 0; $i -lt $emailData.Count; $i += $batchSize) {
    $batch = $emailData[$i..([Math]::Min($i + $batchSize - 1, $emailData.Count - 1))]
    
    # Process this batch
    foreach ($record in $batch) {
        # ... process record ...
    }
    
    # Small pause between batches
    Start-Sleep -Milliseconds 500
}
```

### Cache Templates

```powershell
$templateCache = @{}

function Get-HtmlTemplate-Cached {
    param(
        [string]$TemplateName,
        [string]$TemplateFolder
    )
    
    if ($templateCache.ContainsKey($TemplateName)) {
        Write-Verbose "Template loaded from cache: $TemplateName"
        return $templateCache[$TemplateName]
    }
    
    $content = Get-HtmlTemplate -TemplateName $TemplateName -TemplateFolder $TemplateFolder
    $templateCache[$TemplateName] = $content
    return $content
}
```

---

## Custom Metadata

### Add Priority Field

Add to Excel: "Priority" column with values: High, Normal, Low

```powershell
function Get-EmailPriority {
    param([string]$Priority)
    
    switch ($Priority) {
        'High' { return 2 }  # olImportanceHigh
        'Low' { return 0 }   # olImportanceLow
        default { return 1 } # olImportanceNormal
    }
}

# In email creation:
$mail.Importance = Get-EmailPriority -Priority $record['Priority']
```

### Add Custom Headers

```powershell
$mail.UserProperties.Add("CampaignID", [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olText).Value = $record['Campaign ID']
$mail.UserProperties.Add("ProcessedBy", [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olText).Value = $env:USERNAME
```

---

## Integration with Other Tools

### Export Results to Excel

```powershell
function Export-ResultsToExcel {
    param([string]$OutputPath, [array]$Results)
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Sheets.Item(1)
    
    # Headers
    $worksheet.Cells.Item(1, 1) = "Timestamp"
    $worksheet.Cells.Item(1, 2) = "Email"
    $worksheet.Cells.Item(1, 3) = "Status"
    $worksheet.Cells.Item(1, 4) = "Notes"
    
    # Data
    $row = 2
    foreach ($result in $Results) {
        $worksheet.Cells.Item($row, 1) = $result.Timestamp
        $worksheet.Cells.Item($row, 2) = $result.Email
        $worksheet.Cells.Item($row, 3) = $result.Status
        $worksheet.Cells.Item($row, 4) = $result.Notes
        $row++
    }
    
    $workbook.SaveAs($OutputPath)
    $workbook.Close()
    $excel.Quit()
}
```

### Send Completion Report

```powershell
function Send-CompletionReport {
    param(
        [int]$SuccessCount,
        [int]$FailureCount,
        [string]$ReportRecipient
    )
    
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    
    $mail.To = $ReportRecipient
    $mail.Subject = "Email Batch Processing Complete"
    $mail.Body = @"
The email batch processing has completed.

Summary:
- Successful: $SuccessCount
- Failed: $FailureCount
- Total: $($SuccessCount + $FailureCount)

Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
User: $env:USERNAME
Computer: $env:COMPUTERNAME
"@
    
    $mail.Send()
}
```

### Slack Notification

```powershell
function Send-SlackNotification {
    param(
        [string]$WebhookUrl,
        [int]$SuccessCount,
        [int]$FailureCount
    )
    
    $payload = @{
        text = "Email batch processing complete"
        attachments = @(
            @{
                color = if ($FailureCount -eq 0) { "good" } else { "warning" }
                fields = @(
                    @{
                        title = "Successful"
                        value = $SuccessCount
                        short = $true
                    },
                    @{
                        title = "Failed"
                        value = $FailureCount
                        short = $true
                    }
                )
            }
        )
    } | ConvertTo-Json
    
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $payload -ContentType 'application/json'
}
```

---

## Command Line Parameter Customization

Add custom parameters to your script:

```powershell
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $true)]
    [string]$TemplateFolder,

    [Parameter(Mandatory = $true)]
    [string]$AttachmentFolder,

    [switch]$DryRun,
    
    # New custom parameters
    [int]$BatchSize = 5,
    [int]$DelayBetweenEmails = 100,
    [switch]$SendEmail,
    [switch]$GenerateReport,
    [string]$LogFolder = ".\logs",
    [string]$ReportRecipient = ""
)
```

Usage:
```powershell
.\Send-EmailsFromExcel.ps1 `
    -ExcelPath "emails.xlsx" `
    -TemplateFolder "templates" `
    -AttachmentFolder "attachments" `
    -BatchSize 10 `
    -DelayBetweenEmails 500 `
    -GenerateReport `
    -LogFolder "C:\logs"
```

---

## Notes

- Always backup your original script before making modifications
- Test changes with `-DryRun` first
- Document any customizations you make
- Consider version control for tracking changes
- Performance impacts depend on system resources
- COM object cleanup is critical for stability
