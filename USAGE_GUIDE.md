# PowerShell Email Draft Creator - Usage Guide

## Overview
This script reads an Excel file and creates draft emails in Outlook, automatically:
- Loading HTML email templates
- Attaching variable numbers of files (PDFs, Word docs, etc.)
- Populating recipient information from your Excel data

## Prerequisites
- PowerShell 5.0 or higher
- Microsoft Outlook installed
- Excel COM library installed
- Proper network/COM permissions

## Setup

### 1. Create Your Folder Structure
```
C:\EmailAutomation\
├── excel\
│   └── email_list.xlsx
├── templates\
│   ├── Welcome.html
│   ├── Welcome_F.html
│   ├── ESPP Login.html
│   └── (other .html templates)
└── attachments\
    ├── 123.pdf
    ├── Welcome Aboard.pdf
    ├── Login.pdf
    ├── FAQ.pdf
    ├── Tips and Tricks.pdf
    └── (other attachments)
```

### 2. Prepare Your Excel File
Your Excel file should have these columns:
- **Name**: Recipient's name
- **Email**: Recipient's email address
- **Email Template**: Name of the .html template file (without .html extension)
- **Language**: Language code (EN, FR, etc.)
- **PDF**: Main PDF file name
- **Attachment 1**: Additional attachment file name
- **Attachment 2**: Optional second additional attachment
- **Attachment N**: Add more columns as needed

**Example:**
| Name  | Email            | Email Template | Language | PDF        | Attachment 1           | Attachment 2            |
|-------|------------------|----------------|----------|------------|------------------------|-------------------------|
| Jared | jared@test.com   | Welcome        | EN       | 123.pdf    | Welcome Aboard.pdf     |                         |
| Sam   | sam@test.com     | ESPP Login     | EN       | Login.pdf  | FAQ.pdf                | Tips and Tricks.pdf     |

### 3. Create HTML Templates
Save your email templates as `.html` files in the templates folder. You can include:
- CSS styling
- Variable placeholders (optional - customize script as needed)
- Images (use absolute URLs)

**Example template file: `Welcome.html`**
```html
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        .header { color: #0066cc; margin-bottom: 20px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Welcome to Our Company!</h1>
    </div>
    <p>We're excited to have you join our team.</p>
    <p>Please review the attached documents for important information.</p>
</body>
</html>
```

## Usage

### Basic Usage (Test Mode)
Test your setup without creating actual emails:
```powershell
.\Send-EmailsFromExcel.ps1 `
    -ExcelPath "C:\EmailAutomation\excel\email_list.xlsx" `
    -TemplateFolder "C:\EmailAutomation\templates" `
    -AttachmentFolder "C:\EmailAutomation\attachments" `
    -DryRun
```

### Create Draft Emails
When ready to create actual draft emails:
```powershell
.\Send-EmailsFromExcel.ps1 `
    -ExcelPath "C:\EmailAutomation\excel\email_list.xlsx" `
    -TemplateFolder "C:\EmailAutomation\templates" `
    -AttachmentFolder "C:\EmailAutomation\attachments"
```

### Run with Verbose Output
For detailed logging:
```powershell
.\Send-EmailsFromExcel.ps1 `
    -ExcelPath "C:\EmailAutomation\excel\email_list.xlsx" `
    -TemplateFolder "C:\EmailAutomation\templates" `
    -AttachmentFolder "C:\EmailAutomation\attachments" `
    -Verbose
```

## Features

### ✓ Variable Attachment Handling
- Automatically detects all attachment columns
- Skips empty cells gracefully
- Handles missing files with warnings
- Supports unlimited attachment columns

### ✓ Error Handling
- Validates all paths exist before processing
- Reports missing templates and attachments
- Continues processing other records if one fails
- Provides summary at the end

### ✓ Dry Run Mode
- Preview what will happen before creating emails
- Check template and attachment resolution
- No emails created in Outlook

### ✓ Draft Creation (Not Sent)
- All emails saved to Outlook Drafts folder
- Review before manually sending
- Edit as needed before sending

## Customization

### Change Subject Line Format
Edit this line in the `Create-OutlookDraft` function:
```powershell
# Current:
$subject = "$($record['Email Template']) - $($record['Language'])"

# Or use custom logic:
$subject = "New Notification for $($record['Name'])"
```

### Add Dynamic Content to Templates
Modify `Get-HtmlTemplate` function to support variable replacement:
```powershell
$content = Get-Content -Path $templatePath -Raw -Encoding UTF8
$content = $content -replace '{{NAME}}', $record['Name']
$content = $content -replace '{{EMAIL}}', $record['Email']
return $content
```

### Add Custom Logging
Redirect output to a log file:
```powershell
.\Send-EmailsFromExcel.ps1 ... | Tee-Object -FilePath "C:\Logs\email_run.log"
```

## Troubleshooting

### "Excel.Application is not available"
- Ensure Microsoft Excel is installed
- Run PowerShell as Administrator
- Restart Excel if it's open

### "Outlook.Application is not available"
- Ensure Microsoft Outlook is installed
- Start Outlook manually once
- Check Windows firewall/security settings

### "Template not found"
- Verify template file exists in TemplateFolder
- Ensure filename matches Excel (case-sensitive on some systems)
- Template should be named exactly as in "Email Template" column (without .html)

### Attachments not appearing
- Verify attachment files exist in AttachmentFolder
- Check file extensions match exactly
- Ensure file paths don't contain special characters
- Run with `-Verbose` to see which files are being found

### COM Object errors
- Close any open Excel/Outlook windows
- Run PowerShell as Administrator
- Try restarting your computer

## Security Considerations

- Store Excel files securely (contains email addresses)
- Limit access to attachment folders
- Audit who has access to email templates
- Consider encrypting sensitive template content

## Performance Notes

- Processing 100+ rows may take several minutes
- Each email creation involves COM object interaction
- HTML parsing takes longer for complex templates
- Large attachments may slow processing

## Tips & Best Practices

1. **Test First**: Always run with `-DryRun` before actual execution
2. **Backup**: Keep a backup of your Excel file
3. **Version Templates**: Name versions like `Welcome_v2.html`
4. **Check Drafts**: Review created drafts in Outlook before bulk sending
5. **Batch Processing**: Process in batches of 50-100 for better control
6. **Maintain Lists**: Keep an archive of processed email lists

## Support & Modifications

If you need to modify the script:
- The main logic loop starts at "# Process each row"
- Add custom columns by extending the hashtable
- Modify attachment detection by changing the filter in `Get-AttachmentPaths`
- Add more email metadata by extending the mail object properties

## Example: Adding CC/BCC

To add CC or BCC, modify the `Create-OutlookDraft` function:

```powershell
param(
    ...
    [string]$CcEmail,
    [string]$BccEmail,
    ...
)

# Add after subject line:
if ($CcEmail) { $mail.CC = $CcEmail }
if ($BccEmail) { $mail.BCC = $BccEmail }
```

Then update the call to include:
```powershell
-CcEmail "manager@company.com" `
-BccEmail "archive@company.com"
```
