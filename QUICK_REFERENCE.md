# Quick Reference Sheet

## One-Liner Commands

### Test your setup (no emails created)
```powershell
.\Send-EmailsFromExcel.ps1 -ExcelPath "C:\path\to\email_list.xlsx" -TemplateFolder "C:\path\to\templates" -AttachmentFolder "C:\path\to\attachments" -DryRun
```

### Create draft emails
```powershell
.\Send-EmailsFromExcel.ps1 -ExcelPath "C:\path\to\email_list.xlsx" -TemplateFolder "C:\path\to\templates" -AttachmentFolder "C:\path\to\attachments"
```

### Run with verbose logging
```powershell
.\Send-EmailsFromExcel.ps1 -ExcelPath "C:\path\to\email_list.xlsx" -TemplateFolder "C:\path\to\templates" -AttachmentFolder "C:\path\to\attachments" -Verbose
```

## Execution Policy

If you get "cannot be loaded because running scripts is disabled" error:

```powershell
# Temporarily allow scripts for this session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Then run your script
```

Or use:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\path\to\Send-EmailsFromExcel.ps1" ...
```

## Excel File Requirements

| Column Name      | Required | Data Type | Example                |
|------------------|----------|-----------|------------------------|
| Name             | Yes      | Text      | Jared                  |
| Email            | Yes      | Email     | jared@test.com         |
| Email Template   | Yes      | Text      | Welcome                |
| Language         | No       | Text      | EN, FR                 |
| PDF              | No       | Filename  | 123.pdf                |
| Attachment 1     | No       | Filename  | Welcome Aboard.pdf     |
| Attachment 2     | No       | Filename  | FAQ.pdf                |
| Attachment N     | No       | Filename  | (unlimited columns)    |

## Folder Structure Checklist

```
✓ Excel file path correct
✓ Template folder path correct
✓ Attachment folder path correct
✓ All template files end in .html
✓ Template file names match Excel (exact spelling)
✓ All attachment files exist
✓ No special characters in file names
✓ Outlook installed and tested
✓ Excel installed
```

## What Gets Created

✓ **Draft emails** in Outlook (not sent)
✓ **Located** in Outlook > Drafts folder
✓ **Status**: Ready for review before sending
✓ **Can be edited** before sending
✓ **Attachments**: All included and accessible

## Expected Output

### Successful Run
```
PowerShell Email Draft Creator
==============================

Reading Excel file...
Found 3 email record(s)

Processing record 1 of 3...
✓ Created draft email for: jared@test.com

Processing record 2 of 3...
✓ Created draft email for: jared@test.com

Processing record 3 of 3...
✓ Created draft email for: sam@test.com

==============================
Summary:
  Successful: 3
  Failed: 0
  Total: 3

✓ All emails processed successfully!
```

### Dry Run Output
```
=== DRY RUN: Would create draft email ===
To: jared@test.com (Jared)
Subject: Welcome - EN
Attachments: 2 file(s)
  - 123.pdf
  - Welcome Aboard.pdf
```

## Common Error Messages & Fixes

| Error | Cause | Solution |
|-------|-------|----------|
| "Excel.Application is not available" | Excel not installed or not properly registered | Install Excel or repair Office installation |
| "Outlook.Application is not available" | Outlook not installed or not running | Install Outlook and open it once |
| "Template not found" | HTML file doesn't exist or name mismatch | Verify file exists and name matches exactly |
| "Attachment not found" | File missing from attachment folder | Check file exists and path is correct |
| "Cannot bind argument to parameter 'ExcelPath'" | Path doesn't exist or has quotes issues | Use full path without extra quotes: `"C:\full\path\file.xlsx"` |
| "Script is disabled on this system" | Execution policy blocks scripts | Use `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force` |

## Testing Checklist

- [ ] Run with `-DryRun` first to verify file resolution
- [ ] Check that all templates are found
- [ ] Verify all attachment files are located
- [ ] Create a test with 1-2 rows
- [ ] Review created drafts in Outlook
- [ ] Check attachments are included
- [ ] Test with different template files
- [ ] Run verbose mode to see detailed output

## Pro Tips

1. **Always use full paths** (C:\full\path\) not relative paths
2. **Test with DryRun first** - it's free!
3. **Check Outlook Drafts folder** after running to verify
4. **Use verbose mode** if anything seems wrong
5. **Keep an audit trail** - log the runs
6. **Backup your Excel** before processing
7. **Review templates** before bulk sending

## Modify Subject Line

Find this line in the script:
```powershell
$subject = "$($record['Email Template']) - $($record['Language'])"
```

Change to your preferred format:
```powershell
# Option 1: Just template name
$subject = $record['Email Template']

# Option 2: Template + name
$subject = "$($record['Email Template']) for $($record['Name'])"

# Option 3: Custom text
$subject = "New Employee Onboarding"
```

## Run from Scheduled Task

Create a batch file to run the script:

**run_emails.bat**
```batch
@echo off
cd /d "C:\EmailAutomation"
powershell -ExecutionPolicy Bypass -File "Send-EmailsFromExcel.ps1" `
  -ExcelPath "C:\EmailAutomation\excel\email_list.xlsx" `
  -TemplateFolder "C:\EmailAutomation\templates" `
  -AttachmentFolder "C:\EmailAutomation\attachments"
pause
```

Then schedule with Windows Task Scheduler pointing to this batch file.

## Performance Expectations

- **1-10 emails**: < 1 minute
- **11-50 emails**: 1-3 minutes
- **51-100 emails**: 3-5 minutes
- **100+ emails**: 5-10+ minutes (depends on attachment sizes)

Factors affecting speed:
- Number of attachments per email
- Size of HTML templates
- Outlook performance
- System load

## Getting Help

If something goes wrong:
1. Note the exact error message
2. Run with `-Verbose` for more details
3. Check that paths are correct (full paths only)
4. Verify files exist with File Explorer
5. Try the DryRun mode first
6. Close all Office applications and retry
