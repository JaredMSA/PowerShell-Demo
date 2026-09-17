# Setup & Deployment Checklist

Complete these steps in order to successfully deploy the email automation script.

## Phase 1: Environment Setup

### Prerequisites Check
- [ ] PowerShell 5.0 or higher installed
  ```powershell
  $PSVersionTable.PSVersion
  ```
  Should show version 5.0 or higher

- [ ] Microsoft Outlook installed and accessible
  - [ ] Outlook can be opened and is functional
  - [ ] Email account(s) configured in Outlook
  - [ ] Outlook Drafts folder is accessible

- [ ] Microsoft Excel installed
  - [ ] Excel can be opened
  - [ ] COM components are registered

- [ ] Run PowerShell as Administrator
  - Right-click PowerShell → Run as Administrator

### Execution Policy
- [ ] Check current execution policy:
  ```powershell
  Get-ExecutionPolicy
  ```

- [ ] Set appropriate policy (if needed):
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
  ```
  
  OR for just this session:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
  ```

## Phase 2: Folder Structure Setup

### Create Directory Structure
```
[ ] C:\EmailAutomation\                    (main folder)
    [ ] excel\                              (Excel files)
    [ ] templates\                          (HTML templates)
    [ ] attachments\                        (documents to attach)
    [ ] logs\                               (optional - for logging)
```

Commands to create:
```powershell
New-Item -ItemType Directory -Path "C:\EmailAutomation\excel" -Force
New-Item -ItemType Directory -Path "C:\EmailAutomation\templates" -Force
New-Item -ItemType Directory -Path "C:\EmailAutomation\attachments" -Force
New-Item -ItemType Directory -Path "C:\EmailAutomation\logs" -Force
```

### Script Placement
- [ ] Copy `Send-EmailsFromExcel.ps1` to `C:\EmailAutomation\`
- [ ] Verify script file is not blocked:
  ```powershell
  Unblock-File -Path "C:\EmailAutomation\Send-EmailsFromExcel.ps1"
  ```

## Phase 3: Template Preparation

### Create HTML Email Templates
For each email template in your Excel file:

- [ ] Create `Welcome.html`
  - [ ] Save in `C:\EmailAutomation\templates\`
  - [ ] Test HTML renders properly in Outlook
  - [ ] Check styling, fonts, colors

- [ ] Create `Welcome_F.html` (French version, if needed)
  - [ ] Save in `C:\EmailAutomation\templates\`

- [ ] Create `ESPP Login.html` (and other templates)
  - [ ] Each needs corresponding .html file
  - [ ] Name must match Excel "Email Template" column exactly

### HTML Template Checklist
For each template file:
- [ ] HTML is valid and well-formed
- [ ] CSS styling displays correctly in Outlook
- [ ] No external images/scripts (use absolute URLs if needed)
- [ ] File encoding is UTF-8
- [ ] No special characters in filename

## Phase 4: Attachment Organization

### Gather All Files
- [ ] Create list of all attachment files needed
- [ ] Collect PDFs, Word docs, etc.
- [ ] Verify file names match Excel exactly

### Copy to Attachments Folder
- [ ] Copy all files to `C:\EmailAutomation\attachments\`
- [ ] Verify file structure (no subfolders):
  ```
  C:\EmailAutomation\attachments\
  ├── 123.pdf
  ├── Welcome Aboard.pdf
  ├── Login.pdf
  ├── FAQ.pdf
  └── Tips and Tricks.pdf
  ```

### Attachment Checklist
- [ ] All file names match Excel exactly (case-sensitive on some systems)
- [ ] No special characters in file names
- [ ] All files are accessible and not corrupted
- [ ] File names don't have spaces at start/end
- [ ] Files aren't open in other applications

## Phase 5: Excel File Preparation

### Verify Excel Structure
- [ ] Open your Excel file
- [ ] Verify headers in Row 1:
  - [ ] Name
  - [ ] Email
  - [ ] Email Template
  - [ ] Language
  - [ ] PDF (or attachment filename column)
  - [ ] Attachment 1 (optional additional columns)
  - [ ] Attachment 2, etc.

### Verify Data
- [ ] All email addresses are valid
- [ ] Template names match your .html files exactly
  - [ ] "Welcome" → Welcome.html
  - [ ] "ESPP Login" → ESPP Login.html
- [ ] All attachment filenames match exactly
- [ ] No extra spaces in filenames
- [ ] Empty cells handled (nulls are fine)

### Save Excel File
- [ ] Save as .xlsx format (not .xls or .csv)
- [ ] Save to `C:\EmailAutomation\excel\`
- [ ] Close the file (not open during script execution)

## Phase 6: Testing

### Test 1: Dry Run on Sample Data
- [ ] Create test Excel with 2-3 rows
- [ ] Run with DryRun parameter:
  ```powershell
  cd C:\EmailAutomation
  
  .\Send-EmailsFromExcel.ps1 `
    -ExcelPath "C:\EmailAutomation\excel\email_list.xlsx" `
    -TemplateFolder "C:\EmailAutomation\templates" `
    -AttachmentFolder "C:\EmailAutomation\attachments" `
    -DryRun
  ```

Expected output:
- [ ] Shows "DRY RUN" mode
- [ ] Lists emails that would be created
- [ ] Shows attachments found
- [ ] No errors about missing files
- [ ] No actual emails created

### Test 2: Single Email Creation
- [ ] Run without DryRun:
  ```powershell
  .\Send-EmailsFromExcel.ps1 `
    -ExcelPath "C:\EmailAutomation\excel\email_list.xlsx" `
    -TemplateFolder "C:\EmailAutomation\templates" `
    -AttachmentFolder "C:\EmailAutomation\attachments"
  ```

- [ ] Check Outlook > Drafts folder
- [ ] Verify draft email created:
  - [ ] Correct recipient email address
  - [ ] HTML body loaded properly
  - [ ] All attachments present
  - [ ] Subject line correct

### Test 3: Multiple Emails
- [ ] Run on test Excel with 5+ rows
- [ ] Verify all emails created in Drafts
- [ ] Check varying attachments per row work
- [ ] Verify different templates load correctly

### Test 4: Error Handling
- [ ] Test with missing attachment:
  - [ ] Add a filename that doesn't exist
  - [ ] Script should warn but continue
  
- [ ] Test with missing template:
  - [ ] Reference non-existent template
  - [ ] Script should skip that row with warning

## Phase 7: Pre-Production Validation

### Final Checks
- [ ] All templates are correct and professionally formatted
- [ ] All attachments are final versions (not drafts)
- [ ] Email addresses are verified and correct
- [ ] No test data in production Excel
- [ ] Templates translated correctly (if multi-language)

### Security Review
- [ ] Excel file access restricted to authorized users
- [ ] Attachments folder access secured
- [ ] Template folder backed up
- [ ] No sensitive data in plain text

### Documentation
- [ ] Keep copy of this checklist
- [ ] Document your folder structure
- [ ] Document any custom modifications made
- [ ] Record template-to-file mappings

## Phase 8: Production Deployment

### First Production Run
- [ ] Close all Office applications
- [ ] Open PowerShell as Administrator
- [ ] Navigate to script folder:
  ```powershell
  cd C:\EmailAutomation
  ```

- [ ] Run production script:
  ```powershell
  .\Send-EmailsFromExcel.ps1 `
    -ExcelPath "C:\EmailAutomation\excel\production_emails.xlsx" `
    -TemplateFolder "C:\EmailAutomation\templates" `
    -AttachmentFolder "C:\EmailAutomation\attachments" `
    -Verbose
  ```

- [ ] Review all created drafts in Outlook
- [ ] Verify attachments for each email
- [ ] Check formatting and content
- [ ] Send or schedule sending

### Post-Run Tasks
- [ ] Archive the Excel file used
- [ ] Log the run details (date, number of emails, any issues)
- [ ] Update records of processing
- [ ] Clean up Drafts (delete test emails)

## Phase 9: Ongoing Maintenance

### After Each Run
- [ ] Archive processed Excel files
- [ ] Keep log of runs performed
- [ ] Monitor for any issues reported
- [ ] Update attachment files as needed

### Regular Tasks
- [ ] Review and update templates as needed
- [ ] Keep attachments current
- [ ] Maintain clean folder structure
- [ ] Backup all configuration files

### Troubleshooting Log
Keep a record of any issues and resolutions:
```
Date: _______________
Issue: _______________
Resolution: _______________
Notes: _______________
```

## Quick Start (After Setup Complete)

Once everything is set up, running the script is simple:

1. **Prepare Excel file** with recipient data
2. **Ensure templates exist** in templates folder
3. **Ensure attachments exist** in attachments folder
4. **Run the script:**
   ```powershell
   cd C:\EmailAutomation
   .\Send-EmailsFromExcel.ps1 `
     -ExcelPath "C:\EmailAutomation\excel\email_list.xlsx" `
     -TemplateFolder "C:\EmailAutomation\templates" `
     -AttachmentFolder "C:\EmailAutomation\attachments"
   ```
5. **Review drafts** in Outlook
6. **Send** when ready

## Common Issues & Solutions

| Issue | Solution |
|-------|----------|
| Script won't run | Check execution policy and run as Admin |
| COM object errors | Close Excel/Outlook and retry |
| Files not found | Verify full paths and file names |
| Template not loading | Check HTML file name matches exactly |
| Attachments missing | Verify files exist in attachments folder |
| Outlook freezes | Close and reopen Outlook, reduce batch size |

## Support Contacts

For issues with:
- **PowerShell**: Microsoft PowerShell documentation or Stack Overflow
- **Outlook**: Microsoft Outlook support
- **Excel**: Microsoft Excel support
- **Script**: Review USAGE_GUIDE.md or QUICK_REFERENCE.md

## Success Criteria

✓ You can run the script successfully
✓ Draft emails appear in Outlook Drafts
✓ All attachments are included
✓ HTML formatting displays correctly
✓ Variable attachment handling works
✓ Error messages are clear and helpful
✓ Dry run mode works without creating emails
✓ You can modify templates and scripts easily
