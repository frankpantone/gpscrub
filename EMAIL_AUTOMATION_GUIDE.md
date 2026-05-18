# Email Automation Guide - Hertz Gate Pass Processor

## 🎯 What This Does

**Completely automates your daily workflow:**

1. ✅ **Connects** to your Outlook email (hertzlogistics@hertz.com)
2. ✅ **Searches** CarMax GP Folder for emails with "KUNES" and "EASTON"
3. ✅ **Filters** to only emails received TODAY
4. ✅ **Downloads** all PDF attachments
5. ✅ **Processes** each PDF to extract vehicle data
6. ✅ **Excludes** purchaser emails automatically
7. ✅ **Creates** Outlook email drafts with correct recipients and deadline
8. ✅ **Attaches** PDFs to each draft
9. ✅ **Generates** summary report

**Result:** All email drafts ready to review and send - no manual work needed!

---

## 🚀 Quick Start (3 Steps)

### Step 1: Make sure Outlook is running
- Outlook must be open (doesn't have to be in focus)

### Step 2: Double-click the launcher
```
Double-click: process_today_emails.bat
```

### Step 3: Review drafts in Outlook
- Email drafts will open automatically
- Review and click Send

**That's it!** ✨

---

## 📁 Files Created

### Core Scripts
- **`email_pdf_processor.py`** - Main automation script
- **`process_today_emails.bat`** - Windows launcher (double-click to run)
- **`email_config.py`** - Configuration file (edit keywords, dates, etc.)

### Auto-Generated
- **`downloaded_pdfs/`** - Folder where PDFs are saved
- **`email_processing_report.txt`** - Detailed processing report

---

## ⚙️ Configuration

Edit `email_config.py` to customize:

### Email Settings
```python
EMAIL_ACCOUNT = "hertzlogistics@hertz.com"  # Your email account
FOLDER_NAME = "CarMax GP"                   # Folder to search
```

### Search Settings
```python
KEYWORDS = ["KUNES", "EASTON"]              # Keywords in email subjects
DATE_FILTER = "today"                       # "today", "yesterday", or "2025-10-28"
```

### Processing Settings
```python
CREATE_EMAIL_DRAFTS = True                  # Auto-create drafts
GENERATE_REPORT = True                      # Create summary report
DELETE_PDFS_AFTER_PROCESSING = False        # Keep downloaded PDFs
```

---

## 📊 Example Workflow

### Input: Email in CarMax GP Folder
```
From: sender@example.com
Subject: Gate Pass - KUNES Country Honda - VIN KNDPU3AF3P7143419
Date: Today
Attachments: GR4031029_KNDPU3AF3P7143419.pdf
```

### Automation Process:
1. Script connects to Outlook ✅
2. Finds CarMax GP folder ✅
3. Finds email with "KUNES" in subject ✅
4. Downloads PDF to `downloaded_pdfs/` ✅
5. Extracts: VIN, vehicle info, emails, location ✅
6. Creates Outlook draft:
   - **To:** BSwift@hertz.com; Khadijah.Goudeau@hertz.com
   - **Subject:** HERTZ LOGISTICS CONFIRMATION - VIN # (...) - RESPONSE NEEDED BY 1100 on [tomorrow]
   - **Attachment:** PDF attached
   - **Body:** Pre-filled with vehicle details and questions ✅

### Output: Ready-to-send email draft in Outlook

---

## 📋 Sample Output

```
================================================================================
HERTZ EMAIL PDF PROCESSOR - AUTOMATED WORKFLOW
================================================================================

Configuration:
  Email Account: hertzlogistics@hertz.com
  Folder: CarMax GP
  Keywords: KUNES, EASTON
  Date Filter: today

[INFO] Connecting to Outlook...
[OK] Connected to Outlook successfully!

[INFO] Searching in store: hertzlogistics@hertz.com
[OK] Found account: hertzlogistics@hertz.com
[OK] Found folder: \\hertzlogistics@hertz.com\Inbox\CarMax GP

[INFO] Searching emails in: CarMax GP
[INFO] Keywords: KUNES, EASTON
[INFO] Date: 2025-10-28

  [1] Found: Gate Pass - KUNES Country Honda - VIN KNDPU3AF3P7143419
  [2] Found: Gate Pass - EASTON Ford - VIN 1FTFW1E85MFC12345

[OK] Found 2 matching email(s)

[INFO] Downloading PDFs to: C:\...\downloaded_pdfs

[1/2] Processing: Gate Pass - KUNES Country Honda...
  [OK] Downloaded: GR4031029_KNDPU3AF3P7143419.pdf

[2/2] Processing: Gate Pass - EASTON Ford...
  [OK] Downloaded: GR4567890_1FTFW1E85MFC12345.pdf

[OK] Downloaded 2 PDF(s) total

================================================================================
PROCESSING DOWNLOADED PDFs
================================================================================

[1/2] Processing: GR4031029_KNDPU3AF3P7143419.pdf
--------------------------------------------------------------------------------
  VIN:      KNDPU3AF3P7143419
  Vehicle:  2023 Kia Sportage
  Location: 2825 W. Perimeter Rd Indianapolis, IN 46241
  To:       BSwift@hertz.com, Khadijah.Goudeau@hertz.com
  Excluded: konovak@hertz.com
  [OK] PDF attached: GR4031029_KNDPU3AF3P7143419.pdf
  [OK] Outlook email draft created successfully!
  [OK] Processed successfully

[2/2] Processing: GR4567890_1FTFW1E85MFC12345.pdf
--------------------------------------------------------------------------------
  VIN:      1FTFW1E85MFC12345
  Vehicle:  2021 Ford F-150
  Location: 1234 Main Street Columbus, OH 43215
  To:       logistics@hertz.com
  Excluded: dealer@easton.com
  [OK] PDF attached: GR4567890_1FTFW1E85MFC12345.pdf
  [OK] Outlook email draft created successfully!
  [OK] Processed successfully

[OK] Report saved to: email_processing_report.txt

================================================================================
AUTOMATION COMPLETE
================================================================================

Emails Found: 2
PDFs Downloaded: 2
Successfully Processed: 2
Failed: 0

Email drafts created: Check your Outlook!
================================================================================
```

---

## 🔍 How It Finds Emails

### Email Matching Logic:

1. **Folder Search**: Looks in "CarMax GP" folder
2. **Date Filter**: Only emails received TODAY
3. **Keyword Search**: Email subject must contain ANY of the keywords:
   - "KUNES" (case-insensitive)
   - "EASTON" (case-insensitive)
4. **Has Attachments**: Email must have PDF attachments

### Example Matching Subjects:
- ✅ "Gate Pass - KUNES Country Honda"
- ✅ "KUNES HONDA - Vehicle Pickup"
- ✅ "EASTON Ford Dealership - Gate Release"
- ✅ "Daily Report - KUNES and EASTON"
- ❌ "Gate Pass - OTHER Dealer" (no matching keyword)
- ❌ "KUNES Report from yesterday" (wrong date)

---

## 🛠️ Troubleshooting

### "Cannot connect to Outlook"
**Solution:**
- Make sure Microsoft Outlook is installed
- Open Outlook and let it load completely
- Run the script again

### "Folder 'CarMax GP' not found"
**Solution:**
- Check the exact folder name in Outlook
- Edit `email_config.py` to match the exact name
- Check if folder is under "hertzlogistics@hertz.com" account

**To list all folders:**
The script will show available folders if it can't find the target folder.

### "No matching emails found"
**Solution:**
- Check if there are emails from TODAY
- Verify keywords are in email subjects
- Try changing DATE_FILTER to "yesterday" for testing
- Check if emails have PDF attachments

### "Downloaded PDFs but no drafts created"
**Solution:**
- Check if `CREATE_EMAIL_DRAFTS = True` in `email_config.py`
- Review console output for errors
- Check `email_processing_report.txt` for details

### PDFs not downloading
**Solution:**
- Verify emails have attachments
- Check if attachments are PDFs
- Ensure you have write permissions in the folder

---

## 📝 Advanced Usage

### Search Multiple Keywords
Edit `email_config.py`:
```python
KEYWORDS = ["KUNES", "EASTON", "ANOTHER_DEALER", "FOURTH_DEALER"]
```

### Search Yesterday's Emails
```python
DATE_FILTER = "yesterday"
```

### Search Specific Date
```python
DATE_FILTER = "2025-10-28"
```

### Change Download Folder
```python
DOWNLOAD_FOLDER = "my_custom_folder"
```

### Extract Data Only (No Email Drafts)
```python
CREATE_EMAIL_DRAFTS = False
```

### Different Email Account
```python
EMAIL_ACCOUNT = "different@hertz.com"
```

### Different Folder
```python
FOLDER_NAME = "Different Folder Name"
```

---

## 🔄 Daily Workflow

### Morning Routine (Automated)
1. Check your hertzlogistics@hertz.com inbox
2. Emails with gate passes arrive in CarMax GP folder
3. Double-click `process_today_emails.bat`
4. Script finds and processes all matching emails
5. Review email drafts in Outlook
6. Send drafts

**Time saved:** 5-10 minutes per PDF → 2-3 seconds per PDF ✨

---

## 📊 Processing Report

After each run, check `email_processing_report.txt`:

```
================================================================================
HERTZ EMAIL PDF PROCESSING REPORT
================================================================================

Timestamp: 2025-10-28 09:15:30
Total PDFs: 2
Successful: 2
Failed: 0

================================================================================
DETAILED RESULTS
================================================================================

[1] GR4031029_KNDPU3AF3P7143419.pdf
--------------------------------------------------------------------------------
Status: SUCCESS
VIN: KNDPU3AF3P7143419
Vehicle: 2023 Kia Sportage
Location: 2825 W. Perimeter Rd Indianapolis, IN 46241
Recipients: BSwift@hertz.com, Khadijah.Goudeau@hertz.com
Excluded: konovak@hertz.com

[2] GR4567890_1FTFW1E85MFC12345.pdf
--------------------------------------------------------------------------------
Status: SUCCESS
VIN: 1FTFW1E85MFC12345
Vehicle: 2021 Ford F-150
Location: 1234 Main Street Columbus, OH 43215
Recipients: logistics@hertz.com
Excluded: dealer@easton.com
```

---

## 🎯 Benefits

### Time Savings
- **Manual:** ~5 minutes per PDF
- **Automated:** ~3 seconds per PDF
- **10 PDFs/day:** Save ~50 minutes daily! ⏰

### Accuracy
- ✅ No typos in VINs
- ✅ No wrong recipients
- ✅ Consistent email format
- ✅ Correct deadlines every time

### Consistency
- ✅ Same process every time
- ✅ No missed emails
- ✅ Always follows template
- ✅ Audit trail in reports

---

## 🔐 Security & Privacy

- ✅ **Local processing** - All data stays on your PC
- ✅ **No external services** - Doesn't send data anywhere
- ✅ **Uses your Outlook** - Same security as manual process
- ✅ **Draft-only** - Never sends automatically
- ✅ **You control** - Review before sending

---

## 📞 Support

### For Help:
1. Check console output for error messages
2. Review `email_processing_report.txt`
3. Check `downloaded_pdfs/` folder
4. Verify `email_config.py` settings

### Common Issues:
- **Folder not found** → Check folder name in Outlook
- **No emails found** → Verify date and keywords
- **No PDFs** → Check if emails have attachments
- **Outlook error** → Restart Outlook

---

## 🔄 Version History

### v1.2 (Current) - October 28, 2025
- ✅ Email automation added
- ✅ Automatic PDF download
- ✅ Folder search functionality
- ✅ Keyword and date filtering
- ✅ Batch processing from email
- ✅ Configuration file support

### v1.1 - October 28, 2025
- ✅ Purchaser email exclusion
- ✅ Fixed deadline format (1100 next day)
- ✅ PDF auto-attachment

### v1.0 - October 28, 2025
- Initial release
- Manual PDF processing

---

## ✅ Quick Reference

| Task | Command |
|------|---------|
| **Run daily automation** | Double-click `process_today_emails.bat` |
| **Change keywords** | Edit `email_config.py` → `KEYWORDS` |
| **Change date** | Edit `email_config.py` → `DATE_FILTER` |
| **Change folder** | Edit `email_config.py` → `FOLDER_NAME` |
| **View report** | Open `email_processing_report.txt` |
| **Find PDFs** | Check `downloaded_pdfs/` folder |

---

**Status:** ✅ Fully Automated End-to-End Solution  
**Last Updated:** October 28, 2025  
**Version:** 1.2




