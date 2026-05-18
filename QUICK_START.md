# Quick Start Guide

## First Time Setup

1. **Install Python** (if not already installed)
   - Download from: https://www.python.org/downloads/
   - Make sure to check "Add Python to PATH" during installation

2. **Install Required Packages**
   ```bash
   pip install pypdf pywin32
   ```

3. **Ensure Outlook is Configured**
   - Microsoft Outlook must be installed
   - At least one email account must be configured
   - Outlook does not need to be running (it will launch automatically)

## Processing a Single PDF

### Method 1: Drag and Drop (Simplest)
1. Drag your PDF file into the folder with the scripts
2. Double-click `hertz_email_generator.py`
3. Follow the prompts

### Method 2: Command Line
```bash
# Navigate to the folder
cd "C:\path\to\gatepassscrub"

# Process the PDF (interactive)
python hertz_email_generator.py your_file.pdf

# Or auto-create draft (no prompts)
python hertz_email_generator.py your_file.pdf --auto
```

## Processing Multiple PDFs

### Method 1: Batch Process All PDFs in Folder
```bash
# Place all PDFs in the folder, then run:
python batch_process_pdfs.py --auto
```

### Method 2: Process PDFs from Another Folder
```bash
python batch_process_pdfs.py "C:\path\to\pdf\folder" --auto
```

## What Happens?

1. ✅ Script extracts data from PDF:
   - VIN
   - Year, Make, Model
   - Pickup location
   - Hertz rep emails
   - Contact info

2. ✅ Creates an Outlook email draft:
   - **To:** Hertz representative emails
   - **Subject:** HERTZ LOGISTICS CONFIRMATION - VIN # (...)
   - **Body:** Pre-filled with vehicle details and 5 questions

3. ✅ The draft opens in Outlook for you to:
   - Review the information
   - Make any edits if needed
   - Click "Send" when ready

## Example Output

```
================================================================================
EXTRACTED DATA SUMMARY
================================================================================

PDF File: GR4031029_KNDPU3AF3P7143419.pdf

Hertz Representative Emails:
  - Khadijah.Goudeau@hertz.com
  - BSwift@hertz.com

Vehicle Information:
  VIN:   KNDPU3AF3P7143419
  Year:  2023
  Make:  Kia
  Model: Sportage

Pickup Location:
  2825 W. Perimeter Rd Indianapolis, IN 46241

Contact Phone:
  (317) 775-8055

Order Number:
  4031029
```

## Common Commands

| Task | Command |
|------|---------|
| Process one PDF (interactive) | `python hertz_email_generator.py file.pdf` |
| Process one PDF (auto) | `python hertz_email_generator.py file.pdf --auto` |
| Process all PDFs in folder | `python batch_process_pdfs.py --auto` |
| Extract data only (no emails) | `python batch_process_pdfs.py --no-email` |
| Get help | `python hertz_email_generator.py --help` |

## Tips

- ✅ **Always review drafts** before sending
- ✅ The script creates **drafts only** - it never sends automatically
- ✅ You can edit the draft before sending
- ✅ Multiple emails found? All will be added as recipients
- ✅ Incorrect data? You can manually edit the draft or the source PDF

## Troubleshooting

### "No module named 'pypdf'"
**Solution:** Run `pip install pypdf pywin32`

### "Error creating Outlook draft"
**Solution:** 
- Make sure Outlook is installed
- Ensure Outlook is configured with an email account
- Try opening Outlook manually first

### Incorrect vehicle make/model
**Solution:** 
- Check the PDF for clarity
- Review `pdf_extracted_text.txt` to see what was extracted
- Manually edit the draft before sending

### Script won't run
**Solution:**
- Make sure you're in the correct folder
- Try: `python hertz_email_generator.py --help`
- Check Python is installed: `python --version`

## Advanced Usage

### Custom Response Deadline
Edit line 155 in `hertz_email_generator.py`:
```python
def generate_response_deadline(self, hours_from_now=24):
```
Change `24` to your desired number of hours.

### Add Vehicle Makes
Edit lines 62-65 in `hertz_email_generator.py` to add more vehicle makes to the list.

## File Structure

After running, you'll see these files:

```
gatepassscrub/
├── hertz_email_generator.py     ← Main script
├── batch_process_pdfs.py        ← Batch processor
├── analyze_pdf.py               ← Analysis tool
├── README.md                    ← Full documentation
├── QUICK_START.md              ← This file
├── GR4031029_*.pdf             ← Your PDF files
├── pdf_extracted_text.txt       ← Last extracted text
└── batch_report_*.txt          ← Batch processing report
```

## Next Steps

1. ✅ Test with one PDF first
2. ✅ Review the generated draft
3. ✅ If all looks good, batch process remaining PDFs
4. ✅ Review each draft before sending

---

**Need Help?** Check the full README.md for detailed documentation.




