# Hertz PDF Scrubber & Email Generator

Automated tool to extract vehicle information, pickup location, and Hertz representative emails from gate release PDFs and generate Outlook email drafts for logistics confirmation.

## Features

- ✅ **PDF Data Extraction**: Automatically extracts:
  - Hertz Representative emails
  - Vehicle Information (VIN, Year, Make, Model)
  - Pickup Location (address, city, state, zip)
  - Contact phone numbers
  - Order numbers

- ✅ **Smart Parsing**: 
  - Distinguishes between vehicle makes and dealership names
  - Filters out field labels from actual data
  - Excludes dates from model year detection

- ✅ **Outlook Integration**: 
  - Creates email drafts directly in Microsoft Outlook
  - Pre-populates recipients, subject, and body
  - Automatically attaches the PDF file
  - Follows standard logistics confirmation format

- ✅ **Smart Email Filtering**:
  - Automatically excludes purchaser emails (first email in document)
  - Includes only Hertz representative emails as recipients
  - Prevents accidental emails to purchasers

## Installation

### Requirements
- Python 3.7+
- Microsoft Outlook installed and configured

### Install Dependencies

```bash
pip install pypdf pywin32
```

## Usage

### Basic Usage

Process a single PDF (interactive mode):
```bash
python hertz_email_generator.py
```

Process a specific PDF file:
```bash
python hertz_email_generator.py path/to/file.pdf
```

Auto-create email draft without confirmation:
```bash
python hertz_email_generator.py --auto
```

Or with a specific file:
```bash
python hertz_email_generator.py GR4031029_KNDPU3AF3P7143419.pdf --auto
```

### Command-Line Options

- `--auto` or `-a`: Automatically create Outlook draft without user confirmation
- No arguments: Processes the first PDF found in the current directory (interactive mode)
- PDF filename: Processes the specified PDF file

### Batch Processing

Process multiple PDFs:
```bash
python batch_process_pdfs.py
```

This will:
1. Find all PDFs in the current directory
2. Extract data from each
3. Create email drafts for each vehicle
4. Generate a summary report

## Email Template

The script generates emails with the following format:

**Subject:**
```
HERTZ LOGISTICS CONFIRMATION - VIN # (KNDPU3AF3P7143419) - CONDITION/LOCATION - RESPONSE NEEDED BY 1100 ON 10/29
```
Note: The deadline is always set to 11:00 AM (1100) on the next day.

**Body:**
```
Hello,
 
Please confirm this unit is on site at 2825 W. Perimeter Rd Indianapolis, IN 46241 – Hertz is transporting to the dealer 
 
KNDPU3AF3P7143419 – 2023 Kia Sportage
 
Please answer the following questions regarding the status of the unit so we can ensure its in good condition to be sold.
 
1. Is there any damage on the body of this vehicle? Y or N
2. Is there any glass damage (chips or cracks)? Y or N
3. Is there any paint mismatch (likely from previous repair) Y or N
4. Are all the tires in good condition (no flats or low tires)? Y or N
5. Is the interior clean and debris free? Y or N
 
Best regards,
```

**To:** Automatically populated with Hertz representative emails found in the PDF (excludes purchaser emails)

**CC:** joshua.blankenship@hertz.com (automatically CC'd on every email)

**Attachment:** The original PDF is automatically attached to the email

## Extracted Data Examples

From a typical gate release PDF:

```
Hertz Representative Emails (TO Recipients):
  - Khadijah.Goudeau@hertz.com
  - BSwift@hertz.com

Excluded Emails (Purchaser):
  - konovak@hertz.com [EXCLUDED - First email/Purchaser]

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

## Customization

### Adjust Response Deadline

Edit the `generate_response_deadline()` method in `hertz_email_generator.py`:

```python
def generate_response_deadline(self):
    """Generate response deadline timestamp - 1100 on next day."""
    tomorrow = datetime.now() + timedelta(days=1)
    return tomorrow.strftime("1100 on %m/%d")
```

Currently set to 11:00 AM (1100) on the next day. Change `"1100"` to your preferred time.

### Add More Vehicle Makes

Edit the `makes` list in the `parse_data()` method:

```python
makes = ['CHEVROLET', 'FORD', 'TOYOTA', 'HONDA', 'KIA', 'HYUNDAI', 
         'YOUR_MAKE_HERE', ...]
```

### Customize Email Body

Edit the `create_email_body()` method to adjust the email template.

## Troubleshooting

### "No PDF files found"
- Ensure you're in the correct directory
- Specify the PDF path directly: `python hertz_email_generator.py path/to/file.pdf`

### "Error creating Outlook draft"
- Make sure Microsoft Outlook is installed
- Ensure Outlook is configured with at least one email account
- Check that Outlook is not blocking COM automation (security settings)

### Incorrect Data Extraction
- The script works best with standard Hertz gate release forms
- If data is missing or incorrect, check the extracted text file: `pdf_extracted_text.txt`
- You can manually edit the extraction patterns in the `parse_data()` method

### Unicode/Encoding Errors
- This should be fixed in the latest version
- If issues persist, ensure your terminal supports UTF-8

## Files Generated

- `pdf_extracted_text.txt`: Raw text extracted from the last processed PDF
- `pdf_structure_analysis.md`: Detailed structural analysis (from initial analysis)
- `field_mapping.txt`: Line-by-line field mapping (from initial analysis)

## Security & Privacy

- **Email Scrubbing**: The script identifies all email addresses in PDFs
- **Local Processing**: All processing happens locally on your machine
- **No Data Storage**: No data is sent to external services
- **Draft Only**: Emails are created as drafts, never sent automatically

## Support

For issues or questions:
1. Check the console output for error messages
2. Review the `pdf_extracted_text.txt` file to see what was extracted
3. Verify your PDF follows the standard Hertz gate release format

## Version History

### v1.0 (Current)
- Initial release
- PDF text extraction
- Smart vehicle data parsing
- Outlook email draft generation
- Batch processing support
- Comprehensive error handling

## License

This tool is provided as-is for Hertz internal use.

---

**Note**: This tool is designed for Hertz gate release PDFs. Results may vary with other document formats.

