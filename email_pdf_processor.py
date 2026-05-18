"""
Automated Email PDF Processor for Hertz Gate Passes
Connects to Outlook, downloads PDFs from specific folder, and creates email drafts.
"""

import win32com.client
import os
from datetime import datetime, timedelta, date, time
from pathlib import Path
import re
from typing import Optional, Set, Tuple
from hertz_email_generator import HertzPDFProcessor

# VIN regex pattern
VIN_REGEX = re.compile(r'\b[A-HJ-NPR-Z0-9]{17}\b')


def format_outlook_datetime(dt):
    """Format datetime for Outlook filter."""
    return dt.strftime("%m/%d/%Y %I:%M %p")


def open_mail_folder(namespace, mailbox_name: Optional[str], folder_path: str):
    """
    Open a subfolder by path within a mailbox.
    
    folder_path example: "Inbox/CarMax GP"
    If mailbox_name is None, tries the default store; otherwise opens the top-level
    store by the provided mailbox name (e.g., "hertzlogistics@hertz.com").
    """
    root = None
    try:
        if mailbox_name:
            # Try exact match first
            try:
                root = namespace.Folders[mailbox_name]
            except Exception:
                # Try case-insensitive match and partial match (without @domain)
                search_name_lower = mailbox_name.lower()
                search_name_base = search_name_lower.split('@')[0]  # Get part before @
                
                for store in list(namespace.Folders):
                    store_name = getattr(store, "Name", "")
                    store_name_lower = store_name.lower()
                    
                    # Match if exact, case-insensitive, or base name matches
                    if (store_name_lower == search_name_lower or 
                        store_name_lower == search_name_base or
                        store_name == mailbox_name):
                        root = store
                        break
            if root is None:
                available = ", ".join([getattr(s, "Name", "<unknown>") for s in list(namespace.Folders)])
                raise RuntimeError(
                    f"Outlook mailbox '{mailbox_name}' not found. Available mailboxes: {available}"
                )
        else:
            root = namespace.GetDefaultFolder(6).Parent  # 6 = olFolderInbox
    except Exception:
        available = ", ".join([getattr(s, "Name", "<unknown>") for s in list(namespace.Folders)])
        raise RuntimeError(
            f"Outlook mailbox '{mailbox_name}' not found. Available mailboxes: {available}"
        )
    
    current = root
    for part in [p for p in folder_path.split("/") if p]:
        try:
            current = current.Folders[part]
        except Exception:
            raise RuntimeError(
                f"Outlook folder path not found: '{folder_path}'. Missing segment: '{part}'."
            )
    return current


def _walk_folder_paths(parent, prefix: str):
    """Yield (full_path, folder_obj) for all descendants of parent."""
    for sub in list(parent.Folders):
        path = sub.Name if not prefix else f"{prefix}/{sub.Name}"
        yield path, sub
        # Recurse
        try:
            yield from _walk_folder_paths(sub, path)
        except Exception:
            continue


def discover_folder_paths(namespace, mailbox_name: Optional[str], search_name: str):
    """Return list of folder paths whose name contains search_name (case-insensitive)."""
    try:
        if mailbox_name:
            # Try exact match first
            try:
                root = namespace.Folders[mailbox_name]
            except Exception:
                # Try case-insensitive match and partial match (without @domain)
                root = None
                search_name_lower = mailbox_name.lower()
                search_name_base = search_name_lower.split('@')[0]  # Get part before @
                
                for store in list(namespace.Folders):
                    store_name = getattr(store, "Name", "")
                    store_name_lower = store_name.lower()
                    
                    # Match if exact, case-insensitive, or base name matches
                    if (store_name_lower == search_name_lower or 
                        store_name_lower == search_name_base or
                        store_name == mailbox_name):
                        root = store
                        print(f"[DEBUG] Matched mailbox '{mailbox_name}' to '{store_name}'")
                        break
                
                if root is None:
                    available = ", ".join([getattr(s, "Name", "<unknown>") for s in list(namespace.Folders)])
                    raise RuntimeError(
                        f"Outlook mailbox '{mailbox_name}' not found. Available mailboxes: {available}"
                    )
        else:
            root = namespace.GetDefaultFolder(6).Parent
    except RuntimeError:
        raise
    except Exception:
        raise RuntimeError(
            f"Outlook mailbox '{mailbox_name}' not found. Ensure you have access in Outlook."
        )
    
    matches = []
    term = search_name.strip().lower()
    
    # DEBUG: Track all folders we find
    all_folders = []
    
    # Walk all top-level folders under the mailbox root
    for path, _folder in _walk_folder_paths(root, prefix=""):
        try:
            folder_name = path.split("/")[-1].lower()
            all_folders.append(path)
            
            # More flexible matching: check if search term is substring of folder name
            # OR if all words in search term are in folder name
            search_words = term.split()
            folder_words = folder_name.split()
            
            # Match if: exact substring match OR all search words present in folder name
            if (term in folder_name) or all(word in folder_name for word in search_words):
                matches.append(path)
        except Exception:
            continue
    
    # DEBUG: Print all folders if no matches found
    if not matches:
        print(f"[DEBUG] No matches found for '{search_name}'")
        print(f"[DEBUG] Scanned {len(all_folders)} total folders:")
        for folder_path in sorted(all_folders)[:20]:  # Show first 20
            print(f"  - {folder_path}")
        if len(all_folders) > 20:
            print(f"  ... and {len(all_folders) - 20} more")
    
    # Sort with Inbox-first preference
    matches.sort(key=lambda p: (not p.lower().startswith("inbox/"), p.lower()))
    return matches


class OutlookEmailProcessor:
    """Process emails from Outlook and download PDF attachments."""
    
    def __init__(self, email_account="hertzlogistics@hertz.com", folder_name="CarMax GP"):
        self.email_account = email_account
        self.folder_name = folder_name
        self.outlook = None
        self.namespace = None
        self.download_folder = "downloaded_pdfs"
        
    def connect_to_outlook(self):
        """Connect to Outlook application."""
        try:
            print("\n[INFO] Connecting to Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            print("[OK] Connected to Outlook successfully!")
            return True
        except Exception as e:
            print(f"[ERROR] Failed to connect to Outlook: {e}")
            print("  Make sure Outlook is installed and configured.")
            return False
    
    def find_folder(self, folder_name):
        """Find a folder by name in Outlook using improved discovery."""
        try:
            print(f"[INFO] Discovering '{folder_name}' folders in mailbox: {self.email_account}")
            
            # Discover all folders matching the name
            matches = discover_folder_paths(self.namespace, self.email_account, folder_name)
            
            if not matches:
                print(f"[WARNING] No folders matching '{folder_name}' found")
                # List available mailboxes to help user
                available = ", ".join([getattr(s, "Name", "<unknown>") for s in list(self.namespace.Folders)])
                print(f"[INFO] Available mailboxes: {available}")
                return None
            
            print(f"[OK] Found {len(matches)} matching folder(s):")
            for idx, path in enumerate(matches, 1):
                print(f"  [{idx}] {path}")
            
            # Use the first non-archive folder, or first folder if all are archives
            selected_path = matches[0]
            for path in matches:
                if "archive" not in path.lower():
                    selected_path = path
                    break
            
            print(f"[OK] Using folder: {selected_path}")
            
            # Open the selected folder
            folder = open_mail_folder(self.namespace, self.email_account, selected_path)
            print(f"[OK] Opened folder successfully")
            
            return folder
            
        except Exception as e:
            print(f"[ERROR] Error finding folder: {e}")
            return None
    
    def search_emails(self, folder, keywords, date_filter="today"):
        """
        Search emails in folder by keywords and date.
        
        Args:
            folder: Outlook folder object
            keywords: List of keywords to search for (e.g., ["KUNES", "EASTON"])
            date_filter: "today", "yesterday", "day before yesterday", or specific date (e.g., "11/7/2025" or "11/7")
        """
        try:
            # Set date filter
            if date_filter.lower() == "today":
                target_date = datetime.now().date()
            elif date_filter.lower() == "yesterday":
                target_date = (datetime.now() - timedelta(days=1)).date()
            elif date_filter.lower() == "day before yesterday":
                target_date = (datetime.now() - timedelta(days=2)).date()
            else:
                # Try to parse as specific date
                try:
                    # Try various date formats
                    for fmt in ["%m/%d/%Y", "%m/%d", "%Y-%m-%d", "%m-%d-%Y"]:
                        try:
                            parsed_date = datetime.strptime(date_filter, fmt)
                            # If year not provided, use current year
                            if fmt == "%m/%d":
                                parsed_date = parsed_date.replace(year=datetime.now().year)
                            target_date = parsed_date.date()
                            break
                        except ValueError:
                            continue
                    else:
                        # If no format matched, default to today
                        print(f"[WARNING] Could not parse date '{date_filter}', using today")
                        target_date = datetime.now().date()
                except Exception:
                    target_date = datetime.now().date()
            
            print(f"\n[INFO] Searching emails in: {folder.Name}")
            print(f"[INFO] Keywords: {', '.join(keywords)}")
            print(f"[INFO] Date: {target_date.strftime('%Y-%m-%d')}")
            
            matching_emails = []
            items = folder.Items
            
            # Sort by received time (most recent first)
            items.Sort("[ReceivedTime]", True)
            
            # Search through emails
            count = 0
            for item in items:
                try:
                    # Check if it's a mail item
                    if not hasattr(item, 'ReceivedTime'):
                        continue
                    
                    # Check date
                    received_date = item.ReceivedTime.date()
                    if received_date != target_date:
                        # Skip emails not from target date
                        # If we've gone past today's date, we can stop searching
                        if received_date < target_date:
                            break
                        continue
                    
                    # Check for keywords in subject
                    subject = item.Subject if hasattr(item, 'Subject') else ""
                    
                    # Check if any keyword is in the subject
                    if any(keyword.upper() in subject.upper() for keyword in keywords):
                        matching_emails.append(item)
                        count += 1
                        print(f"  [{count}] Found: {subject[:60]}...")
                
                except Exception as e:
                    # Skip items that cause errors
                    continue
            
            print(f"\n[OK] Found {len(matching_emails)} matching email(s)")
            return matching_emails
            
        except Exception as e:
            print(f"[ERROR] Error searching emails: {e}")
            return []
    
    def get_existing_pdfs(self, output_folder=None):
        """Get list of already downloaded PDF filenames."""
        if output_folder is None:
            output_folder = self.download_folder
        
        output_folder = os.path.abspath(output_folder)
        
        if not os.path.exists(output_folder):
            return set()
        
        # Get all PDF filenames (not full paths)
        existing_pdfs = set()
        for filename in os.listdir(output_folder):
            if filename.lower().endswith('.pdf'):
                existing_pdfs.add(filename)
        
        return existing_pdfs
    
    def download_pdfs(self, emails, output_folder=None):
        """
        Download PDF attachments from emails.
        Only downloads PDFs that don't already exist (prevents duplicates).
        
        Args:
            emails: List of Outlook email items
            output_folder: Folder to save PDFs (default: downloaded_pdfs)
        """
        if output_folder is None:
            output_folder = self.download_folder
        
        # Create output folder with absolute path if it doesn't exist
        output_folder = os.path.abspath(output_folder)
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        
        # Get list of already downloaded PDFs
        existing_pdfs = self.get_existing_pdfs(output_folder)
        
        if existing_pdfs:
            print(f"\n[INFO] Found {len(existing_pdfs)} existing PDF(s) in download folder")
            print(f"[INFO] Will skip duplicates and only download new PDFs")
        
        print(f"\n[INFO] Downloading PDFs to: {os.path.abspath(output_folder)}")
        
        downloaded_pdfs = []
        skipped_pdfs = []
        
        for idx, email in enumerate(emails, 1):
            try:
                subject = email.Subject if hasattr(email, 'Subject') else "Unknown"
                print(f"\n[{idx}/{len(emails)}] Processing: {subject[:60]}...")
                
                if email.Attachments.Count == 0:
                    print("  [INFO] No attachments found")
                    continue
                
                # Process each attachment
                for attachment in email.Attachments:
                    filename = attachment.FileName
                    
                    # Check if it's a PDF
                    if filename.lower().endswith('.pdf'):
                        # Check if this PDF already exists (skip duplicates)
                        if filename in existing_pdfs:
                            print(f"  [SKIP] Already exists: {filename}")
                            skipped_pdfs.append(filename)
                            continue
                        
                        # Save to output folder
                        filepath = os.path.join(output_folder, filename)
                        
                        # Double-check file doesn't exist (safety check)
                        if os.path.exists(filepath):
                            print(f"  [SKIP] File already exists: {filename}")
                            skipped_pdfs.append(filename)
                            continue
                        
                        attachment.SaveAsFile(filepath)
                        downloaded_pdfs.append(filepath)
                        print(f"  [OK] Downloaded: {filename} (NEW)")
                    else:
                        print(f"  [SKIP] Not a PDF: {filename}")
                
            except Exception as e:
                print(f"  [ERROR] Failed to process email: {e}")
                continue
        
        print(f"\n[OK] Downloaded {len(downloaded_pdfs)} NEW PDF(s)")
        if skipped_pdfs:
            print(f"[INFO] Skipped {len(skipped_pdfs)} existing PDF(s) (already downloaded)")
        
        return downloaded_pdfs
    
    def process_pdfs(self, pdf_files, create_email_drafts=True):
        """
        Process downloaded PDFs and optionally create email drafts.
        
        Args:
            pdf_files: List of PDF file paths
            create_email_drafts: Whether to create Outlook drafts
        """
        print("\n" + "="*80)
        print("PROCESSING DOWNLOADED PDFs")
        print("="*80)
        
        results = []
        
        for idx, pdf_path in enumerate(pdf_files, 1):
            print(f"\n[{idx}/{len(pdf_files)}] Processing: {os.path.basename(pdf_path)}")
            print("-" * 80)
            
            try:
                # Create processor
                processor = HertzPDFProcessor(pdf_path)
                processor.extract_text()
                processor.parse_data()
                
                # Show summary
                print(f"  VIN:      {processor.data.get('vin', 'N/A')}")
                print(f"  Vehicle:  {processor.data.get('year', 'N/A')} {processor.data.get('make', 'N/A')} {processor.data.get('model', 'N/A')}")
                print(f"  Location: {processor.data.get('full_location', 'N/A')}")
                
                hertz_emails = processor.data.get('hertz_rep_emails', [])
                excluded_emails = processor.data.get('excluded_emails', [])
                
                # Get CC recipient from config
                cc_recipients = "joshua.blankenship@hertz.com"
                try:
                    import email_config
                    if hasattr(email_config, 'CC_RECIPIENTS'):
                        cc_recipients = email_config.CC_RECIPIENTS
                except ImportError:
                    pass
                
                print(f"  To:       {', '.join(hertz_emails) if hertz_emails else 'None'}")
                print(f"  CC:       {cc_recipients}")
                if excluded_emails:
                    print(f"  Excluded: {', '.join(excluded_emails)}")
                
                # Create email draft if requested
                if create_email_drafts:
                    success = processor.create_outlook_draft()
                    results.append({
                        'file': os.path.basename(pdf_path),
                        'success': success,
                        'data': processor.data
                    })
                else:
                    results.append({
                        'file': os.path.basename(pdf_path),
                        'success': True,
                        'data': processor.data
                    })
                
                print(f"  [OK] Processed successfully")
                
            except Exception as e:
                print(f"  [ERROR] Failed to process: {e}")
                results.append({
                    'file': os.path.basename(pdf_path),
                    'success': False,
                    'error': str(e)
                })
        
        return results
    
    def generate_report(self, results, output_file="email_processing_report.txt"):
        """Generate a summary report of processing results."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("HERTZ EMAIL PDF PROCESSING REPORT\n")
            f.write("="*80 + "\n")
            f.write(f"\nTimestamp: {timestamp}\n")
            f.write(f"Total PDFs: {len(results)}\n")
            f.write(f"Successful: {sum(1 for r in results if r['success'])}\n")
            f.write(f"Failed: {sum(1 for r in results if not r['success'])}\n")
            f.write("\n" + "="*80 + "\n")
            f.write("DETAILED RESULTS\n")
            f.write("="*80 + "\n")
            
            for idx, result in enumerate(results, 1):
                f.write(f"\n[{idx}] {result['file']}\n")
                f.write("-" * 80 + "\n")
                
                if result['success']:
                    data = result.get('data', {})
                    f.write(f"Status: SUCCESS\n")
                    f.write(f"VIN: {data.get('vin', 'N/A')}\n")
                    f.write(f"Vehicle: {data.get('year', 'N/A')} {data.get('make', 'N/A')} {data.get('model', 'N/A')}\n")
                    f.write(f"Location: {data.get('full_location', 'N/A')}\n")
                    f.write(f"Recipients: {', '.join(data.get('hertz_rep_emails', []))}\n")
                    if data.get('excluded_emails'):
                        f.write(f"Excluded: {', '.join(data.get('excluded_emails', []))}\n")
                else:
                    f.write(f"Status: FAILED\n")
                    f.write(f"Error: {result.get('error', 'Unknown error')}\n")
        
        print(f"\n[OK] Report saved to: {output_file}")


def main():
    """Main automation workflow."""
    print("="*80)
    print("HERTZ EMAIL PDF PROCESSOR - AUTOMATED WORKFLOW")
    print("="*80)
    
    # Import configuration
    try:
        import email_config as config
        email_account = config.EMAIL_ACCOUNT
        folder_name = config.FOLDER_NAME
        keywords = config.KEYWORDS
        date_filter = config.DATE_FILTER
        create_drafts = config.CREATE_EMAIL_DRAFTS
        generate_report = config.GENERATE_REPORT
        report_file = config.REPORT_FILENAME
    except ImportError:
        print("[WARNING] email_config.py not found, using defaults")
        email_account = "hertzlogistics@hertz.com"
        folder_name = "CarMax GP"
        keywords = ["KUNES", "EASTON"]
        date_filter = "today"
        create_drafts = True
        generate_report = True
        report_file = "email_processing_report.txt"
    
    print(f"\nConfiguration:")
    print(f"  Email Account: {email_account}")
    print(f"  Folder: {folder_name}")
    print(f"  Keywords: {', '.join(keywords)}")
    print(f"  Date Filter: {date_filter}")
    
    # Create processor
    processor = OutlookEmailProcessor(email_account, folder_name)
    
    # Step 1: Connect to Outlook
    if not processor.connect_to_outlook():
        print("\n[ERROR] Cannot proceed without Outlook connection")
        return
    
    # Step 2: Find the folder
    folder = processor.find_folder(folder_name)
    if not folder:
        print(f"\n[ERROR] Cannot find folder: {folder_name}")
        print("\nAvailable folders:")
        # List some folders to help user
        try:
            for account in processor.namespace.Folders:
                print(f"  - {account.Name}")
                if account.Folders.Count > 0:
                    for subfolder in account.Folders:
                        print(f"    - {subfolder.Name}")
        except:
            pass
        return
    
    # Step 3: Search for emails
    emails = processor.search_emails(folder, keywords, date_filter)
    
    if not emails:
        print("\n[INFO] No matching emails found")
        print("  Try checking:")
        print("    - Date filter (are there emails from today?)")
        print("    - Keywords (are they in email subjects?)")
        print("    - Folder name (is it correct?)")
        return
    
    # Step 4: Download PDFs
    pdf_files = processor.download_pdfs(emails)
    
    if not pdf_files:
        print("\n[WARNING] No PDFs downloaded")
        return
    
    # Step 5: Process PDFs and create email drafts
    results = processor.process_pdfs(pdf_files, create_email_drafts=create_drafts)
    
    # Step 6: Generate report
    if generate_report:
        processor.generate_report(results, report_file)
    
    # Summary
    print("\n" + "="*80)
    print("AUTOMATION COMPLETE")
    print("="*80)
    print(f"\nEmails Found: {len(emails)}")
    print(f"New PDFs Downloaded: {len(pdf_files)}")
    print(f"Successfully Processed: {sum(1 for r in results if r['success'])}")
    print(f"Failed: {sum(1 for r in results if not r['success'])}")
    
    if len(pdf_files) > 0:
        print(f"\n[OK] Email drafts created: Check your Outlook!")
        print(f"     {len(pdf_files)} new draft(s) ready to review and send")
    else:
        print(f"\n[INFO] No new PDFs to process")
        print(f"     All gate passes from today have already been processed")
        print(f"     Re-run later if new emails arrive")
    
    print("="*80)


if __name__ == "__main__":
    main()

