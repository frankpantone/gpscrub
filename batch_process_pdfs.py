"""
Batch PDF Processor for Hertz Gate Release Forms
Process multiple PDFs and create email drafts for each.
"""

import sys
from pathlib import Path
from datetime import datetime
from hertz_email_generator import HertzPDFProcessor


def batch_process(directory=".", auto_create=True, create_emails=True):
    """
    Process all PDFs in a directory.
    
    Args:
        directory: Directory to search for PDFs (default: current directory)
        auto_create: If True, create drafts without confirmation
        create_emails: If True, create Outlook drafts; if False, only extract data
    """
    print("="*80)
    print("HERTZ BATCH PDF PROCESSOR")
    print("="*80)
    
    # Find all PDF files
    pdf_path = Path(directory)
    pdf_files = list(pdf_path.glob('*.pdf'))
    
    if not pdf_files:
        print(f"\n[ERROR] No PDF files found in: {directory}")
        return
    
    print(f"\nFound {len(pdf_files)} PDF file(s)")
    print("-" * 80)
    
    results = []
    successful = 0
    failed = 0
    
    for idx, pdf_file in enumerate(pdf_files, 1):
        print(f"\n[{idx}/{len(pdf_files)}] Processing: {pdf_file.name}")
        print("-" * 80)
        
        try:
            # Process the PDF
            processor = HertzPDFProcessor(str(pdf_file))
            processor.extract_text()
            processor.parse_data()
            
            # Store results
            result = {
                'file': pdf_file.name,
                'success': True,
                'data': processor.data.copy()
            }
            
            # Print summary
            print(f"  VIN:        {processor.data.get('vin', 'N/A')}")
            print(f"  Vehicle:    {processor.data.get('year', 'N/A')} {processor.data.get('make', 'N/A')} {processor.data.get('model', 'N/A')}")
            print(f"  Location:   {processor.data.get('full_location', 'N/A')}")
            print(f"  Emails:     {', '.join(processor.data.get('hertz_rep_emails', ['N/A']))}")
            
            # Create email draft if requested
            if create_emails:
                if auto_create:
                    success = processor.create_outlook_draft()
                    result['email_created'] = success
                else:
                    # Ask for confirmation
                    response = input(f"\n  Create email draft for {processor.data.get('vin', 'this vehicle')}? (Y/n): ").strip().lower()
                    if response in ['y', 'yes', '']:
                        success = processor.create_outlook_draft()
                        result['email_created'] = success
                    else:
                        print("  Skipped email creation")
                        result['email_created'] = False
            
            results.append(result)
            successful += 1
            
        except Exception as e:
            print(f"  [ERROR] Failed to process {pdf_file.name}: {e}")
            results.append({
                'file': pdf_file.name,
                'success': False,
                'error': str(e)
            })
            failed += 1
    
    # Print summary report
    print("\n" + "="*80)
    print("BATCH PROCESSING SUMMARY")
    print("="*80)
    print(f"\nTotal Files Processed: {len(pdf_files)}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    
    if successful > 0:
        print(f"\n{'File':<40} {'VIN':<20} {'Vehicle':<30}")
        print("-" * 90)
        for result in results:
            if result['success']:
                file_name = result['file'][:38] + '..' if len(result['file']) > 40 else result['file']
                vin = result['data'].get('vin', 'N/A')[:18] + '..' if len(result['data'].get('vin', 'N/A')) > 20 else result['data'].get('vin', 'N/A')
                vehicle = f"{result['data'].get('year', '')} {result['data'].get('make', '')} {result['data'].get('model', '')}"
                vehicle = vehicle[:28] + '..' if len(vehicle) > 30 else vehicle
                print(f"{file_name:<40} {vin:<20} {vehicle:<30}")
    
    if failed > 0:
        print(f"\nFailed Files:")
        for result in results:
            if not result['success']:
                print(f"  - {result['file']}: {result.get('error', 'Unknown error')}")
    
    # Save detailed report to file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = f"batch_report_{timestamp}.txt"
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("HERTZ BATCH PROCESSING REPORT\n")
        f.write("="*80 + "\n")
        f.write(f"\nTimestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Total Files: {len(pdf_files)}\n")
        f.write(f"Successful: {successful}\n")
        f.write(f"Failed: {failed}\n")
        f.write("\n" + "="*80 + "\n")
        f.write("DETAILED RESULTS\n")
        f.write("="*80 + "\n")
        
        for idx, result in enumerate(results, 1):
            f.write(f"\n[{idx}] {result['file']}\n")
            f.write("-" * 80 + "\n")
            
            if result['success']:
                data = result['data']
                f.write(f"Status: SUCCESS\n")
                f.write(f"VIN: {data.get('vin', 'N/A')}\n")
                f.write(f"Vehicle: {data.get('year', 'N/A')} {data.get('make', 'N/A')} {data.get('model', 'N/A')}\n")
                f.write(f"Location: {data.get('full_location', 'N/A')}\n")
                f.write(f"Contact Phone: {data.get('contact_phone', 'N/A')}\n")
                f.write(f"Order Number: {data.get('order_number', 'N/A')}\n")
                f.write(f"Hertz Rep Emails:\n")
                for email in data.get('hertz_rep_emails', []):
                    f.write(f"  - {email}\n")
                if 'email_created' in result:
                    f.write(f"Email Draft Created: {'Yes' if result['email_created'] else 'No'}\n")
            else:
                f.write(f"Status: FAILED\n")
                f.write(f"Error: {result.get('error', 'Unknown error')}\n")
    
    print(f"\n[OK] Detailed report saved to: {report_file}")
    print("\n" + "="*80)


def main():
    """Main entry point for batch processing."""
    # Parse command-line arguments
    auto_mode = '--auto' in sys.argv or '-a' in sys.argv
    no_email = '--no-email' in sys.argv
    
    # Check for directory argument
    directory = "."
    for arg in sys.argv[1:]:
        if not arg.startswith('-') and Path(arg).is_dir():
            directory = arg
            break
    
    print(f"\nSearching directory: {directory}")
    
    # Run batch processing
    batch_process(
        directory=directory,
        auto_create=auto_mode,
        create_emails=not no_email
    )


if __name__ == "__main__":
    if '--help' in sys.argv or '-h' in sys.argv:
        print("""
Hertz Batch PDF Processor
========================

Usage: python batch_process_pdfs.py [DIRECTORY] [OPTIONS]

Arguments:
  DIRECTORY         Directory containing PDFs (default: current directory)

Options:
  --auto, -a        Automatically create all email drafts without confirmation
  --no-email        Extract data only, don't create email drafts
  --help, -h        Show this help message

Examples:
  python batch_process_pdfs.py                  # Process all PDFs in current directory
  python batch_process_pdfs.py --auto           # Auto-create all drafts
  python batch_process_pdfs.py C:\\PDFs --auto  # Process PDFs in specific directory
  python batch_process_pdfs.py --no-email       # Extract data only, no emails

Output:
  - Console summary of all processed files
  - Outlook email drafts (unless --no-email is used)
  - Detailed report file: batch_report_YYYYMMDD_HHMMSS.txt
""")
        sys.exit(0)
    
    main()




