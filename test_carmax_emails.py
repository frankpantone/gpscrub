"""
Quick test to check for emails in CarMax GP folders
"""

import win32com.client
from datetime import datetime


def test_carmax_emails():
    """Test if there are emails in CarMax GP folders."""
    print("="*80)
    print("TESTING CARMAX GP FOLDERS FOR EMAILS")
    print("="*80)
    
    try:
        print("\n[INFO] Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("[OK] Connected!")
        
        today = datetime.now().date()
        print(f"\n[INFO] Checking for emails from: {today.strftime('%Y-%m-%d')}")
        
        # Search all stores for CarMax GP
        for store in namespace.Stores:
            store_name = store.DisplayName
            
            # Skip archives for now, focus on main accounts
            if "archive" in store_name.lower():
                continue
                
            print(f"\n[CHECKING] Account: {store_name}")
            
            try:
                root_folder = store.GetRootFolder()
                check_folder_recursive(root_folder, today)
            except Exception as e:
                print(f"  [SKIP] Cannot access: {e}")
        
        # Also check the main namespace folders
        print(f"\n[CHECKING] Main Outlook Folders")
        for folder in namespace.Folders:
            folder_name = folder.Name
            print(f"\n[FOLDER] {folder_name}")
            check_folder_recursive(folder, today)
        
        print("\n" + "="*80)
        print("[OK] Test complete!")
        print("="*80)
        
    except Exception as e:
        print(f"\n[ERROR] {e}")


def check_folder_recursive(folder, target_date, level=0):
    """Recursively check folders for CarMax GP and count emails."""
    try:
        folder_name = folder.Name
        indent = "  " * level
        
        # Check if this is a CarMax folder
        if "carmax" in folder_name.lower():
            print(f"{indent}[FOUND] CarMax folder: {folder_name}")
            
            try:
                items = folder.Items
                total_count = items.Count
                print(f"{indent}  Total emails in folder: {total_count}")
                
                # Count today's emails
                today_count = 0
                keywords_found = []
                
                # Check recent emails (last 20)
                items.Sort("[ReceivedTime]", True)
                
                for i, item in enumerate(items):
                    if i >= 20:  # Only check first 20
                        break
                    
                    try:
                        if not hasattr(item, 'ReceivedTime'):
                            continue
                        
                        received_date = item.ReceivedTime.date()
                        subject = item.Subject if hasattr(item, 'Subject') else ""
                        
                        if received_date == target_date:
                            today_count += 1
                            print(f"{indent}    [{today_count}] {subject[:70]}")
                            
                            # Check for keywords
                            if "KUNES" in subject.upper():
                                keywords_found.append("KUNES")
                            if "EASTON" in subject.upper():
                                keywords_found.append("EASTON")
                    except:
                        continue
                
                print(f"{indent}  Emails from today: {today_count}")
                if keywords_found:
                    print(f"{indent}  Keywords found: {', '.join(set(keywords_found))}")
                
                if today_count > 0:
                    print(f"{indent}  [***] THIS FOLDER HAS TODAY'S EMAILS!")
                    print(f"{indent}  Use this in email_config.py:")
                    print(f"{indent}  FOLDER_NAME = \"{folder_name}\"")
                
            except Exception as e:
                print(f"{indent}  [ERROR] Cannot read emails: {e}")
        
        # Check subfolders
        if folder.Folders.Count > 0:
            for subfolder in folder.Folders:
                check_folder_recursive(subfolder, target_date, level + 1)
    
    except Exception as e:
        pass


if __name__ == "__main__":
    test_carmax_emails()
    input("\nPress Enter to exit...")

















