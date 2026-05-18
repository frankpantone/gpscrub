"""
Find all CarMax GP folders in all accounts
"""

import win32com.client


def find_all_carmax_folders():
    """Find all folders containing 'CarMax' in all accounts."""
    print("="*80)
    print("FINDING ALL CARMAX FOLDERS")
    print("="*80)
    
    try:
        print("\n[INFO] Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("[OK] Connected successfully!")
        
        carmax_folders = []
        
        print("\n[INFO] Searching all accounts for CarMax folders...")
        print("-" * 80)
        
        # Search through all stores
        for store in namespace.Stores:
            store_name = store.DisplayName
            print(f"\n[SEARCHING] {store_name}")
            
            try:
                root_folder = store.GetRootFolder()
                found_folders = search_for_carmax_recursive(root_folder, store_name)
                carmax_folders.extend(found_folders)
            except Exception as e:
                print(f"  [SKIP] Cannot access store: {e}")
                continue
        
        print("\n" + "="*80)
        print("ALL CARMAX FOLDERS FOUND")
        print("="*80)
        
        if not carmax_folders:
            print("\n[WARNING] No CarMax folders found!")
        else:
            for idx, (path, folder, store) in enumerate(carmax_folders, 1):
                print(f"\n[{idx}] {folder.Name}")
                print(f"    Account: {store}")
                print(f"    Path: {path}")
                
                # Try to count emails in this folder
                try:
                    email_count = folder.Items.Count
                    print(f"    Emails in folder: {email_count}")
                    
                    # Check if this is the main (not archive) account
                    if "archive" not in store.lower():
                        print(f"    [***] This is in the MAIN account (not archive)!")
                except:
                    print(f"    Emails: Unable to count")
        
        print("\n" + "="*80)
        print("RECOMMENDATION")
        print("="*80)
        
        # Find the best folder (non-archive with emails)
        best_folder = None
        for path, folder, store in carmax_folders:
            if "archive" not in store.lower():
                try:
                    if folder.Items.Count > 0:
                        best_folder = (path, folder, store)
                        break
                except:
                    pass
        
        if best_folder:
            path, folder, store = best_folder
            print(f"\n[RECOMMENDED] Use this folder:")
            print(f"  Folder Name: {folder.Name}")
            print(f"  Account: {store}")
            print(f"  Full Path: {path}")
            print(f"\n  In email_config.py, set:")
            print(f'  FOLDER_NAME = "{folder.Name}"')
        else:
            print("\n[INFO] Check which folder has today's emails and update email_config.py")
        
        return carmax_folders
        
    except Exception as e:
        print(f"\n[ERROR] {e}")
        return []


def search_for_carmax_recursive(folder, store_name, parent_path=""):
    """Recursively search for folders containing 'carmax'."""
    found = []
    
    try:
        folder_name = folder.Name
        current_path = f"{parent_path}/{folder_name}" if parent_path else folder_name
        
        # Check if this folder has "carmax" in the name
        if "carmax" in folder_name.lower():
            print(f"  [FOUND] {current_path}")
            found.append((current_path, folder, store_name))
        
        # Search subfolders
        if folder.Folders.Count > 0:
            for subfolder in folder.Folders:
                found.extend(search_for_carmax_recursive(subfolder, store_name, current_path))
    
    except Exception as e:
        # Some folders may not be accessible
        pass
    
    return found


if __name__ == "__main__":
    find_all_carmax_folders()
    input("\nPress Enter to exit...")

















