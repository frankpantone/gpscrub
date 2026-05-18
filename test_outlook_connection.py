"""
Test Outlook Connection and List Folders
Use this to verify your Outlook setup and find folder names.
"""

import win32com.client


def test_outlook_connection():
    """Test connection to Outlook and list available folders."""
    print("="*80)
    print("OUTLOOK CONNECTION TEST")
    print("="*80)
    
    try:
        print("\n[1/3] Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("[OK] Connected successfully!")
        
        print("\n[2/3] Listing email accounts...")
        print("-" * 80)
        
        accounts_found = []
        for store in namespace.Stores:
            print(f"\n[ACCOUNT] {store.DisplayName}")
            accounts_found.append(store.DisplayName)
            
            # Check if this might be the hertzlogistics account
            if "hertz" in store.DisplayName.lower() or "logistics" in store.DisplayName.lower():
                print("   [***] This looks like the Hertz Logistics account!")
        
        if not accounts_found:
            print("[WARNING] No accounts found")
            return False
        
        print(f"\n[OK] Found {len(accounts_found)} account(s)")
        
        print("\n[3/3] Listing folders...")
        print("-" * 80)
        
        all_folders = []
        
        for folder in namespace.Folders:
            print(f"\n[FOLDER] {folder.Name}")
            all_folders.append(folder.Name)
            
            # List subfolders
            try:
                if folder.Folders.Count > 0:
                    for subfolder in folder.Folders:
                        print(f"   +-- {subfolder.Name}")
                        all_folders.append(f"{folder.Name}/{subfolder.Name}")
                        
                        # List sub-subfolders (one level deep)
                        try:
                            if subfolder.Folders.Count > 0:
                                for subsubfolder in subfolder.Folders:
                                    print(f"       +-- {subsubfolder.Name}")
                                    all_folders.append(f"{folder.Name}/{subfolder.Name}/{subsubfolder.Name}")
                                    
                                    # Check if this is the CarMax GP folder
                                    if "carmax" in subsubfolder.Name.lower() and "gp" in subsubfolder.Name.lower():
                                        print(f"           [***] This might be your CarMax GP folder!")
                                    elif "carmax" in subsubfolder.Name.lower():
                                        print(f"           [***] CarMax folder found!")
                        except:
                            pass
                        
                        # Check if this is the CarMax GP folder
                        if "carmax" in subfolder.Name.lower() and "gp" in subfolder.Name.lower():
                            print(f"       [***] This might be your CarMax GP folder!")
                        elif "carmax" in subfolder.Name.lower():
                            print(f"       [***] CarMax folder found!")
            except:
                pass
        
        print("\n" + "="*80)
        print("SUMMARY")
        print("="*80)
        print(f"\n[OK] Outlook Connection: Working")
        print(f"[OK] Email Accounts: {len(accounts_found)}")
        print(f"[OK] Folders Found: {len(all_folders)}")
        
        print("\n" + "="*80)
        print("NEXT STEPS")
        print("="*80)
        print("\n1. Look for your CarMax GP folder in the list above (marked with [***])")
        print("2. Copy the exact folder name")
        print("3. Edit email_config.py and set FOLDER_NAME to that exact name")
        print("\nExample:")
        print('   FOLDER_NAME = "CarMax GP"')
        print('   or')
        print('   FOLDER_NAME = "Inbox/CarMax GP"')
        print("\n4. Also note the email account name (look for hertzlogistics)")
        print("5. Edit email_config.py and set EMAIL_ACCOUNT accordingly")
        
        print("\n" + "="*80)
        print("[OK] TEST COMPLETE")
        print("="*80)
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Error: {e}")
        print("\nTroubleshooting:")
        print("  - Make sure Microsoft Outlook is installed")
        print("  - Make sure Outlook is open and logged in")
        print("  - Try opening Outlook manually first")
        return False


if __name__ == "__main__":
    test_outlook_connection()
    input("\nPress Enter to exit...")

