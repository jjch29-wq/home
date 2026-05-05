import sys
import os
import unittest
from unittest.mock import MagicMock

# Add current directory to path
sys.path.append(os.getcwd())

def test_load_data():
    print("Testing MaterialManager.load_data()...")
    
    # Mock Tkinter
    mock_tk = MagicMock()
    sys.modules['tkinter'] = mock_tk
    sys.modules['tkinter.ttk'] = MagicMock()
    sys.modules['tkinter.messagebox'] = MagicMock()
    sys.modules['tkinter.filedialog'] = MagicMock()
    sys.modules['tkinter.simpledialog'] = MagicMock()
    sys.modules['tkcalendar'] = MagicMock()
    
    # Import the class
    from MaterialManager_10 import MaterialManager
    
    # Mock root
    root = MagicMock()
    
    try:
        # We need to bypass __init__ because it creates widgets
        # Or just mock create_widgets and other UI methods
        MaterialManager.create_widgets = MagicMock()
        MaterialManager.update_registration_combos = MagicMock()
        MaterialManager.setup_keyboard_shortcuts = MagicMock()
        MaterialManager.load_tab_config = MagicMock()
        MaterialManager.apply_autocomplete_to_all_comboboxes = MagicMock()
        
        app = MaterialManager(root)
        print("Successfully instantiated MaterialManager and called load_data().")
        
        # Check if dataframes are loaded and columns are normalized
        print(f"Materials DF columns: {list(app.materials_df.columns)}")
        print(f"Transactions DF columns: {list(app.transactions_df.columns)}")
        
        if 'MaterialID' in app.transactions_df.columns:
            print("SUCCESS: 'MaterialID' found in Transactions (normalized from 'Material ID').")
        else:
            print("FAILURE: 'MaterialID' NOT found in Transactions.")
            
    except Exception as e:
        print(f"Verification FAILED: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # Ensure we use the correct filename (MaterialManager-10.py -> MaterialManager_10)
    # We'll create a temporary copy if needed or just rename for import
    import shutil
    if not os.path.exists("MaterialManager_10.py"):
        shutil.copy2("MaterialManager-10.py", "MaterialManager_10.py")
    
    test_load_data()
    
    # Cleanup temporary file
    if os.path.exists("MaterialManager_10.py"):
        os.remove("MaterialManager_10.py")
