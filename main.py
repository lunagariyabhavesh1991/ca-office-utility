import sys
import os

# Ensure the parent directory is in the path to allow absolute imports within the project
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from ui.app_window import AppWindow
import tkinter.colorchooser
import tkinter.filedialog
import tkinter.messagebox
import tkinter.simpledialog
import win32crypt
import win32api
import pywintypes
import pythoncom

def main():
    """
    Main entry point for the CA Office PDF Utility application.
    Initializes the main window and starts the GUI event loop.
    """
    try:
        # Initialize COM for win32crypt in frozen context
        pythoncom.CoInitialize()
        
        app = AppWindow()
        app.mainloop()
    except Exception as e:
        import traceback
        print(f"An error occurred: {e}")
        traceback.print_exc()
        input("Press Enter to exit...")
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == "__main__":
    main()
