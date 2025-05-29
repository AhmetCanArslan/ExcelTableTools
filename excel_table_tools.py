#!/usr/bin/env python3
"""
Excel Table Tools - Main Launcher Script
This script launches the Excel Table Tools application from the project root directory.
"""

import os
import sys

# Add project paths to Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, 'src')
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Import the main module and run the application
if __name__ == "__main__":
    try:
        from src.main import ExcelEditorApp
        import tkinter as tk
        
        root = tk.Tk()
        app = ExcelEditorApp(root)
        
        # Add proper close handling to prevent multiple window issues
        def on_closing():
            try:
                # Cancel any ongoing operations
                if hasattr(app, 'operation_manager'):
                    app.operation_manager.cancel_processing()
                # Properly exit the application
                root.quit()
                root.destroy()
            except Exception:
                pass
            finally:
                # Ensure the program exits completely
                import sys
                sys.exit(0)
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
    except ImportError as e:
        print(f"Error importing required modules: {e}")
        print(f"Current Python path: {sys.path}")
        input("Press Enter to exit...")
        sys.exit(1)