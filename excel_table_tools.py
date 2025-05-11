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

# Import the main module and run the application
if __name__ == "__main__":
    from src.main import ExcelEditorApp
    import tkinter as tk
    
    root = tk.Tk()
    app = ExcelEditorApp(root)
    root.mainloop()