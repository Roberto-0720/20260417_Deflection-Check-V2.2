"""
Deflection Check Tool - SAP2000
================================
Check beam deflection (L/N) connecting directly to SAP2000 via COM API.
Supports physical members (groups of broken elements) and virtual mesh nodes.

Usage: python main.py
Requirements: Python 3.8+, Windows, SAP2000 running, comtypes, openpyxl
"""
import sys, os
import tkinter as tk
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ui.main_window import MainWindow

def main():
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
