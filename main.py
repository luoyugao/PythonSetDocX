import tkinter as tk
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import doc_parameters_manager as dpm


def main():
    dpm.load_settings()
    
    root = tk.Tk()
    
    from main_form import MainForm
    app = MainForm(root)
    
    root.mainloop()


if __name__ == "__main__":
    main()