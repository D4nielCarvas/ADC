import sys
import os
import tkinter as tk
from gui.main_window import MainWindow

# Configuração para High DPI
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

if __name__ == "__main__":
    # Garante que imports funcionem
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    
    root = tk.Tk()
    
    # ícone (se existir)
    # try:
    #     root.iconbitmap("assets/icon.ico")
    # except: pass
    
    app = MainWindow(root)
    root.mainloop()
