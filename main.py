import tkinter as tk
from gui_app import AtenaCommanderApp
from inclui_dados_na_base import incluir_dados_na_base

if __name__ == "__main__":
    root = tk.Tk()
    app = AtenaCommanderApp(root)
    root.mainloop()