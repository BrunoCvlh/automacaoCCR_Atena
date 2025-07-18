import time
import tkinter as tk
from tkinter import messagebox
import threading
from selenium_automation import SeleniumHandler
import calendar

class AtenaCommanderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Coordenação de Controladoria - GCO")
        self.root.geometry("550x350")

        self.root.grid_rowconfigure(0, weight=1)
        for i in range(1, 12):
            self.root.grid_rowconfigure(i, weight=1)
        self.root.grid_rowconfigure(12, weight=2)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(3, weight=0)
        self.root.grid_columnconfigure(4, weight=1)

        self.selenium_handler = SeleniumHandler() # Instancia o SeleniumHandler

        self.label = tk.Label(root, text="Baixar Análise Comparativa", font=("Arial", 16, "bold"), fg="#333")
        self.label.grid(row=0, column=1, columnspan=4, pady=15, sticky="nsew")

        self.email_label = tk.Label(root, text="Email:", font=("Arial", 12))
        self.email_label.grid(row=1, column=1, sticky="w", padx=20, pady=5)
        self.email_entry = tk.Entry(root, font=("Arial", 12), width=30, bd=2, relief="groove")
        self.email_entry.grid(row=1, column=2, columnspan=3, sticky="ew", padx=20, pady=5)

        self.password_label = tk.Label(root, text="Senha:", font=("Arial", 12))
        self.password_label.grid(row=2, column=1, sticky="w", padx=20, pady=5)
        self.password_entry = tk.Entry(root, font=("Arial", 12), show="*", width=30, bd=2, relief="groove")
        self.password_entry.grid(row=2, column=2, columnspan=3, sticky="ew", padx=20, pady=5)

        self.competencia_label = tk.Label(root, text="Competência (MM/YYYY):", font=("Arial", 12))
        self.competencia_label.grid(row=3, column=1, sticky="w", padx=20, pady=5)

        self.competencia_mes_entry = tk.Entry(root, font=("Arial", 12), width=5, bd=2, relief="groove")
        self.competencia_mes_entry.grid(row=3, column=2, sticky="e", padx=(20, 5), pady=5)

        self.separator_label = tk.Label(root, text="/", font=("Arial", 12, "bold"))
        self.separator_label.grid(row=3, column=3, sticky="nsew", pady=5)

        self.competencia_ano_entry = tk.Entry(root, font=("Arial", 12), width=7, bd=2, relief="groove")
        self.competencia_ano_entry.grid(row=3, column=4, sticky="w", padx=(5, 20), pady=5)

        self.extract_report_button = tk.Button(root, text="Extrair Relatório", command=self.start_login_thread,
                                               font=("Arial", 12, "bold"), bg="#305374", fg="white",
                                               activebackground="#305374", activeforeground="white",
                                               relief="raised", bd=3, cursor="hand2")
        self.extract_report_button.grid(row=5, column=1, columnspan=4, sticky="nsew", padx=20, pady=10)

        self.loading_label = tk.Label(root,
                                       text="Aguarde enquanto o relatório é acessado...\nNão feche o programa até o download ser concluído.",
                                       font=("Arial", 10, "italic"), fg="blue", wraplength=400)
        self.loading_label.grid(row=7, column=1, columnspan=4, sticky="nsew", padx=20, pady=5)
        self.loading_label.grid_forget()
        
        self.status_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
        self.status_label.grid(row=8, column=1, columnspan=4, sticky="nsew", padx=20, pady=5)

    def start_login_thread(self):
        self.status_label.config(text="")
        self.loading_label.grid(row=7, column=1, columnspan=4, sticky="nsew", padx=20, pady=5)
        
        self.extract_report_button.config(state=tk.DISABLED)

        login_thread = threading.Thread(target=self._login_in_thread)
        login_thread.start()

    def _login_in_thread(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        mes = self.competencia_mes_entry.get()
        ano = self.competencia_ano_entry.get()

        if not all([email, password, mes, ano]):
            messagebox.showerror("Erro de Validação", "Todos os campos devem ser preenchidos.")
            self.loading_label.grid_forget()
            self.extract_report_button.config(state=tk.NORMAL)
            return

        try:
            success = self.selenium_handler._login(email, password, mes, ano)
            
            self.loading_label.grid_forget()

            if success:
                self.status_label.config(text="Download da análise comparativa concluído!", fg="green")
            else:
                self.status_label.config(text="Erro durante o login ou navegação.", fg="red")

        except Exception as e:
            self.loading_label.grid_forget()
            self.status_label.config(text=f"Ocorreu um erro inesperado: {e}", fg="red")
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")
        finally:
            self.selenium_handler.quit_driver()
            self.extract_report_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = AtenaCommanderApp(root)
    root.mainloop()