import time
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk  # Adicione esta linha
import threading
from selenium_automation import SeleniumHandler
import calendar

class AtenaCommanderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Coordenação de Controladoria - GCO")
        self.root.geometry("750x800")

        self.root.grid_rowconfigure(0, weight=1)
        for i in range(1, 15):
            self.root.grid_rowconfigure(i, weight=1)
        self.root.grid_rowconfigure(15, weight=2)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(3, weight=0)
        self.root.grid_columnconfigure(4, weight=1)

        self.selenium_handler = SeleniumHandler()

        self.label = tk.Label(root, text="Baixar Análise Comparativa", font=("Arial", 16, "bold"), fg="#333")
        self.label.grid(row=0, column=1, columnspan=4, pady=15, sticky="nsew")

        # Título da primeira parte
        self.dados_section_label = tk.Label(root, text="1ª Parte: Dados de Acesso e Competência", font=("Arial", 13, "bold"), fg="#305374")
        self.dados_section_label.grid(row=1, column=1, columnspan=4, sticky="w", padx=20, pady=(0, 5))

        self.email_label = tk.Label(root, text="Email:", font=("Arial", 12))
        self.email_label.grid(row=2, column=1, sticky="w", padx=20, pady=5)
        self.email_entry = tk.Entry(root, font=("Arial", 12), width=30, bd=2, relief="groove")
        self.email_entry.grid(row=2, column=2, columnspan=3, sticky="ew", padx=20, pady=5)

        self.password_label = tk.Label(root, text="Senha:", font=("Arial", 12))
        self.password_label.grid(row=3, column=1, sticky="w", padx=20, pady=5)
        self.password_entry = tk.Entry(root, font=("Arial", 12), show="*", width=30, bd=2, relief="groove")
        self.password_entry.grid(row=3, column=2, columnspan=3, sticky="ew", padx=20, pady=5)

        self.competencia_label = tk.Label(root, text="Competência (MM/YYYY):", font=("Arial", 12))
        self.competencia_label.grid(row=4, column=1, sticky="w", padx=20, pady=5)

        # Campo para mês (2 dígitos)
        self.competencia_mes_entry = tk.Entry(root, font=("Arial", 12), width=3, bd=2, relief="groove", justify="center")
        self.competencia_mes_entry.grid(row=4, column=2, sticky="e", padx=(20, 2), pady=5)
        self.competencia_mes_entry.config(validate="key", validatecommand=(self.root.register(lambda P: len(P) <= 2 and (P.isdigit() or P == "")), "%P"))

        # Separador "/"
        self.competencia_separator = tk.Label(root, text="/", font=("Arial", 12, "bold"))
        self.competencia_separator.grid(row=4, column=3, sticky="w", pady=5)

        # Campo para ano (4 dígitos)
        self.competencia_ano_entry = tk.Entry(root, font=("Arial", 12), width=5, bd=2, relief="groove", justify="center")
        self.competencia_ano_entry.grid(row=4, column=4, sticky="w", padx=(2, 20), pady=5)
        self.competencia_ano_entry.config(validate="key", validatecommand=(self.root.register(lambda P: len(P) <= 4 and (P.isdigit() or P == "")), "%P"))

        # Linha separadora visual
        self.separator = ttk.Separator(root, orient='horizontal')
        self.separator.grid(row=5, column=1, columnspan=4, sticky="ew", padx=20, pady=(10, 0))  # Espaço só abaixo

        # Título da segunda parte
        self.tratamento_section_label = tk.Label(root, text="2ª Parte: Tratamento do Arquivo", font=("Arial", 13, "bold"), fg="#305374")
        self.tratamento_section_label.grid(row=6, column=1, columnspan=4, sticky="w", padx=20, pady=(20, 5))  # Espaço acima

        self.tratamento_file_label = tk.Label(root, text="Anexar Arquivo para Tratamento:", font=("Arial", 12))
        self.tratamento_file_label.grid(row=7, column=1, sticky="w", padx=20, pady=5)
        self.tratamento_file_path_label = tk.Label(root, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="gray")
        self.tratamento_file_path_label.grid(row=7, column=2, columnspan=2, sticky="ew", padx=20, pady=5)
        self.tratamento_file_button = tk.Button(root, text="Selecionar Arquivo", command=self.select_tratamento_file,
                                                font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2)
        self.tratamento_file_button.grid(row=7, column=4, sticky="ew", padx=20, pady=5)
        self.tratamento_file_path = ""

        self.tratamento_button = tk.Button(root, text="Tratar Arquivo", command=self.tratar_arquivo,
                                           font=("Arial", 12, "bold"), bg="#305374", fg="white", relief="raised", bd=2)
        self.tratamento_button.grid(row=8, column=1, columnspan=4, sticky="ew", padx=20, pady=5)

        # Linha separadora visual para terceira parte
        self.separator2 = ttk.Separator(root, orient='horizontal')
        self.separator2.grid(row=9, column=1, columnspan=4, sticky="ew", padx=20, pady=(10, 0))  # Espaço só abaixo

        # Título da terceira parte
        self.planilha_section_label = tk.Label(root, text="3ª Parte: Anexar Planilha e Alterar Valor Contingencial", font=("Arial", 13, "bold"), fg="#305374")
        self.planilha_section_label.grid(row=10, column=1, columnspan=4, sticky="w", padx=20, pady=(20, 5))  # Espaço acima

        self.planilha_label = tk.Label(root, text="Anexar Planilha:", font=("Arial", 12))
        self.planilha_label.grid(row=11, column=1, sticky="w", padx=20, pady=5)
        self.planilha_path_label = tk.Label(root, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="gray")
        self.planilha_path_label.grid(row=11, column=2, columnspan=2, sticky="ew", padx=20, pady=5)
        self.planilha_button = tk.Button(root, text="Selecionar Arquivo", command=self.select_planilha,
                                         font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2)
        self.planilha_button.grid(row=11, column=4, sticky="ew", padx=20, pady=5)
        self.planilha_path = ""

        self.contingencia_label = tk.Label(root, text="Valor Contingencial:", font=("Arial", 12))
        self.contingencia_label.grid(row=12, column=1, sticky="w", padx=20, pady=5)
        self.contingencia_entry = tk.Entry(root, font=("Arial", 12), width=30, bd=2, relief="groove")
        self.contingencia_entry.grid(row=12, column=2, columnspan=3, sticky="ew", padx=20, pady=5)

        self.extract_report_button = tk.Button(root, text="Extrair Relatório", command=self.start_login_thread,
                                               font=("Arial", 12, "bold"), bg="#305374", fg="white",
                                               activebackground="#305374", activeforeground="white",
                                               relief="raised", bd=3, cursor="hand2")
        self.extract_report_button.grid(row=13, column=1, columnspan=4, sticky="nsew", padx=20, pady=10)

        self.loading_label = tk.Label(root,
                                       text="Aguarde enquanto o relatório é acessado...\nNão feche o programa até o download ser concluído.",
                                       font=("Arial", 10, "italic"), fg="blue", wraplength=400)
        self.loading_label.grid(row=14, column=1, columnspan=4, sticky="nsew", padx=20, pady=5) 
        self.loading_label.grid_forget()
        
        self.status_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
        self.status_label.grid(row=15, column=1, columnspan=4, sticky="nsew", padx=20, pady=5)

        # Linha separadora visual para quarta parte
        self.separator3 = ttk.Separator(root, orient='horizontal')
        self.separator3.grid(row=16, column=1, columnspan=4, sticky="ew", padx=20, pady=(0, 0))  # Espaço só abaixo

        # Título da quarta parte
        self.envio_section_label = tk.Label(root, text="4ª Parte: Enviar Dados para a Base", font=("Arial", 13, "bold"), fg="#305374")
        self.envio_section_label.grid(row=17, column=1, columnspan=4, sticky="w", padx=20, pady=(20, 5))  # Espaço acima

        # Campo para selecionar caminho da base
        self.base_path_label = tk.Label(root, text="Selecionar caminho da base:", font=("Arial", 12))
        self.base_path_label.grid(row=18, column=1, sticky="w", padx=20, pady=5)
        self.base_path_display = tk.Label(root, text="Nenhum caminho selecionado", font=("Arial", 10), fg="gray")
        self.base_path_display.grid(row=18, column=2, columnspan=2, sticky="ew", padx=20, pady=5)
        self.base_path_button = tk.Button(root, text="Procurar", command=self.select_base_path,
                                          font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2)
        self.base_path_button.grid(row=18, column=4, sticky="ew", padx=20, pady=5)
        self.base_path = ""

        # Botão de envio
        self.envio_button = tk.Button(root, text="Enviar dados para a base", command=self.enviar_dados_base,
                                      font=("Arial", 12, "bold"), bg="#305374", fg="white", relief="raised", bd=2)
        self.envio_button.grid(row=19, column=1, columnspan=4, sticky="ew", padx=20, pady=10)

    def select_planilha(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.planilha_path = file_path
            self.planilha_path_label.config(text=f"Arquivo selecionado: {file_path.split('/')[-1]}")
        else:
            self.planilha_path = ""
            self.planilha_path_label.config(text="Nenhum arquivo selecionado")

    def select_tratamento_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Todos os arquivos", "*")]
        )
        if file_path:
            self.tratamento_file_path = file_path
            self.tratamento_file_path_label.config(text=f"Arquivo selecionado: {file_path.split('/')[-1]}")
        else:
            self.tratamento_file_path = ""
            self.tratamento_file_path_label.config(text="Nenhum arquivo selecionado")

    def start_login_thread(self):
        self.status_label.config(text="")
        self.loading_label.grid(row=9, column=1, columnspan=4, sticky="nsew", padx=20, pady=5)
        
        self.extract_report_button.config(state=tk.DISABLED)

        login_thread = threading.Thread(target=self._login_in_thread)
        login_thread.start()

    def _login_in_thread(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        mes = self.competencia_mes_entry.get()
        ano = self.competencia_ano_entry.get()
        competencia = f"{mes}/{ano}"
        valor_contingencial = self.contingencia_entry.get()
        planilha_anexada = self.planilha_path

        if not all([email, password, mes, ano, valor_contingencial, planilha_anexada]):
            messagebox.showerror("Erro de Validação", "Todos os campos devem ser preenchidos, incluindo o anexo da planilha.")
            self.loading_label.grid_forget()
            self.extract_report_button.config(state=tk.NORMAL)
            return

        try:
            # Ajuste para passar a competência como MM/YYYY
            success = self.selenium_handler._login(email, password, competencia)
            
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

    def tratar_arquivo(self):
        if not self.tratamento_file_path:
            messagebox.showerror("Erro", "Selecione um arquivo para tratamento.")
            return
        # Adicione aqui a lógica de tratamento do arquivo conforme necessário
        messagebox.showinfo("Tratamento", f"Tratamento do arquivo realizado: {self.tratamento_file_path}")

    def select_base_path(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo da base",
            filetypes=[("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.base_path = file_path
            self.base_path_display.config(text=f"Caminho selecionado: {file_path}")
        else:
            self.base_path = ""
            self.base_path_display.config(text="Nenhum caminho selecionado")

    def enviar_dados_base(self):
        if not self.base_path:
            messagebox.showerror("Erro", "Selecione o caminho da base antes de enviar os dados.")
            return
        # Adicione aqui a lógica de envio dos dados para a base
        messagebox.showinfo("Envio", f"Dados enviados para a base: {self.base_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AtenaCommanderApp(root)
    root.mainloop()