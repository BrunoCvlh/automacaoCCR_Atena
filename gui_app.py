import time
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import threading
from selenium_automation import SeleniumHandler
from tratamento_da_planilha import tratar_planilha
from ajuste_contingencial import ajustar_contingencial
from inclui_dados_na_base import incluir_dados_na_base
import calendar

class AtenaCommanderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Coordenação de Controladoria - GCO")
        self.root.geometry("850x900")

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=0)
        for i in range(1, 5):
            self.root.grid_rowconfigure(i, weight=1)
        self.root.grid_rowconfigure(5, weight=0)

        self.selenium_handler = SeleniumHandler()
        self.tratamento_file_path = ""
        self.planilha_path = ""
        self.base_path = ""
        self.tratada_path = ""

        self.main_title_label = tk.Label(root, text="Automação de Relatórios e Bases - Atena Commander", font=("Arial", 18, "bold"), fg="#1A2F4B")
        self.main_title_label.grid(row=0, column=0, columnspan=2, pady=20, sticky="nsew")

        self.access_frame = ttk.LabelFrame(root, text=" 1ª Passo: Acesso e Download de Relatório ", padding=(10, 10))
        self.access_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        self.access_frame.grid_columnconfigure(0, weight=0)
        self.access_frame.grid_columnconfigure(1, weight=1)
        self.access_frame.grid_columnconfigure(2, weight=0)
        self.access_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.access_frame, text="Email:", font=("Arial", 12)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.email_entry = tk.Entry(self.access_frame, font=("Arial", 12), bd=2, relief="groove")
        self.email_entry.grid(row=0, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.access_frame, text="Senha:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.password_entry = tk.Entry(self.access_frame, font=("Arial", 12), show="*", bd=2, relief="groove")
        self.password_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.access_frame, text="Competência (MM/YYYY):", font=("Arial", 12)).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.competencia_mes_entry = tk.Entry(self.access_frame, font=("Arial", 12), width=3, bd=2, relief="groove", justify="center")
        self.competencia_mes_entry.grid(row=2, column=1, sticky="e", padx=(5, 2), pady=5)
        self.competencia_mes_entry.config(validate="key", validatecommand=(self.root.register(lambda P: len(P) <= 2 and (P.isdigit() or P == "")), "%P"))

        tk.Label(self.access_frame, text="/", font=("Arial", 12, "bold")).grid(row=2, column=2, sticky="w", pady=5)

        self.competencia_ano_entry = tk.Entry(self.access_frame, font=("Arial", 12), width=5, bd=2, relief="groove", justify="center")
        self.competencia_ano_entry.grid(row=2, column=3, sticky="w", padx=(2, 5), pady=5)
        self.competencia_ano_entry.config(validate="key", validatecommand=(self.root.register(lambda P: len(P) <= 4 and (P.isdigit() or P == "")), "%P"))

        self.extract_report_button = tk.Button(self.access_frame, text="Extrair Relatório de Análise Comparativa", command=self.start_login_thread,
                                               font=("Arial", 10, "bold"), bg="#305374", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.extract_report_button.grid(row=3, column=0, columnspan=4, sticky="", padx=5, pady=10)

        self.treatment_frame = ttk.LabelFrame(root, text=" 2ª Passo: Tratamento do Arquivo ", padding=(10, 10))
        self.treatment_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        self.treatment_frame.grid_columnconfigure(0, weight=0)
        self.treatment_frame.grid_columnconfigure(1, weight=1)
        self.treatment_frame.grid_columnconfigure(2, weight=0)
        self.treatment_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.treatment_frame, text="Competência (MM/YYYY):", font=("Arial", 12)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.tratamento_competencia_mes_entry = tk.Entry(self.treatment_frame, font=("Arial", 12), width=3, bd=2, relief="groove", justify="center")
        self.tratamento_competencia_mes_entry.grid(row=0, column=1, sticky="e", padx=(5, 2), pady=5)
        self.tratamento_competencia_mes_entry.config(validate="key", validatecommand=(self.root.register(lambda P: len(P) <= 2 and (P.isdigit() or P == "")), "%P"))
        tk.Label(self.treatment_frame, text="/", font=("Arial", 12, "bold")).grid(row=0, column=2, sticky="w", pady=5)
        self.tratamento_competencia_ano_entry = tk.Entry(self.treatment_frame, font=("Arial", 12), width=5, bd=2, relief="groove", justify="center")
        self.tratamento_competencia_ano_entry.grid(row=0, column=3, sticky="w", padx=(2, 5), pady=5)
        self.tratamento_competencia_ano_entry.config(validate="key", validatecommand=(self.root.register(lambda P: len(P) <= 4 and (P.isdigit() or P == "")), "%P"))

        tk.Label(self.treatment_frame, text="Anexar Arquivo para Tratamento:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.tratamento_file_path_label = tk.Label(self.treatment_frame, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="gray")
        self.tratamento_file_path_label.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tratamento_file_button = tk.Button(self.treatment_frame, text="Selecionar Arquivo", command=self.select_tratamento_file,
                                                 font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.tratamento_file_button.grid(row=1, column=3, sticky="ew", padx=5, pady=5)

        self.tratamento_button = tk.Button(self.treatment_frame, text="Tratar Arquivo", command=self.tratar_arquivo,
                                           font=("Arial", 10, "bold"), bg="#305374", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.tratamento_button.grid(row=2, column=0, columnspan=4, sticky="", padx=5, pady=10)

        self.contingency_frame = ttk.LabelFrame(root, text=" 3ª Passo: Ajustar Valor Contingencial ", padding=(10, 10))
        self.contingency_frame.grid(row=3, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        self.contingency_frame.grid_columnconfigure(0, weight=0)
        self.contingency_frame.grid_columnconfigure(1, weight=1)
        self.contingency_frame.grid_columnconfigure(2, weight=1)
        self.contingency_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.contingency_frame, text="Anexar Planilha para Ajuste:", font=("Arial", 12)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.planilha_path_label = tk.Label(self.contingency_frame, text="Nenhum arquivo selecionado", font=("Arial", 10), fg="gray")
        self.planilha_path_label.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.planilha_button = tk.Button(self.contingency_frame, text="Selecionar Arquivo", command=self.select_planilha,
                                         font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.planilha_button.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.contingency_frame, text="Novo Valor Contingencial:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.contingencia_entry = tk.Entry(self.contingency_frame, font=("Arial", 12), bd=2, relief="groove")
        self.contingencia_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        self.adjust_contingency_button = tk.Button(self.contingency_frame, text="Ajustar Valor Contingencial na Planilha", command=self.ajustar_valor_contingencial,
                                                   font=("Arial", 10, "bold"), bg="#305374", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.adjust_contingency_button.grid(row=2, column=0, columnspan=4, sticky="", padx=5, pady=10)

        self.send_data_frame = ttk.LabelFrame(root, text=" 4ª Passo: Enviar Dados para a Base ", padding=(10, 10))
        self.send_data_frame.grid(row=4, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        self.send_data_frame.grid_columnconfigure(0, weight=0)
        self.send_data_frame.grid_columnconfigure(1, weight=1)
        self.send_data_frame.grid_columnconfigure(2, weight=1)
        self.send_data_frame.grid_columnconfigure(3, weight=1)

        tk.Label(self.send_data_frame, text="Caminho da Base de Dados:", font=("Arial", 12)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.base_path_display = tk.Label(self.send_data_frame, text="Nenhum caminho selecionado", font=("Arial", 10), fg="gray")
        self.base_path_display.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.base_path_button = tk.Button(self.send_data_frame, text="Procurar Base", command=self.select_base_path,
                                          font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.base_path_button.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        tk.Label(self.send_data_frame, text="Planilha Tratada para Envio:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.tratada_path_label = tk.Label(self.send_data_frame, text="Nenhuma planilha tratada selecionada", font=("Arial", 10), fg="gray")
        self.tratada_path_label.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.tratada_path_button = tk.Button(self.send_data_frame, text="Selecionar Planilha Tratada", command=self.select_tratada_file,
                                             font=("Arial", 10), bg="#D3D3D3", fg="black", relief="raised", bd=2, width=20)
        self.tratada_path_button.grid(row=1, column=3, sticky="ew", padx=5, pady=5)

        self.envio_button = tk.Button(self.send_data_frame, text="Enviar Dados para a Base", command=self.enviar_dados_base,
                                       font=("Arial", 10, "bold"), bg="#305374", fg="white", relief="raised", bd=2, cursor="hand2", width=40)
        self.envio_button.grid(row=2, column=0, columnspan=4, sticky="", padx=5, pady=10)

        self.loading_label = tk.Label(root, text="Aguarde enquanto a operação é processada...\nNão feche o programa até a conclusão.",
                                      font=("Arial", 10, "italic"), fg="blue", wraplength=700, justify="center")
        self.loading_label.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=20, pady=5)
        self.loading_label.grid_forget()

        self.status_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
        self.status_label.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=20, pady=5)

    def select_planilha(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.planilha_path = file_path
            self.planilha_path_label.config(text=f"Arquivo selecionado: {file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]}")
        else:
            self.planilha_path = ""
            self.planilha_path_label.config(text="Nenhum arquivo selecionado")

    def select_tratamento_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.tratamento_file_path = file_path
            self.tratamento_file_path_label.config(text=f"Arquivo selecionado: {file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]}")
        else:
            self.tratamento_file_path = ""
            self.tratamento_file_path_label.config(text="Nenhum arquivo selecionado")

    def select_tratada_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.tratada_path = file_path
            self.tratada_path_label.config(text=f"Planilha tratada: {file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]}")
        else:
            self.tratada_path = ""
            self.tratada_path_label.config(text="Nenhuma planilha tratada selecionada")

    def start_login_thread(self):
        self.status_label.config(text="")
        self.loading_label.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=20, pady=5)
        
        self.extract_report_button.config(state=tk.DISABLED)
        self.tratamento_button.config(state=tk.DISABLED)
        self.adjust_contingency_button.config(state=tk.DISABLED)
        self.envio_button.config(state=tk.DISABLED)

        login_thread = threading.Thread(target=self._login_in_thread)
        login_thread.start()

    def _login_in_thread(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        mes = self.competencia_mes_entry.get()
        ano = self.competencia_ano_entry.get()
        if not all([email, password, mes, ano]):
            messagebox.showerror("Erro de Validação", "Preencha o Email, Senha e Competência para realizar a extração.")
            self.loading_label.grid_forget()
            self.extract_report_button.config(state=tk.NORMAL)
            self.tratamento_button.config(state=tk.NORMAL)
            self.adjust_contingency_button.config(state=tk.NORMAL)
            self.envio_button.config(state=tk.NORMAL)
            return

        try:
            success = self.selenium_handler._login(email, password, mes, ano)
            self.loading_label.grid_forget()
            if success:
                self.status_label.config(text="Download da análise comparativa concluído!", fg="green")
            else:
                self.status_label.config(text="Erro durante o login ou navegação. Verifique as credenciais e a competência.", fg="red")
        except Exception as e:
            self.loading_label.grid_forget()
            self.status_label.config(text=f"Ocorreu um erro inesperado: {e}", fg="red")
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")
        finally:
            self.selenium_handler.quit_driver()
            self.extract_report_button.config(state=tk.NORMAL)
            self.tratamento_button.config(state=tk.NORMAL)
            self.adjust_contingency_button.config(state=tk.NORMAL)
            self.envio_button.config(state=tk.NORMAL)


    def tratar_arquivo(self):
        self.status_label.config(text="")
        if not self.tratamento_file_path:
            messagebox.showerror("Erro", "Selecione um arquivo para tratamento.")
            return
        
        mes = self.tratamento_competencia_mes_entry.get()
        ano = self.tratamento_competencia_ano_entry.get()
        
        if not all([mes, ano]):
            messagebox.showerror("Erro de Validação", "Preencha a competência (Mês/Ano) para o tratamento do arquivo.")
            return

        competencia = f"{mes}/{ano}"
        
        self.extract_report_button.config(state=tk.DISABLED)
        self.tratamento_button.config(state=tk.DISABLED)
        self.adjust_contingency_button.config(state=tk.DISABLED)
        self.envio_button.config(state=tk.DISABLED)
        self.loading_label.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=20, pady=5)

        try:
            tratado_path = tratar_planilha(self.tratamento_file_path, competencia)
            if tratado_path.startswith("ERRO:"):
                messagebox.showerror("Erro", f"Erro ao tratar arquivo:\n{tratado_path}")
                self.status_label.config(text="Erro no tratamento do arquivo.", fg="red")
            else:
                messagebox.showinfo("Tratamento", f"Tratamento do arquivo realizado!\nArquivo salvo em:\n{tratado_path}")
                self.status_label.config(text="Arquivo tratado com sucesso!", fg="green")
                self.tratada_path = tratado_path
                self.tratada_path_label.config(text=f"Planilha tratada: {tratado_path.split('/')[-1] if '/' in tratado_path else tratado_path.split('\\')[-1]}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado durante o tratamento: {e}")
            self.status_label.config(text=f"Erro inesperado no tratamento: {e}", fg="red")
        finally:
            self.loading_label.grid_forget()
            self.extract_report_button.config(state=tk.NORMAL)
            self.tratamento_button.config(state=tk.NORMAL)
            self.adjust_contingency_button.config(state=tk.NORMAL)
            self.envio_button.config(state=tk.NORMAL)


    def select_base_path(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo da base",
            filetypes=[("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.base_path = file_path
            nome_arquivo = file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]
            self.base_path_display.config(text=f"Caminho selecionado: {nome_arquivo}")
        else:
            self.base_path = ""
            self.base_path_display.config(text="Nenhum caminho selecionado")

    def enviar_dados_base(self):
        self.status_label.config(text="")
        if not self.base_path:
            messagebox.showerror("Erro", "Selecione o caminho da base antes de enviar os dados.")
            return
        if not self.tratada_path:
            messagebox.showerror("Erro", "Selecione a planilha tratada.")
            return
        
        self.extract_report_button.config(state=tk.DISABLED)
        self.tratamento_button.config(state=tk.DISABLED)
        self.adjust_contingency_button.config(state=tk.DISABLED)
        self.envio_button.config(state=tk.DISABLED)
        self.loading_label.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=20, pady=5)

        try:
            resultado = incluir_dados_na_base(self.tratada_path, self.base_path)
            if resultado is True:
                messagebox.showinfo("Envio", "Dados incluídos na base com sucesso!")
                self.status_label.config(text="Dados incluídos na base!", fg="green")
            else:
                messagebox.showerror("Erro", f"Erro ao incluir dados na base:\n{resultado}")
                self.status_label.config(text="Erro ao incluir dados na base.", fg="red")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado ao enviar dados para a base: {e}")
            self.status_label.config(text=f"Erro inesperado no envio: {e}", fg="red")
        finally:
            self.loading_label.grid_forget()
            self.extract_report_button.config(state=tk.NORMAL)
            self.tratamento_button.config(state=tk.NORMAL)
            self.adjust_contingency_button.config(state=tk.NORMAL)
            self.envio_button.config(state=tk.NORMAL)

    def ajustar_valor_contingencial(self):
        self.status_label.config(text="")
        if not self.planilha_path:
            messagebox.showerror("Erro", "Selecione a planilha antes de ajustar o valor contingencial.")
            return
        novo_valor = self.contingencia_entry.get()
        if not novo_valor:
            messagebox.showerror("Erro", "Preencha o valor contingencial.")
            return
        
        self.extract_report_button.config(state=tk.DISABLED)
        self.tratamento_button.config(state=tk.DISABLED)
        self.adjust_contingency_button.config(state=tk.DISABLED)
        self.envio_button.config(state=tk.DISABLED)
        self.loading_label.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=20, pady=5)

        try:
            resultado = ajustar_contingencial(self.planilha_path, novo_valor)
            if resultado is True:
                messagebox.showinfo("Contingencial", "Valor contingencial ajustado com sucesso!")
                self.status_label.config(text="Valor contingencial ajustado!", fg="green")
            else:
                messagebox.showerror("Erro", f"Erro ao ajustar valor contingencial:\n{resultado}")
                self.status_label.config(text="Erro ao ajustar valor contingencial.", fg="red")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado ao ajustar valor contingencial: {e}")
            self.status_label.config(text=f"Erro inesperado no ajuste: {e}", fg="red")
        finally:
            self.loading_label.grid_forget()
            self.extract_report_button.config(state=tk.NORMAL)
            self.tratamento_button.config(state=tk.NORMAL)
            self.adjust_contingency_button.config(state=tk.NORMAL)
            self.envio_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = AtenaCommanderApp(root)
    root.mainloop()