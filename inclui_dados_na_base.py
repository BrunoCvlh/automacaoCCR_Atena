from tkinter import messagebox
import pandas as pd

def enviar_dados_base(self):
    if not self.base_path:
        messagebox.showerror("Erro", "Selecione o caminho da base antes de enviar os dados.")
        return
    if not self.planilha_path:
        messagebox.showerror("Erro", "Selecione a planilha de origem.")
        return
    resultado = incluir_dados_na_base(self.planilha_path, self.base_path)
    if resultado is True:
        messagebox.showinfo("Envio", "Dados incluídos na base com sucesso!")
    else:
        messagebox.showerror("Erro", f"Erro ao incluir dados: {resultado}")

def incluir_dados_na_base(planilha_origem, planilha_base):
    try:
        df_origem = pd.read_excel(planilha_origem)
        df_base = pd.read_excel(planilha_base)
        df_base = pd.concat([df_base, df_origem], ignore_index=True)
        df_base.to_excel(planilha_base, index=False)
        return True
    except Exception as e:
        return f"ERRO: {e}"