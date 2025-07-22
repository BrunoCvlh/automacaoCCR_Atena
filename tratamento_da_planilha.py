import pandas as pd
import os
from datetime import datetime

def tratar_planilha(file_path, competencia):
    try:
        df = pd.read_excel(file_path)
        df = df.iloc[4:].reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        df = df.iloc[:, [0, 3, 4]]
        df.columns = [df.columns[0], "Orçado", "Realizado"]
        try:
            competencia_date = datetime.strptime(competencia, "%m/%Y")
            primeiro_dia = competencia_date.strftime("%d/%m/%Y")
        except Exception:
            primeiro_dia = ""
        df["Data Competência"] = primeiro_dia
        # Excluir as últimas duas linhas
        if len(df) > 2:
            df = df.iloc[:-2].reset_index(drop=True)
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        tratado_path = os.path.join(downloads_folder, "planilha_tratada.xlsx")
        df.to_excel(tratado_path, index=False)
        return tratado_path
    except Exception as e:
        return f"ERRO: {e}"