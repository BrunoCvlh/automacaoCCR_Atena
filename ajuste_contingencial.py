import pandas as pd

def ajustar_contingencial(file_path, novo_valor):
    try:
        df = pd.read_excel(file_path)
        mask = df.iloc[:, 0] == "1.06 - CONTIGENCIAL"
        df.loc[mask, df.columns[1]] = novo_valor
        df.to_excel(file_path, index=False)
        return True
    except Exception as e:
        return f"ERRO: {e}"