import pandas as pd
import os

item = "01.01.04.000575"
arquivo_excel = r"integrao_a_28-01-2026 (1).xlsx"

if os.path.exists(arquivo_excel):
    print(f"Lendo {arquivo_excel}...")
    try:
        df = pd.read_excel(arquivo_excel)
        print(f"Dimensões: {df.shape}")
        print(f"Colunas: {list(df.columns)}")
        
        # Procurar o item em todas as colunas
        encontrado = False
        for col in df.columns:
            if item in str(df[col].astype(str)).replace(" ", ""):
                print(f"\nEncontrado em coluna: {col}")
                mask = df[col].astype(str).str.contains(item, na=False)
                linhas = df[mask]
                print(linhas.to_string())
                encontrado = True
        
        if not encontrado:
            print(f"\nItem {item} NÃO encontrado no arquivo!")
            print("\nProcurando por itens similares...")
            for col in df.columns:
                conteudo = df[col].astype(str)
                if "01.01.04" in conteudo.str.cat():
                    print(f"\nEncontrado 01.01.04.* em coluna '{col}':")
                    mask = conteudo.str.contains("01.01.04", na=False)
                    print(df[mask][[col]].head(10).to_string())
    except Exception as e:
        print(f"Erro: {e}")
else:
    print(f"Arquivo não encontrado: {arquivo_excel}")
