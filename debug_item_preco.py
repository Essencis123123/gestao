"""
Debug: Rastrear item 01.01.04.000575 (Preço 51,91) nos arquivos
"""
import os
import pandas as pd
import json
from pathlib import Path

# Pasta do projeto
pasta_projeto = r"C:\Users\2700024\Desktop\Ferramenta de gestão"

# Item que estamos procurando
item_código = "01.01.04.000575"
preço_procurado = 51.91

print(f"Procurando item {item_código} com preço {preço_procurado}")
print("=" * 80)

# Lista de arquivos para verificar
arquivos_para_verificar = [
    r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2025\Relatório de Detalhes da Distribuição por Conta.xlsx",
    r"integrao_a_28-01-2026 (1).xlsx",
]

# Adicionar arquivos do workspace
for root, dirs, files in os.walk(pasta_projeto):
    for file in files:
        if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.csv'):
            full_path = os.path.join(root, file)
            if full_path not in arquivos_para_verificar and "\\Backup\\" not in full_path and "\\build\\" not in full_path:
                arquivos_para_verificar.append(full_path)

resultados_encontrados = []

for arquivo in arquivos_para_verificar:
    if not os.path.exists(arquivo):
        print(f"❌ Arquivo não encontrado: {arquivo}")
        continue
    
    try:
        print(f"\n📄 Verificando: {arquivo}")
        
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            # Tentar ler Excel
            try:
                xls = pd.ExcelFile(arquivo)
                print(f"   Abas disponíveis: {xls.sheet_names}")
                
                for sheet_name in xls.sheet_names:
                    try:
                        df = pd.read_excel(arquivo, sheet_name=sheet_name)
                        print(f"   Verificando aba '{sheet_name}' ({len(df)} linhas, {len(df.columns)} colunas)")
                        
                        # Procurar por todas as colunas que possam conter o código do item ou preço
                        for col in df.columns:
                            col_str = str(col).lower()
                            
                            # Procurar o código do item
                            if item_código in str(df[col].astype(str)).replace(' ', ''):
                                mask = df[col].astype(str).str.contains(item_código, case=False, na=False)
                                encontrados = df[mask]
                                print(f"      ✓ Encontrado em coluna '{col}':")
                                for idx, row in encontrados.iterrows():
                                    print(f"         Linha {idx + 2}: {row.to_dict()}")
                                    resultados_encontrados.append({
                                        'arquivo': arquivo,
                                        'aba': sheet_name,
                                        'linha': idx + 2,
                                        'dados': row.to_dict()
                                    })
                    except Exception as e:
                        print(f"   ⚠ Erro ao ler aba '{sheet_name}': {str(e)}")
                        
            except Exception as e:
                print(f"   ⚠ Erro ao ler Excel: {str(e)}")
        
        elif arquivo.endswith('.csv'):
            # Tentar ler CSV
            try:
                df = pd.read_csv(arquivo)
                print(f"   Verificando CSV ({len(df)} linhas, {len(df.columns)} colunas)")
                
                for col in df.columns:
                    if item_código in str(df[col].astype(str)).replace(' ', ''):
                        mask = df[col].astype(str).str.contains(item_código, case=False, na=False)
                        encontrados = df[mask]
                        print(f"      ✓ Encontrado em coluna '{col}':")
                        for idx, row in encontrados.iterrows():
                            print(f"         Linha {idx + 1}: {row.to_dict()}")
                            resultados_encontrados.append({
                                'arquivo': arquivo,
                                'aba': 'CSV',
                                'linha': idx + 1,
                                'dados': row.to_dict()
                            })
            except Exception as e:
                print(f"   ⚠ Erro ao ler CSV: {str(e)}")
                
    except Exception as e:
        print(f"❌ Erro processando {arquivo}: {str(e)}")

print("\n" + "=" * 80)
print("RESUMO DOS RESULTADOS:")
print("=" * 80)

if resultados_encontrados:
    print(f"\n✓ Item {item_código} encontrado em {len(resultados_encontrados)} local(is)\n")
    for i, resultado in enumerate(resultados_encontrados, 1):
        print(f"\n{i}. {resultado['arquivo']}")
        print(f"   Aba/Seção: {resultado['aba']}")
        print(f"   Linha: {resultado['linha']}")
        print(f"   Dados:")
        for chave, valor in resultado['dados'].items():
            print(f"      {chave}: {valor}")
else:
    print(f"\n❌ Item {item_código} NÃO foi encontrado em nenhum arquivo!")
    
    # Dica: procurar por parte do código
    print(f"\n💡 Procurando por partes do código ({item_código[:8]}* ou *{item_código[-6:]})...")
    
    parte_inicio = item_código[:8]
    parte_fim = item_código[-6:]
    
    for arquivo in arquivos_para_verificar:
        if not os.path.exists(arquivo):
            continue
        try:
            if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
                xls = pd.ExcelFile(arquivo)
                for sheet_name in xls.sheet_names:
                    try:
                        df = pd.read_excel(arquivo, sheet_name=sheet_name)
                        for col in df.columns:
                            mask = (df[col].astype(str).str.contains(parte_inicio, case=False, na=False) or 
                                   df[col].astype(str).str.contains(parte_fim, case=False, na=False))
                            if mask.any():
                                encontrados = df[mask]
                                print(f"   Encontrado similar em {arquivo} (aba '{sheet_name}'):")
                                print(f"      {encontrados.head().to_string()}")
                                print()
                    except:
                        pass
        except:
            pass

print("\n" + "=" * 80)
print("Análise concluída!")
