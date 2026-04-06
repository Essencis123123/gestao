"""
Debug: Rastrear item 01.01.01.000009 (Combustível) nos arquivos Excel
"""
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

ITEM_BUSCA = '01.01.01.000009'

# Arquivos
arquivo_2025 = r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2025\Relatório de Detalhes da Distribuição por Conta.xlsx"
arquivo_2026 = r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2026\janeiro\Relatório de Detalhes da Distribuição por Conta.xlsx"

print("=" * 80)
print(f"RASTREANDO ITEM: {ITEM_BUSCA}")
print("=" * 80)

# Carregar 2025
print("\n📂 Carregando arquivo 2025...")
df_2025 = pd.read_excel(arquivo_2025)
for col in ['Crédito', 'Débito', 'Quantidade']:
    if col in df_2025.columns:
        df_2025[col] = pd.to_numeric(df_2025[col], errors='coerce').fillna(0)
if 'Data da Transação' in df_2025.columns:
    df_2025['Data da Transação'] = pd.to_datetime(df_2025['Data da Transação'], errors='coerce')

# Carregar 2026
print("📂 Carregando arquivo 2026 (Janeiro)...")
df_2026 = pd.read_excel(arquivo_2026)
for col in ['Crédito', 'Débito', 'Quantidade']:
    if col in df_2026.columns:
        df_2026[col] = pd.to_numeric(df_2026[col], errors='coerce').fillna(0)
if 'Data da Transação' in df_2026.columns:
    df_2026['Data da Transação'] = pd.to_datetime(df_2026['Data da Transação'], errors='coerce')

# Buscar item em 2025
print(f"\n{'='*80}")
print(f"BUSCA NO ARQUIVO 2025:")
print(f"{'='*80}")
registros_2025 = df_2025[df_2025['Item'] == ITEM_BUSCA]
print(f"Total de registros encontrados: {len(registros_2025)}")

if len(registros_2025) > 0:
    # Filtrar Ordem de compra
    if 'Tipo de Referência' in registros_2025.columns:
        oc_2025 = registros_2025[registros_2025['Tipo de Referência'] == 'Ordem de compra']
        print(f"Registros 'Ordem de compra': {len(oc_2025)}")
        
        # Filtrar Débito > 0 e Quantidade > 0
        oc_validos = oc_2025[(oc_2025['Débito'] > 0) & (oc_2025['Quantidade'] > 0)]
        print(f"Registros válidos (Débito > 0 e Qtd > 0): {len(oc_validos)}")
        
        if len(oc_validos) > 0:
            for _, row in oc_validos.iterrows():
                preco = row['Débito'] / row['Quantidade']
                desc = row.get('Descrição do Item', 'N/A')
                print(f"  📌 {row['Data da Transação'].strftime('%m/%Y') if pd.notna(row['Data da Transação']) else '?'} | "
                      f"Débito: R$ {row['Débito']:,.2f} | Qtd: {row['Quantidade']:,.2f} | "
                      f"Preço Unit: R$ {preco:,.2f} | Desc: {desc}")
    else:
        print("Coluna 'Tipo de Referência' não encontrada")
else:
    # Tentar busca parcial
    print("❌ Nenhum registro encontrado com correspondência exata.")
    print("\n🔍 Buscando por correspondência parcial...")
    parcial = df_2025[df_2025['Item'].astype(str).str.contains('01.01.01', na=False)]
    itens_encontrados = parcial['Item'].unique()
    print(f"Itens que começam com '01.01.01': {len(itens_encontrados)}")
    for it in itens_encontrados[:20]:
        desc = ''
        if 'Descrição do Item' in parcial.columns:
            desc_rows = parcial[parcial['Item'] == it]['Descrição do Item'].dropna()
            desc = desc_rows.iloc[0] if len(desc_rows) > 0 else ''
        print(f"  → {it} | {desc}")

# Buscar item em 2026
print(f"\n{'='*80}")
print(f"BUSCA NO ARQUIVO 2026 (Janeiro):")
print(f"{'='*80}")
registros_2026 = df_2026[df_2026['Item'] == ITEM_BUSCA]
print(f"Total de registros encontrados: {len(registros_2026)}")

if len(registros_2026) > 0:
    if 'Tipo de Referência' in registros_2026.columns:
        oc_2026 = registros_2026[registros_2026['Tipo de Referência'] == 'Ordem de compra']
        print(f"Registros 'Ordem de compra': {len(oc_2026)}")
        
        oc_validos_2026 = oc_2026[(oc_2026['Débito'] > 0) & (oc_2026['Quantidade'] > 0)]
        print(f"Registros válidos (Débito > 0 e Qtd > 0): {len(oc_validos_2026)}")
        
        if len(oc_validos_2026) > 0:
            for _, row in oc_validos_2026.iterrows():
                preco = row['Débito'] / row['Quantidade']
                desc = row.get('Descrição do Item', 'N/A')
                print(f"  📌 {row['Data da Transação'].strftime('%m/%Y') if pd.notna(row['Data da Transação']) else '?'} | "
                      f"Débito: R$ {row['Débito']:,.2f} | Qtd: {row['Quantidade']:,.2f} | "
                      f"Preço Unit: R$ {preco:,.2f} | Desc: {desc}")
else:
    print("❌ Nenhum registro encontrado com correspondência exata.")
    print("\n🔍 Buscando por correspondência parcial...")
    parcial_2026 = df_2026[df_2026['Item'].astype(str).str.contains('01.01.01', na=False)]
    itens_encontrados_2026 = parcial_2026['Item'].unique()
    print(f"Itens que começam com '01.01.01': {len(itens_encontrados_2026)}")
    for it in itens_encontrados_2026[:20]:
        desc = ''
        if 'Descrição do Item' in parcial_2026.columns:
            desc_rows = parcial_2026[parcial_2026['Item'] == it]['Descrição do Item'].dropna()
            desc = desc_rows.iloc[0] if len(desc_rows) > 0 else ''
        print(f"  → {it} | {desc}")

# Simular lógica da análise
print(f"\n{'='*80}")
print(f"SIMULAÇÃO DA LÓGICA DE ANÁLISE:")
print(f"{'='*80}")
df_combined = pd.concat([df_2025, df_2026], ignore_index=True)
if 'Tipo de Referência' in df_combined.columns:
    df_combined = df_combined[df_combined['Tipo de Referência'] == 'Ordem de compra']
df_combined = df_combined[(df_combined['Débito'] > 0) & (df_combined['Quantidade'] > 0)].copy()
df_combined['Preco_Unitario'] = df_combined['Débito'] / df_combined['Quantidade']
df_combined['Mes'] = pd.to_datetime(df_combined['Data da Transação']).dt.to_period('M')

item_df = df_combined[df_combined['Item'] == ITEM_BUSCA]
print(f"Registros após filtros: {len(item_df)}")

if len(item_df) > 0:
    print("\nRegistros do item:")
    for _, row in item_df.iterrows():
        print(f"  {row['Mes']} | Preço: R$ {row['Preco_Unitario']:,.2f} | Qtd: {row['Quantidade']:,.2f}")
    
    # Mês anterior
    data_atual = datetime.now()
    data_mes_anterior = data_atual - relativedelta(months=1)
    mes_anterior = data_mes_anterior.strftime('%Y-%m')
    print(f"\nMês anterior (para análise): {mes_anterior}")
    
    item_df['Mes_str'] = item_df['Mes'].astype(str)
    item_df['Ano'] = item_df['Mes'].apply(lambda x: x.year)
    
    precos_2025_item = item_df[item_df['Ano'] == 2025]['Preco_Unitario']
    print(f"Preços 2025: {list(precos_2025_item.values)}")
    print(f"Média 2025 (preço inicial): R$ {precos_2025_item.mean():,.2f}" if len(precos_2025_item) > 0 else "Sem dados 2025")
    
    precos_mes_ant = item_df[item_df['Mes_str'] == mes_anterior]['Preco_Unitario']
    print(f"Preços no mês anterior ({mes_anterior}): {list(precos_mes_ant.values)}")
    
    if len(precos_mes_ant) == 0:
        print(f"⚠️ ITEM SERÁ EXCLUÍDO: sem dados no mês anterior ({mes_anterior})")
    else:
        print(f"Preço final (média mês anterior): R$ {precos_mes_ant.mean():,.2f}")
else:
    print("❌ Item não encontrado nos dados combinados")
    
    # Buscar itens similares nos dados combinados (combustível)
    print("\n🔍 Buscando itens com 'combust' na descrição...")
    if 'Descrição do Item' in df_combined.columns:
        combust = df_combined[df_combined['Descrição do Item'].astype(str).str.lower().str.contains('combust', na=False)]
        itens_combust = combust[['Item', 'Descrição do Item']].drop_duplicates(subset='Item')
        print(f"Itens encontrados com 'combustível': {len(itens_combust)}")
        for _, row in itens_combust.iterrows():
            print(f"  → {row['Item']} | {row['Descrição do Item']}")
