"""Script para investigar de onde vem o item 01.01.18.000009 na análise de preços."""
import pandas as pd
import os

item_busca = '01.01.18.000009'

# 1. Verificar no arquivo de 2025
arquivo_2025 = r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2025\Relatório de Detalhes da Distribuição por Conta.xlsx"
print("=" * 80)
print(f"INVESTIGAÇÃO DO ITEM: {item_busca}")
print("=" * 80)

if os.path.exists(arquivo_2025):
    print(f"\n📁 Carregando arquivo 2025...")
    df_2025 = pd.read_excel(arquivo_2025)
    
    # Buscar o item
    mask = df_2025['Item'].astype(str).str.contains(item_busca, na=False)
    encontrados = df_2025[mask]
    
    if len(encontrados) > 0:
        print(f"✅ ENCONTRADO em 2025: {len(encontrados)} registros")
        # Filtrar apenas ordens de compra
        if 'Tipo de Referência' in encontrados.columns:
            oc = encontrados[encontrados['Tipo de Referência'] == 'Ordem de compra']
            print(f"   → Ordens de Compra: {len(oc)} registros")
        
        cols_mostrar = ['Item', 'Descrição do Item', 'Data da Transação', 'Débito', 'Crédito', 'Quantidade', 'Tipo de Referência']
        cols_disponiveis = [c for c in cols_mostrar if c in encontrados.columns]
        print(f"\n   Detalhes:")
        for _, row in encontrados[cols_disponiveis].iterrows():
            debito = float(row.get('Débito', 0)) if pd.notna(row.get('Débito')) else 0
            qtd = float(row.get('Quantidade', 0)) if pd.notna(row.get('Quantidade')) else 0
            preco_unit = debito / qtd if qtd > 0 else 0
            print(f"   Data: {row.get('Data da Transação', 'N/A')} | Tipo: {row.get('Tipo de Referência', 'N/A')} | Débito: {debito:.2f} | Qtd: {qtd:.2f} | Preço Unit: {preco_unit:.2f}")
    else:
        print(f"❌ NÃO encontrado no arquivo de 2025")
else:
    print(f"⚠️ Arquivo 2025 não encontrado: {arquivo_2025}")

# 2. Verificar no arquivo de Janeiro 2026
arquivo_2026 = r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2026\janeiro\Relatório de Detalhes da Distribuição por Conta.xlsx"
print(f"\n{'='*80}")

if os.path.exists(arquivo_2026):
    print(f"\n📁 Carregando arquivo Janeiro 2026...")
    df_2026 = pd.read_excel(arquivo_2026)
    
    # Buscar o item
    mask = df_2026['Item'].astype(str).str.contains(item_busca, na=False)
    encontrados = df_2026[mask]
    
    if len(encontrados) > 0:
        print(f"✅ ENCONTRADO em Jan/2026: {len(encontrados)} registros")
        if 'Tipo de Referência' in encontrados.columns:
            oc = encontrados[encontrados['Tipo de Referência'] == 'Ordem de compra']
            print(f"   → Ordens de Compra: {len(oc)} registros")
        
        cols_mostrar = ['Item', 'Descrição do Item', 'Data da Transação', 'Débito', 'Crédito', 'Quantidade', 'Tipo de Referência']
        cols_disponiveis = [c for c in cols_mostrar if c in encontrados.columns]
        print(f"\n   Detalhes:")
        for _, row in encontrados[cols_disponiveis].iterrows():
            debito = float(row.get('Débito', 0)) if pd.notna(row.get('Débito')) else 0
            qtd = float(row.get('Quantidade', 0)) if pd.notna(row.get('Quantidade')) else 0
            preco_unit = debito / qtd if qtd > 0 else 0
            print(f"   Data: {row.get('Data da Transação', 'N/A')} | Tipo: {row.get('Tipo de Referência', 'N/A')} | Débito: {debito:.2f} | Qtd: {qtd:.2f} | Preço Unit: {preco_unit:.2f}")
    else:
        print(f"❌ NÃO encontrado no arquivo de Janeiro 2026")
else:
    print(f"⚠️ Arquivo Jan/2026 não encontrado: {arquivo_2026}")

# 3. Simular o cálculo que o sistema faz
print(f"\n{'='*80}")
print("SIMULAÇÃO DO CÁLCULO DO SISTEMA:")
print(f"{'='*80}")

dfs = []
if os.path.exists(arquivo_2025):
    df1 = pd.read_excel(arquivo_2025)
    for col in ['Crédito', 'Débito', 'Quantidade']:
        if col in df1.columns:
            df1[col] = pd.to_numeric(df1[col], errors='coerce').fillna(0)
    dfs.append(df1)
    
if os.path.exists(arquivo_2026):
    df2 = pd.read_excel(arquivo_2026)
    for col in ['Crédito', 'Débito', 'Quantidade']:
        if col in df2.columns:
            df2[col] = pd.to_numeric(df2[col], errors='coerce').fillna(0)
    dfs.append(df2)

if dfs:
    df = pd.concat(dfs, ignore_index=True)
    
    # Filtrar
    if 'Tipo de Referência' in df.columns:
        df = df[df['Tipo de Referência'] == 'Ordem de compra']
    
    mask = df['Item'].astype(str).str.contains(item_busca, na=False)
    df_item = df[mask]
    df_item = df_item[(df_item['Débito'] > 0) & (df_item['Quantidade'] > 0)].copy()
    
    if len(df_item) > 0:
        df_item['Preco_Unitario'] = df_item['Débito'] / df_item['Quantidade']
        df_item['Mes'] = pd.to_datetime(df_item['Data da Transação']).dt.to_period('M')
        
        print(f"\nRegistros filtrados (OC, Débito>0, Qtd>0): {len(df_item)}")
        for _, row in df_item.iterrows():
            print(f"  Mês: {row['Mes']} | Débito: {row['Débito']:.2f} | Qtd: {row['Quantidade']:.2f} | Preço Unit: {row['Preco_Unitario']:.2f}")
        
        # Agrupar
        precos_mensais = df_item.groupby('Mes').agg({'Preco_Unitario': 'mean'}).reset_index()
        precos_mensais['Mes_str'] = precos_mensais['Mes'].astype(str)
        precos_mensais['Ano'] = precos_mensais['Mes'].apply(lambda x: x.year)
        
        print(f"\nPreço médio por mês:")
        for _, row in precos_mensais.iterrows():
            print(f"  {row['Mes_str']}: R$ {row['Preco_Unitario']:.2f}")
        
        # Cálculo do sistema
        precos_2025 = precos_mensais[precos_mensais['Ano'] == 2025]['Preco_Unitario']
        preco_inicial = precos_2025.mean() if len(precos_2025) > 0 else precos_mensais.iloc[0]['Preco_Unitario']
        
        mes_anterior = '2026-01'  # Fev 2026, então mês anterior = Jan 2026
        precos_jan = precos_mensais[precos_mensais['Mes_str'] == mes_anterior]['Preco_Unitario']
        preco_final = precos_jan.mean() if len(precos_jan) > 0 else precos_mensais.iloc[-1]['Preco_Unitario']
        
        variacao = ((preco_final - preco_inicial) / preco_inicial * 100) if preco_inicial > 0 else 0
        
        print(f"\n📊 RESULTADO DO CÁLCULO:")
        print(f"  Preço Inicial (média 2025): R$ {preco_inicial:.2f}")
        print(f"  Preço Final (Jan/2026):     R$ {preco_final:.2f}")
        print(f"  Variação:                   {variacao:+.1f}%")
        print(f"  Variação R$:                R$ {preco_final - preco_inicial:.2f}")
    else:
        print("❌ Nenhum registro encontrado após filtros (OC + Débito>0 + Qtd>0)")
