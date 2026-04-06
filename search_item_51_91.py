"""
Debug: Rastrear item 01.01.04.000575 (Preço 51,91) nos arquivos Excel de Relatórios
"""
import pandas as pd
from datetime import datetime

ITEM_BUSCA = '01.01.04.000575'
PRECO_PROCURADO = 51.91

# Arquivos de relatórios contábeis
arquivos = [
    r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2025\Relatório de Detalhes da Distribuição por Conta.xlsx",
    r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2026\janeiro\Relatório de Detalhes da Distribuição por Conta.xlsx",
    r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil\2026\fevereiro\Relatório de Detalhes da Distribuição por Conta.xlsx",
]

print("=" * 80)
print(f"RASTREANDO ITEM: {ITEM_BUSCA}")
print(f"PRECO PROCURADO: R$ {PRECO_PROCURADO}")
print("=" * 80)

total_encontrados = 0

for arquivo in arquivos:
    print(f"\n📂 Tentando arquivo: {arquivo}")
    try:
        df = pd.read_excel(arquivo)
        
        # Converter colunas numéricas
        for col in ['Crédito', 'Débito', 'Quantidade', 'Valor', 'Preco', 'Preço']:
            if col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except:
                    pass
        
        # Buscar item
        if 'Item' in df.columns:
            registros = df[df['Item'] == ITEM_BUSCA]
            
            if len(registros) > 0:
                print(f"   ENCONTRADO! Total de registros: {len(registros)}")
                total_encontrados += len(registros)
                
                # Mostrar detalhes
                for idx, row in registros.iterrows():
                    print(f"\n   Linha {idx + 2}:")
                    
                    # Mostrar todas as colunas relevantes
                    for col in df.columns:
                        valor = row[col]
                        if pd.notna(valor):
                            print(f"      {col}: {valor}")
                    
                    # Se há Débito e Quantidade, calcular preço
                    if 'Débito' in row and 'Quantidade' in row:
                        if pd.notna(row['Débito']) and pd.notna(row['Quantidade']) and row['Quantidade'] > 0:
                            preco_calc = row['Débito'] / row['Quantidade']
                            print(f"      >>> PREÇO CALCULADO: R$ {preco_calc:.2f}")
                            if abs(preco_calc - PRECO_PROCURADO) < 0.01:
                                print(f"      *** CORRESPONDE AO PREÇO PROCURADO ***")
            else:
                print(f"   Nenhum registro encontrado")
                
                # Procurar por parte do item
                parcial = df[df['Item'].astype(str).str.contains('01.01.04', na=False)]
                if len(parcial) > 0:
                    itens_unicos = parcial['Item'].unique()
                    print(f"   Itens com '01.01.04': {len(itens_unicos)}")
                    for it in itens_unicos[:10]:
                        desc = parcial[parcial['Item'] == it]['Descrição do Item'].iloc[0] if 'Descrição do Item' in parcial.columns else ''
                        print(f"      → {it} | {desc}")
        else:
            print(f"   Coluna 'Item' não encontrada (colunas: {list(df.columns)[:5]}...)")
            
    except FileNotFoundError as e:
        print(f"   ❌ Arquivo não encontrado: {e}")
    except Exception as e:
        print(f"   ❌ Erro ao processar: {e}")

print("\n" + "=" * 80)
if total_encontrados > 0:
    print(f"RESULTADO: Item {ITEM_BUSCA} encontrado em {total_encontrados} registro(s)")
else:
    print(f"RESULTADO: Item {ITEM_BUSCA} NAO encontrado nos arquivos consultados")
    print("\nDica: O item pode estar")
    print("  - Em um mês diferente em: G:\\SUPRIMENTOS\\SUPRIMENTOS\\PROJETOS\\Relatórios\\Geral contabil\\2026\\")
    print("  - Registrado com um número de item ligeiramente diferente")
    print("  - Ainda não integrado ao relatório contábil")
