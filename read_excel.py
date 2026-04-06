import pandas as pd

# Ler a planilha
df = pd.read_excel(r'C:\Users\2700024\Desktop\Relatórios\Pagamentos\2026\janeiro\SOLVI - Relatório de Títulos Pagos_AP_XLSX.xlsx', skiprows=2)

# A primeira linha contém os nomes das colunas
df.columns = df.iloc[0]
df = df[1:].reset_index(drop=True)

print('Colunas completas:')
for i, col in enumerate(df.columns):
    print(f'{i}: {col}')

print('\n\nPrimeiras 5 linhas:')
print(df.head())

print('\n\nColunas específicas que precisamos:')
colunas_importantes = ['Fornecedor', 'Valor da NFF', 'Forma de Pagamento', 'ISS', 'IR', 'INSS', 'PIS', 'COFINS', 'CSLL', 'Condição de Pagamento']
for col in colunas_importantes:
    if col in df.columns:
        print(f'✓ {col}')
    else:
        print(f'✗ {col} (não encontrada)')
        
print('\n\nAmostra de todas as colunas (primeiras 10):')
print(df.iloc[:5, :20])
