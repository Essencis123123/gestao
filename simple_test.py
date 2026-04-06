import pandas as pd
import os

print("Teste básico de leitura...")

# Verificar se os arquivos existem
csv_path = r'G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Conferencia_NFE\Synchro\relatorio-recepcionadas.csv'
xls_path = r'G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Conferencia_NFE\Oracle\export.xls'
cnpj_path = r'G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Conferencia_NFE\cnpj_clientes.txt'

print(f"CSV existe: {os.path.exists(csv_path)}")
print(f"XLS existe: {os.path.exists(xls_path)}")
print(f"CNPJ existe: {os.path.exists(cnpj_path)}")

if os.path.exists(csv_path):
    print("Tentando ler CSV...")
    try:
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
        print(f"Arquivo lido: {len(lines)} linhas")
        print(f"Primeira linha: {repr(lines[0][:50])}")
        print("✅ CSV OK")
    except Exception as e:
        print(f"❌ Erro no CSV: {e}")

if os.path.exists(cnpj_path):
    print("Tentando ler CNPJ...")
    try:
        with open(cnpj_path, 'r', encoding='utf-8') as f:
            cnpjs = f.readlines()
        print(f"CNPJs lidos: {len(cnpjs)} linhas")
        print("✅ CNPJ OK")
    except Exception as e:
        print(f"❌ Erro no CNPJ: {e}")

print("Teste concluído.")