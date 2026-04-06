import pandas as pd

# Teste simples do parsing CSV
csv_path = r'G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Conferencia_NFE\Synchro\relatorio-recepcionadas.csv'

print("Testando parsing manual...")

try:
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()

    print(f"Total de linhas: {len(lines)}")

    # Processar cabeçalho
    header_line = lines[0].strip()
    print(f"Cabeçalho cru: {repr(header_line[:100])}")

    if header_line.startswith('"') and header_line.endswith('"'):
        header_line = header_line[1:-1]  # Remover aspas externas

    headers = [h.strip('"').strip() for h in header_line.split(';')]
    print(f"Cabeçalho processado: {headers[:5]}...")

    # Processar primeira linha de dados
    if len(lines) > 1:
        data_line = lines[1].strip()
        print(f"Primeira linha de dados crua: {repr(data_line[:100])}")

        if data_line.startswith('"') and data_line.endswith('"'):
            data_line = data_line[1:-1]

        fields = data_line.split(';')
        fields = [f.strip('"').strip() for f in fields]
        print(f"Campos processados: {fields[:5]}...")

        print("✅ Parsing funcionando!")

except Exception as e:
    print(f"❌ Erro: {e}")