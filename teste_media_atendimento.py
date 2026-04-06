"""
Script para testar o cálculo da Média de Atendimento 
(mesmo cálculo do main.py, incluindo cards em andamento)
"""

import json
from datetime import datetime, timezone

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"Total de cards: {len(cards)}")

# Calcular tempo de ciclo (incluindo cards em andamento)
tempo_ciclo = []

for card in cards:
    created = card.get('created_at', '')
    finished = card.get('finished_at', '')
    
    if created:
        try:
            dt_cri = datetime.fromisoformat(created.replace('Z', '+00:00'))
            
            # Se finalizado, usar data de conclusão; senão, calcular até agora
            if finished:
                dt_fin = datetime.fromisoformat(finished.replace('Z', '+00:00'))
            else:
                dt_fin = datetime.now(timezone.utc)
            
            dias = (dt_fin - dt_cri).total_seconds() / (3600 * 24)
            # Armazena dias, ano de criação e ano da conclusão/atual
            tempo_ciclo.append((dias, dt_cri.year, dt_fin.year, card.get('title', 'Sem título')))
        except Exception as e:
            print(f"Erro no card {card.get('id')}: {e}")

print(f"\nTotal de cards com tempo calculado: {len(tempo_ciclo)}")

# Média de atendimento do ANO atual (apenas cards CRIADOS e CONCLUÍDOS em 2026)
hoje = datetime.now(timezone.utc)
ano_atual = hoje.year
tempo_ciclo_ano = [t for t in tempo_ciclo if t[1] == ano_atual and t[2] == ano_atual]

print(f"\nCards criados E concluídos em {ano_atual}: {len(tempo_ciclo_ano)}")

if tempo_ciclo_ano:
    tempo_ciclo_medio = sum(t[0] for t in tempo_ciclo_ano) / len(tempo_ciclo_ano)
    print(f"Média de Atendimento: {tempo_ciclo_medio:.2f} dias")
    
    print(f"\nPrimeiros 10 cards:")
    for i, (dias, ano_cri, ano_fin, titulo) in enumerate(tempo_ciclo_ano[:10], 1):
        print(f"  {i}. {dias:.2f} dias | {titulo}")
else:
    print("Nenhum card criado e concluído em 2026")

# Verificar também apenas cards criados em 2026 (incluindo em andamento)
tempo_ciclo_criados_2026 = [t for t in tempo_ciclo if t[1] == ano_atual]
print(f"\n\nCards criados em {ano_atual} (incluindo em andamento): {len(tempo_ciclo_criados_2026)}")

if tempo_ciclo_criados_2026:
    tempo_medio_todos = sum(t[0] for t in tempo_ciclo_criados_2026) / len(tempo_ciclo_criados_2026)
    print(f"Média considerando em andamento: {tempo_medio_todos:.2f} dias")
