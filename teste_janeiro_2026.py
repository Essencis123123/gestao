"""
Teste: Média de Atendimento apenas de cards CRIADOS em Janeiro/2026
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
            # Armazena dias, ano de criação, mês de criação e ano da conclusão/atual
            tempo_ciclo.append((dias, dt_cri.year, dt_fin.year, dt_cri.strftime('%Y-%m'), card.get('title', 'Sem título')))
        except Exception as e:
            print(f"Erro: {e}")

# Média de atendimento apenas de cards CRIADOS em Janeiro/2026
mes_atual = '2026-01'
tempo_ciclo_mes = [t for t in tempo_ciclo if t[3] == mes_atual]

print(f"\n📊 Cards criados em Janeiro/2026: {len(tempo_ciclo_mes)}")

if tempo_ciclo_mes:
    tempo_ciclo_medio = sum(t[0] for t in tempo_ciclo_mes) / len(tempo_ciclo_mes)
    print(f"✅ Média de Atendimento: {tempo_ciclo_medio:.2f} dias")
    
    print(f"\n📋 Primeiros 15 cards:")
    for i, (dias, _, _, mes, titulo) in enumerate(tempo_ciclo_mes[:15], 1):
        status = "✓" if dias < 10 else "⚠"
        print(f"  {status} {i:2}. {dias:6.2f} dias | {titulo[:50]}")
else:
    print("❌ Nenhum card criado em Janeiro/2026")
