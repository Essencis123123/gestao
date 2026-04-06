"""
Script para testar o cálculo das fases (gráfico)
"""

import json
from datetime import datetime, timezone
from collections import defaultdict

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"Total de cards: {len(cards)}")

# === ANÁLISE POR FASE ===
fases_tempo = defaultdict(list)

for card in cards:
    # Buscar tempo de fase nos campos do Pipefy (prioridade)
    tempos_por_fase_card = {}
    for field in card.get('fields', []):
        fn = (field.get('name') or '').lower()
        fv = field.get('value')
        
        # Procurar por campos "Tempo total na fase X (dias)"
        if 'tempo total na fase' in fn and '(dias)' in fn and fv:
            try:
                # Extrair nome da fase do campo
                fase_nome = fn.replace('tempo total na fase', '').replace('(dias)', '').strip().title()
                # Converter vírgula para ponto (formato brasileiro -> Python)
                valor_str = str(fv).replace(',', '.')
                dias = float(valor_str)
                # Usar o maior valor se houver múltiplas entradas para mesma fase
                if fase_nome not in tempos_por_fase_card or dias > tempos_por_fase_card[fase_nome]:
                    tempos_por_fase_card[fase_nome] = dias
            except:
                pass
    
    # Se encontrou valores do Pipefy, usar esses (cada fase uma vez por card)
    if tempos_por_fase_card:
        for fase_nome, dias in tempos_por_fase_card.items():
            horas = dias * 24
            fases_tempo[fase_nome].append(horas)
    else:
        # Fallback: calcular manualmente (incluindo cards em andamento)
        fases_processadas = {}
        for ph in card.get('phases_history', []):
            ph_name = ph.get('phase', {}).get('name', '')
            first_in = ph.get('firstTimeIn')
            last_out = ph.get('lastTimeOut')
            
            if first_in and ph_name:
                try:
                    start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                    # Se não saiu da fase, calcular até agora (como Pipefy)
                    if last_out:
                        end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                    else:
                        end = datetime.now(timezone.utc)
                    
                    hours = (end - start).total_seconds() / 3600
                    # Armazenar apenas o maior tempo para cada fase
                    if ph_name not in fases_processadas or hours > fases_processadas[ph_name]:
                        fases_processadas[ph_name] = hours
                except:
                    pass
        
        # Adicionar as fases processadas aos tempos
        for ph_name, hours in fases_processadas.items():
            fases_tempo[ph_name].append(hours)

# Tempo médio e total por fase
fases_analise = []
for fase, tempos in fases_tempo.items():
    if tempos:
        media_horas = sum(tempos) / len(tempos)
        fases_analise.append({
            'nome': fase,
            'media_horas': media_horas,
            'media_dias': media_horas / 24,
            'total_passagens': len(tempos),
            'max_horas': max(tempos),
            'min_horas': min(tempos)
        })

fases_analise.sort(key=lambda x: x['media_horas'], reverse=True)

print(f"\n📊 TOP FASES POR TEMPO MÉDIO:")
print("=" * 80)
for i, f in enumerate(fases_analise[:10], 1):
    print(f"{i:2}. {f['nome'][:40]:40} | {f['media_dias']:6.1f} dias ({f['total_passagens']:3} cards)")
