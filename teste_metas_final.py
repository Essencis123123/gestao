"""
Teste Final: Verificar ajustes de metas e filtros
"""

import json
from datetime import datetime, timezone
from collections import defaultdict

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"📦 Total de cards: {len(cards)}\n")

hoje = datetime.now(timezone.utc)

# === 1. TEMPO POR FASE (gráfico - apenas cards atualmente na fase) ===
print("=" * 80)
print("1️⃣  TEMPO POR FASE (GRÁFICO)")
print("=" * 80)

fases_tempo = defaultdict(list)

for card in cards:
    phase = card.get('current_phase', {})
    phase_name = phase.get('name', '') if phase else ''
    
    if not phase_name:
        continue
    
    # Calcular tempo
    for ph in card.get('phases_history', []):
        ph_name = ph.get('phase', {}).get('name', '')
        
        if ph_name == phase_name:
            first_in = ph.get('firstTimeIn')
            last_out = ph.get('lastTimeOut')
            
            if first_in:
                try:
                    start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                    if last_out:
                        end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                    else:
                        end = hoje
                    
                    dias = (end - start).total_seconds() / (3600 * 24)
                    fases_tempo[phase_name].append(dias)
                    break
                except:
                    pass

# Filtrar e ordenar
fases_analise = []
for fase, tempos in fases_tempo.items():
    # Filtrar apenas reprovados/cancelados
    if any(x in fase.lower() for x in ['reprovada', 'cancelada']):
        continue
    
    if tempos:
        media_dias = sum(tempos) / len(tempos)
        fases_analise.append((fase, media_dias, len(tempos)))

fases_analise.sort(key=lambda x: x[1], reverse=True)

for i, (fase, media, total) in enumerate(fases_analise[:8], 1):
    print(f"{i}. {fase[:35]:35} | {media:6.2f} dias ({total:2} cards)")

# === 2. METAS (apenas cards atualmente na fase) ===
print("\n" + "=" * 80)
print("2️⃣  METAS DE TEMPO POR FASE")
print("=" * 80)

tempo_mensal = defaultdict(lambda: defaultdict(list))

for card in cards:
    phase = card.get('current_phase', {})
    phase_name = phase.get('name', '') if phase else ''
    
    if not phase_name:
        continue
    
    # Calcular tempo na fase atual
    for ph in card.get('phases_history', []):
        ph_name = ph.get('phase', {}).get('name', '')
        
        if ph_name == phase_name:
            first_in = ph.get('firstTimeIn')
            last_out = ph.get('lastTimeOut')
            
            if first_in:
                try:
                    start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                    if last_out:
                        end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                    else:
                        end = hoje
                    
                    dias = (end - start).total_seconds() / (3600 * 24)
                    
                    # Determinar mês
                    finished = card.get('finished_at', '')
                    if finished:
                        try:
                            dt_fin = datetime.fromisoformat(finished.replace('Z', '+00:00'))
                            mes = dt_fin.strftime('%Y-%m')
                        except:
                            mes = hoje.strftime('%Y-%m')
                    else:
                        mes = hoje.strftime('%Y-%m')
                    
                    if mes >= '2026-01':
                        tempo_mensal[mes][phase_name].append(dias)
                    break
                except:
                    pass

print(f"Mês: 2026-01\n")
if '2026-01' in tempo_mensal:
    fases_metas = []
    for fase, tempos in tempo_mensal['2026-01'].items():
        # Filtrar apenas reprovados/cancelados
        if any(x in fase.lower() for x in ['reprovada', 'cancelada']):
            continue
        
        if tempos:
            media = sum(tempos) / len(tempos)
            fases_metas.append((fase, media, len(tempos)))
    
    fases_metas.sort(key=lambda x: x[1], reverse=True)
    
    for i, (fase, media, total) in enumerate(fases_metas[:8], 1):
        print(f"{i}. {fase[:35]:35} | {media:6.2f} dias ({total:2} cards)")
else:
    print("Sem dados")
