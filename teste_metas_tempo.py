"""
Script para testar Metas de Tempo por Fase (cálculo mensal)
"""

import json
from datetime import datetime, timezone
from collections import defaultdict

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"Total de cards: {len(cards)}")

# === ANÁLISE TEMPO MENSAL POR FASE ===
tempo_mensal_por_fase = defaultdict(lambda: defaultdict(list))

for card in cards:
    # Buscar tempo de fase nos campos do Pipefy (prioridade)
    tempos_por_fase_card = {}
    for field in card.get('fields', []):
        fn = (field.get('name') or '').lower()
        fv = field.get('value')
        
        # Procurar por campos "Tempo total na fase X (dias)"
        if 'tempo total na fase' in fn and '(dias)' in fn and fv:
            try:
                fase_nome = fn.replace('tempo total na fase', '').replace('(dias)', '').strip().title()
                valor_str = str(fv).replace(',', '.')
                dias = float(valor_str)
                if fase_nome not in tempos_por_fase_card or dias > tempos_por_fase_card[fase_nome]:
                    tempos_por_fase_card[fase_nome] = dias
            except:
                pass
    
    # Se encontrou valores do Pipefy, adicionar ao agrupamento mensal
    if tempos_por_fase_card:
        finished = card.get('finished_at', '')
        if finished:
            try:
                dt_fin = datetime.fromisoformat(finished.replace('Z', '+00:00'))
                mes = dt_fin.strftime('%Y-%m')
            except:
                mes = datetime.now(timezone.utc).strftime('%Y-%m')
        else:
            mes = datetime.now(timezone.utc).strftime('%Y-%m')
        
        if mes >= '2026-01':
            for ph_name, dias in tempos_por_fase_card.items():
                tempo_mensal_por_fase[mes][ph_name].append(dias)
    else:
        # Fallback: calcular manualmente (incluindo cards em andamento)
        fases_calculadas = {}
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
                    
                    # Calcular tempo em dias
                    dias = (end - start).total_seconds() / (3600 * 24)
                    
                    # Armazenar apenas o maior tempo para cada fase
                    if ph_name not in fases_calculadas or dias > fases_calculadas[ph_name]:
                        fases_calculadas[ph_name] = dias
                except:
                    pass
        
        # Adicionar as fases calculadas ao agrupamento mensal
        if fases_calculadas:
            finished = card.get('finished_at', '')
            if finished:
                try:
                    dt_fin = datetime.fromisoformat(finished.replace('Z', '+00:00'))
                    mes = dt_fin.strftime('%Y-%m')
                except:
                    mes = datetime.now(timezone.utc).strftime('%Y-%m')
            else:
                mes = datetime.now(timezone.utc).strftime('%Y-%m')
            
            if mes >= '2026-01':
                for ph_name, dias in fases_calculadas.items():
                    tempo_mensal_por_fase[mes][ph_name].append(dias)

# Calcular médias mensais por fase
metas_por_fase = {}
for mes, fases_dict in tempo_mensal_por_fase.items():
    for fase, tempos in fases_dict.items():
        if fase not in metas_por_fase:
            metas_por_fase[fase] = []
        media_mes = sum(tempos) / len(tempos) if tempos else 0
        metas_por_fase[fase].append(media_mes)

# Calcular média geral por fase
print("\n📊 METAS DE TEMPO POR FASE (Janeiro 2026):")
print("=" * 80)

fases_ordenadas = sorted(metas_por_fase.items(), key=lambda x: sum(x[1])/len(x[1]) if x[1] else 0, reverse=True)

for fase, medias_mensais in fases_ordenadas[:10]:
    if medias_mensais:
        media_geral = sum(medias_mensais) / len(medias_mensais)
        print(f"{fase[:40]:40} | {media_geral:6.1f} dias (média de {len(medias_mensais)} mês(es))")

# Mostrar detalhes da fase Cotação
print("\n\n🔍 DETALHES - FASE COTAÇÃO:")
print("=" * 80)
if 'Cotação' in tempo_mensal_por_fase.get('2026-01', {}):
    valores = tempo_mensal_por_fase['2026-01']['Cotação']
    print(f"Janeiro 2026: {len(valores)} cards")
    print(f"Média: {sum(valores)/len(valores):.2f} dias")
    print(f"Mínimo: {min(valores):.2f} dias")
    print(f"Máximo: {max(valores):.2f} dias")
