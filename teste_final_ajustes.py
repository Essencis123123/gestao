"""
Teste Final: Verificação completa dos ajustes
"""

import json
from datetime import datetime, timezone
from collections import defaultdict

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"📦 Total de cards: {len(cards)}\n")

hoje = datetime.now(timezone.utc)

# === 1. TEMPO POR FASE (apenas cards atualmente na fase) ===
print("=" * 80)
print("1️⃣  TEMPO POR FASE (gráfico)")
print("=" * 80)

fases_tempo = defaultdict(list)

for card in cards:
    phase = card.get('current_phase', {})
    phase_name = phase.get('name', '') if phase else ''
    
    if not phase_name:
        continue
    
    # Tentar buscar dos campos do Pipefy
    tempo_campo = None
    for field in card.get('fields', []):
        fn = (field.get('name') or '').lower()
        fv = field.get('value')
        
        if 'tempo total na fase' in fn and phase_name.lower() in fn and '(dias)' in fn and fv:
            try:
                valor_str = str(fv).replace(',', '.')
                tempo_campo = float(valor_str)
                break
            except:
                pass
    
    if tempo_campo is not None:
        fases_tempo[phase_name].append(tempo_campo)
    else:
        # Fallback: calcular manualmente
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

# Calcular médias e filtrar Solicitação - Reprovada/Cancelada
fases_analise = []
for fase, tempos in fases_tempo.items():
    # Filtrar reprovados/cancelados
    if 'reprovada' in fase.lower() or 'cancelada' in fase.lower():
        continue
    
    if tempos:
        media_dias = sum(tempos) / len(tempos)
        fases_analise.append({
            'nome': fase,
            'media_dias': media_dias,
            'total': len(tempos)
        })

fases_analise.sort(key=lambda x: x['media_dias'], reverse=True)

for i, f in enumerate(fases_analise[:8], 1):
    print(f"{i}. {f['nome'][:35]:35} | {f['media_dias']:6.2f} dias ({f['total']:2} cards)")

# === 2. MÉDIA DE ATENDIMENTO (mesma da Cotação) ===
print("\n" + "=" * 80)
print("2️⃣  MÉDIA DE ATENDIMENTO")
print("=" * 80)

tempo_ciclo_medio = 0
for f in fases_analise:
    if f['nome'] == 'Cotação':
        tempo_ciclo_medio = f['media_dias']
        break

print(f"Média de Atendimento: {tempo_ciclo_medio:.2f} dias (igual à fase Cotação)")

# === 3. METAS DE TEMPO POR FASE ===
print("\n" + "=" * 80)
print("3️⃣  METAS DE TEMPO POR FASE")
print("=" * 80)

tempo_mensal_por_fase = defaultdict(lambda: defaultdict(list))

for card in cards:
    phase = card.get('current_phase', {})
    phase_name = phase.get('name', '') if phase else ''
    
    # Buscar tempos do Pipefy
    tempos_por_fase = {}
    for field in card.get('fields', []):
        fn = (field.get('name') or '').lower()
        fv = field.get('value')
        
        if 'tempo total na fase' in fn and '(dias)' in fn and fv:
            try:
                fase_nome = fn.replace('tempo total na fase', '').replace('(dias)', '').strip().title()
                valor_str = str(fv).replace(',', '.')
                dias = float(valor_str)
                if fase_nome not in tempos_por_fase or dias > tempos_por_fase[fase_nome]:
                    tempos_por_fase[fase_nome] = dias
            except:
                pass
    
    if tempos_por_fase:
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
            for fase_nome, dias in tempos_por_fase.items():
                # Filtrar reprovados/cancelados
                if 'reprovada' not in fase_nome.lower() and 'cancelada' not in fase_nome.lower():
                    tempo_mensal_por_fase[mes][fase_nome].append(dias)

# Calcular médias
print(f"Mês de referência: 2026-01\n")
if '2026-01' in tempo_mensal_por_fase:
    fases_metas = []
    for fase, tempos in tempo_mensal_por_fase['2026-01'].items():
        if tempos:
            media = sum(tempos) / len(tempos)
            fases_metas.append((fase, media, len(tempos)))
    
    fases_metas.sort(key=lambda x: x[1], reverse=True)
    
    for i, (fase, media, total) in enumerate(fases_metas[:8], 1):
        print(f"{i}. {fase[:35]:35} | {media:6.2f} dias ({total:2} cards)")
else:
    print("Sem dados para Janeiro/2026")
