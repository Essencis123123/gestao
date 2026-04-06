"""
Teste: Cálculos considerando APENAS cards criados em Janeiro/2026
"""

import json
from datetime import datetime, timezone
from collections import defaultdict

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"📦 Total de cards no arquivo: {len(cards)}")

mes_atual = '2026-01'
hoje = datetime.now(timezone.utc)

# Filtrar apenas cards criados em Janeiro/2026
cards_janeiro = []
for card in cards:
    created = card.get('created_at', '')
    if created:
        try:
            dt_cri = datetime.fromisoformat(created.replace('Z', '+00:00'))
            mes_criacao = dt_cri.strftime('%Y-%m')
            if mes_criacao == mes_atual:
                cards_janeiro.append(card)
        except:
            pass

print(f"📅 Cards criados em Janeiro/2026: {len(cards_janeiro)}")

# === 1. MÉDIA DE ATENDIMENTO ===
tempo_ciclo = []
for card in cards_janeiro:
    created = card.get('created_at', '')
    finished = card.get('finished_at', '')
    
    if created:
        try:
            dt_cri = datetime.fromisoformat(created.replace('Z', '+00:00'))
            
            if finished:
                dt_fin = datetime.fromisoformat(finished.replace('Z', '+00:00'))
            else:
                dt_fin = hoje
            
            dias = (dt_fin - dt_cri).total_seconds() / (3600 * 24)
            tempo_ciclo.append(dias)
        except:
            pass

if tempo_ciclo:
    media_atendimento = sum(tempo_ciclo) / len(tempo_ciclo)
    print(f"\n✅ MÉDIA DE ATENDIMENTO: {media_atendimento:.2f} dias")
else:
    print("\n❌ Sem dados de tempo de ciclo")

# === 2. TEMPO POR FASE (GRÁFICO) ===
fases_tempo = defaultdict(list)

for card in cards_janeiro:
    # Tentar buscar dos campos do Pipefy primeiro
    tempos_por_fase_card = {}
    for field in card.get('fields', []):
        fn = (field.get('name') or '').lower()
        fv = field.get('value')
        
        if 'tempo total na fase' in fn and '(dias)' in fn and fv:
            try:
                fase_nome = fn.replace('tempo total na fase', '').replace('(dias)', '').strip().title()
                valor_str = str(fv).replace(',', '.')
                dias = float(valor_str)
                if fase_nome not in tempos_por_fase_card or dias > tempos_por_fase_card[fase_nome]:
                    tempos_por_fase_card[fase_nome] = dias
            except:
                pass
    
    if tempos_por_fase_card:
        for fase_nome, dias in tempos_por_fase_card.items():
            fases_tempo[fase_nome].append(dias)
    else:
        # Fallback: calcular manualmente
        fases_processadas = {}
        for ph in card.get('phases_history', []):
            ph_name = ph.get('phase', {}).get('name', '')
            first_in = ph.get('firstTimeIn')
            last_out = ph.get('lastTimeOut')
            
            if first_in and ph_name:
                try:
                    start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                    if last_out:
                        end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                    else:
                        end = hoje
                    
                    dias = (end - start).total_seconds() / (3600 * 24)
                    if ph_name not in fases_processadas or dias > fases_processadas[ph_name]:
                        fases_processadas[ph_name] = dias
                except:
                    pass
        
        for ph_name, dias in fases_processadas.items():
            fases_tempo[ph_name].append(dias)

# Calcular médias
fases_analise = []
for fase, tempos in fases_tempo.items():
    if tempos:
        media_dias = sum(tempos) / len(tempos)
        fases_analise.append({
            'nome': fase,
            'media_dias': media_dias,
            'total': len(tempos)
        })

fases_analise.sort(key=lambda x: x['media_dias'], reverse=True)

print(f"\n📊 TEMPO MÉDIO POR FASE (TOP 10):")
print("=" * 80)
for i, f in enumerate(fases_analise[:10], 1):
    print(f"{i:2}. {f['nome'][:40]:40} | {f['media_dias']:6.2f} dias ({f['total']:3} cards)")

# Detalhe da Cotação
cotacao_tempos = fases_tempo.get('Cotação', [])
if cotacao_tempos:
    print(f"\n🔍 DETALHES - FASE COTAÇÃO:")
    print(f"   Média: {sum(cotacao_tempos)/len(cotacao_tempos):.2f} dias")
    print(f"   Mínimo: {min(cotacao_tempos):.2f} dias")
    print(f"   Máximo: {max(cotacao_tempos):.2f} dias")
    print(f"   Total de cards: {len(cotacao_tempos)}")
