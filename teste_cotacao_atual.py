"""
Teste: Tempo na fase Cotação apenas para cards que ESTÃO na Cotação atualmente
"""

import json
from datetime import datetime, timezone

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"📦 Total de cards: {len(cards)}")

hoje = datetime.now(timezone.utc)

# Filtrar apenas cards que estão ATUALMENTE na fase Cotação
cards_na_cotacao = []
for card in cards:
    phase = card.get('current_phase', {})
    phase_name = phase.get('name', '') if phase else ''
    
    if phase_name == 'Cotação':
        cards_na_cotacao.append(card)

print(f"📊 Cards atualmente na fase Cotação: {len(cards_na_cotacao)}")

# Calcular tempo na Cotação
tempos_cotacao = []

for card in cards_na_cotacao:
    titulo = card.get('title', 'Sem título')
    
    # Tentar buscar dos campos do Pipefy primeiro
    tempo_campo = None
    for field in card.get('fields', []):
        fn = (field.get('name') or '').lower()
        fv = field.get('value')
        
        if 'tempo total na fase' in fn and 'cotação' in fn and '(dias)' in fn and fv:
            try:
                valor_str = str(fv).replace(',', '.')
                tempo_campo = float(valor_str)
                break
            except:
                pass
    
    if tempo_campo is not None:
        tempos_cotacao.append((tempo_campo, titulo, 'campo'))
    else:
        # Fallback: calcular manualmente
        for ph in card.get('phases_history', []):
            ph_name = ph.get('phase', {}).get('name', '')
            
            if ph_name == 'Cotação':
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
                        tempos_cotacao.append((dias, titulo, 'calculado'))
                        break
                    except:
                        pass

print(f"\n📈 Cards com tempo calculado: {len(tempos_cotacao)}")

if tempos_cotacao:
    media = sum(t[0] for t in tempos_cotacao) / len(tempos_cotacao)
    print(f"\n✅ TEMPO MÉDIO NA FASE COTAÇÃO: {media:.2f} dias")
    print(f"   Mínimo: {min(t[0] for t in tempos_cotacao):.2f} dias")
    print(f"   Máximo: {max(t[0] for t in tempos_cotacao):.2f} dias")
    
    print(f"\n📋 TODOS OS CARDS NA COTAÇÃO:")
    print("=" * 90)
    tempos_cotacao.sort(key=lambda x: x[0], reverse=True)
    
    for i, (dias, titulo, origem) in enumerate(tempos_cotacao, 1):
        status = "🟢" if dias < 5 else "🟡" if dias < 10 else "🔴"
        print(f"{status} {i:2}. {dias:7.2f} dias | {titulo[:60]}")
else:
    print("\n❌ Nenhum card com tempo calculado")
