"""
Teste de filtro de mês
"""
import json
from datetime import datetime, timezone

# Carregar dados
with open('pipefy_compras_servicos.json', 'r', encoding='utf-8') as f:
    cards = json.load(f)

print(f"Total original: {len(cards)} cards\n")

# Testar filtro Janeiro 2026
mes_filtro = '2026-01'
cards_filtrados = []

for card in cards:
    created = card.get('created_at', '')
    if created:
        try:
            dt_cri = datetime.fromisoformat(created.replace('Z', '+00:00'))
            mes_criacao = dt_cri.strftime('%Y-%m')
            if mes_criacao == mes_filtro:
                cards_filtrados.append(card)
        except:
            pass

print(f"Filtrado {mes_filtro}: {len(cards_filtrados)} cards")

# Contar por mês
meses = {}
for card in cards:
    created = card.get('created_at', '')
    if created:
        try:
            dt_cri = datetime.fromisoformat(created.replace('Z', '+00:00'))
            mes = dt_cri.strftime('%Y-%m')
            meses[mes] = meses.get(mes, 0) + 1
        except:
            pass

print(f"\nDistribuição por mês:")
for mes in sorted(meses.keys(), reverse=True)[:6]:
    print(f"  {mes}: {meses[mes]} cards")
