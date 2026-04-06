"""Análise simplificada do item 01.01.18.000009"""

# Dados encontrados no arquivo 2025:
print("=" * 70)
print("ITEM: 01.01.18.000009 - HIPOCLORITO DE SODIO 12% BOMBONA DE 60KG")
print("=" * 70)

print("\nRegistros em 2025 (APENAS Ordens de Compra com Débito > 0):")
print("-" * 70)

# Registros encontrados (Tipo = Ordem de compra, Débito > 0, Qtd > 0):
registros = [
    {"data": "2025-01-31", "debito": 4000.00, "qtd": 25.00, "preco_unit": 160.00},
    {"data": "2025-05-27", "debito": 1924.00, "qtd": 13.00, "preco_unit": 148.00},
    {"data": "2025-07-14", "debito": 1776.00, "qtd": 600.00, "preco_unit": 2.96},
]

for r in registros:
    print(f"  {r['data']} | Débito: R$ {r['debito']:,.2f} | Qtd: {r['qtd']:.0f} | Preço Unit: R$ {r['preco_unit']:.2f}")

print("\n⚠️ PROBLEMA IDENTIFICADO:")
print("  Em 14/Jul/2025: Débito R$ 1.776 / Quantidade 600 = R$ 2,96")
print("  → A quantidade 600 parece ser em KG (bombona de 60KG x 10)")  
print("  → Enquanto em Jan/2025 foi: R$ 4.000 / 25 = R$ 160,00")
print("  → E em Mai/2025 foi: R$ 1.924 / 13 = R$ 148,00")
print("  → As unidades de medida NÃO são consistentes!")

# Cálculo do sistema
precos = [160.00, 148.00, 2.96]  # Média mensal por mês
media_2025 = sum(precos) / len(precos)
print(f"\n📊 O SISTEMA CALCULA:")
print(f"  Preço médio 2025 (inicial): R$ {media_2025:.2f}") 
print(f"  → Mas mistura R$ 160, R$ 148 e R$ 2,96!")
print(f"  → O R$ 2,96 distorce completamente a análise")

# Se não há dados de Jan/2026, usa último valor disponível
print(f"\n  Como o item NÃO existe em Jan/2026:")
print(f"  → O sistema usa o último preço disponível: R$ 2,96 (Jul/2025)")
print(f"  → Variação: (2,96 - 103,65) / 103,65 = -97,1%")
print(f"\n  O valor R$ 103,65 que aparece como 'Preço Inicial' é a")
print(f"  média de 2025: (160 + 148 + 2,96) / 3 = R$ {media_2025:.2f}")
