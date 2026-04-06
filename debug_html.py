from main import DataHandler, generate_movimentacoes_html

dh = DataHandler()
analise = dh.get_movimentacoes_analise()
html = generate_movimentacoes_html(analise)

# Encontrar e exibir o JavaScript
import re
script_match = re.search(r'<script>(.*?)</script>', html, re.DOTALL)
if script_match:
    js_code = script_match.group(1)
    # Mostrar primeiras 2000 caracteres
    print("=== JAVASCRIPT GERADO ===")
    print(js_code[:2000])
    print("\n=== FIM (primeiros 2000 chars) ===")
else:
    print("Script não encontrado!")
