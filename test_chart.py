from main import DataHandler, generate_movimentacoes_html

dh = DataHandler()
analise = dh.get_movimentacoes_analise()
html = generate_movimentacoes_html(analise)

# Verificar se o container do gráfico existe
print('Container do gráfico existe:', 'chartGruposConsolidado' in html)
print('Canvas existe:', '<canvas id="chartGruposConsolidado">' in html)

# Verificar se o JavaScript está presente
print('JavaScript load event:', 'window.addEventListener' in html)
print('Chart constructor:', 'new Chart' in html)

# Verificar dados
print('gruposLabels length check:', 'gruposLabels.length > 0' in html)
print('Chart.js check:', 'typeof Chart !== \'undefined\'' in html)

# Verificar se os dados estão sendo passados
import re
grupos_labels_match = re.search(r'const gruposLabels = (\[.*?\]);', html, re.DOTALL)
if grupos_labels_match:
    print('gruposLabels encontrado no HTML')
    print('gruposLabels valor:', grupos_labels_match.group(1)[:100])
else:
    print('gruposLabels NÃO encontrado no HTML')