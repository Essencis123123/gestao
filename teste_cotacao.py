"""
Script para extrair e calcular "Tempo total na fase Cotação (dias)"
diretamente da API do Pipefy
"""
import json
import os
import sys
from datetime import datetime, timezone
from collections import defaultdict

# Adicionar o diretório atual ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Configuração
PIPE_ID = '304103091'  # ID do pipe de Contratação de Serviços

def calcular_tempo_fase(card, nome_fase='Cotação'):
    """Calcula o tempo que um card ficou em uma fase específica"""
    phases_history = card.get('phases_history', [])
    
    tempo_total = 0
    for ph in phases_history:
        phase = ph.get('phase', {})
        phase_name = phase.get('name', '')
        
        # Verificar se é a fase desejada (case insensitive)
        if phase_name.lower() == nome_fase.lower():
            first_in = ph.get('firstTimeIn')
            last_out = ph.get('lastTimeOut')
            
            if first_in:
                try:
                    start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                    # Se não saiu ainda, calcular até agora (como o Pipefy faz)
                    if last_out:
                        end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                    else:
                        end = datetime.now(timezone.utc)
                    
                    dias = (end - start).total_seconds() / (3600 * 24)
                    tempo_total += dias
                except:
                    pass
    
    return tempo_total

def extrair_tempos_cotacao():
    """Extrai os valores de tempo na fase Cotação"""
    
    # Primeiro, tentar ler do arquivo JSON salvo
    json_file = 'pipefy_compras_servicos.json'
    
    if os.path.exists(json_file):
        print(f"✅ Arquivo encontrado: {json_file}")
        print(f"📂 Carregando dados salvos...")
        
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                cards = json.load(f)
            print(f"✅ {len(cards)} cards carregados do arquivo!")
        except Exception as e:
            print(f"❌ Erro ao ler arquivo: {e}")
            return
    else:
        print("❌ Arquivo pipefy_compras_servicos.json não encontrado!")
        print("\n📝 INSTRUÇÕES:")
        print("1. Execute o programa principal (main.py)")
        print("2. Aguarde carregar 'Contratação de Serviços'")
        print("3. O arquivo será criado automaticamente")
        print("4. Execute este script novamente")
        return
    
    print(f"\n📊 Calculando tempo na fase Cotação...")
    
    # Calcular tempo na fase Cotação para cada card
    valores = []
    cards_com_cotacao = []
    cards_sem_cotacao = []
    
    for card in cards:
        tempo = calcular_tempo_fase(card, 'Cotação')
        
        if tempo > 0:
            valores.append(tempo)
            cards_com_cotacao.append({
                'id': card.get('id'),
                'titulo': card.get('title', 'Sem título')[:50],
                'tempo': tempo,
                'created': card.get('created_at', ''),
                'finished': card.get('finished_at', '')
            })
        else:
            cards_sem_cotacao.append({
                'id': card.get('id'),
                'titulo': card.get('title', 'Sem título')[:50]
            })
    
    # Mostrar resultados
    print("\n" + "=" * 80)
    print(f"📈 ANÁLISE DOS DADOS")
    print("=" * 80)
    print(f"Total de cards: {len(cards)}")
    print(f"Cards que passaram pela Cotação: {len(cards_com_cotacao)}")
    print(f"Cards que não passaram pela Cotação: {len(cards_sem_cotacao)}")
    
    if not valores:
        print("\n❌ Nenhum card passou pela fase Cotação!")
        return
    
    # Calcular estatísticas
    print("\n" + "=" * 80)
    print("📊 ESTATÍSTICAS - TEMPO NA FASE COTAÇÃO")
    print("=" * 80)
    media = sum(valores) / len(valores)
    print(f"Total de valores: {len(valores)}")
    print(f"Média: {media:.6f} dias")
    print(f"Mínimo: {min(valores):.6f} dias")
    print(f"Máximo: {max(valores):.6f} dias")
    print(f"Soma total: {sum(valores):.2f} dias")
    
    # Comparar com o valor do Excel
    print("\n" + "=" * 80)
    print("🔍 COMPARAÇÃO COM EXCEL")
    print("=" * 80)
    valor_excel = 8.707
    diferenca = abs(media - valor_excel)
    print(f"Valor no Excel (informado): {valor_excel:.6f} dias")
    print(f"Valor calculado pela API:  {media:.6f} dias")
    print(f"Diferença:                 {diferenca:.6f} dias ({(diferenca/valor_excel)*100:.2f}%)")
    
    if diferenca < 0.01:
        print("✅ VALORES IDÊNTICOS!")
    elif diferenca < 0.1:
        print("✅ VALORES MUITO PRÓXIMOS (diferença aceitável)")
    elif diferenca < 1:
        print("⚠️  VALORES PRÓXIMOS (pequena diferença)")
    else:
        print("❌ VALORES DIFERENTES (verificar metodologia)")
    
    # Mostrar detalhes dos cards (primeiros 30)
    print("\n" + "=" * 80)
    print("📋 CARDS COM TEMPO NA COTAÇÃO (primeiros 30)")
    print("=" * 80)
    for i, card in enumerate(cards_com_cotacao[:30], 1):
        print(f"{i:3d}. {card['tempo']:10.6f} dias | {card['titulo']}")
    
    if len(cards_com_cotacao) > 30:
        print(f"\n... e mais {len(cards_com_cotacao) - 30} cards")
    
    # Exportar para CSV
    print("\n" + "=" * 80)
    print("💾 EXPORTANDO VALORES")
    print("=" * 80)
    
    # CSV com todos os valores
    output_file = 'valores_cotacao_api.csv'
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("Card ID,Título,Tempo Cotação (dias),Data Criação,Data Conclusão\n")
        for card in cards_com_cotacao:
            f.write(f"{card['id']},\"{card['titulo']}\",{card['tempo']:.6f},{card['created']},{card['finished']}\n")
    print(f"✅ Dados completos exportados para: {output_file}")
    
    # CSV simples só com valores
    output_file2 = 'valores_cotacao_simples.csv'
    with open(output_file2, 'w', encoding='utf-8') as f:
        f.write("Tempo (dias)\n")
        for v in valores:
            f.write(f"{v:.6f}\n")
    print(f"✅ Valores simples exportados para: {output_file2}")
    
    # Salvar JSON completo
    json_file = 'dados_completos_cotacao.json'
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump({
            'total_cards': len(cards),
            'cards_com_cotacao': len(cards_com_cotacao),
            'media': media,
            'minimo': min(valores),
            'maximo': max(valores),
            'valores': valores,
            'cards': cards_com_cotacao
        }, f, ensure_ascii=False, indent=2)
    print(f"✅ Dados JSON exportados para: {json_file}")

if __name__ == '__main__':
    extrair_tempos_cotacao()
    print("\n✅ Análise concluída!")
    input("\nPressione ENTER para fechar...")
