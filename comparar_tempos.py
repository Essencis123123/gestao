"""
Script para comparar tempos calculados pela API vs Excel do Pipefy
"""
import pandas as pd
import json
import os
from datetime import datetime, timezone

def calcular_tempo_fase_api(card, nome_fase='Cotação'):
    """Calcula o tempo que um card ficou em uma fase específica pela API"""
    phases_history = card.get('phases_history', [])
    
    tempo_total = 0
    for ph in phases_history:
        phase = ph.get('phase', {})
        phase_name = phase.get('name', '')
        
        if phase_name.lower() == nome_fase.lower():
            first_in = ph.get('firstTimeIn')
            last_out = ph.get('lastTimeOut')
            
            if first_in and last_out:
                try:
                    start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                    end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                    dias = (end - start).total_seconds() / (3600 * 24)
                    tempo_total += dias
                except:
                    pass
    
    return tempo_total

def main():
    print("=" * 80)
    print("🔍 COMPARAÇÃO: API vs EXCEL DO PIPEFY")
    print("=" * 80)
    
    # 1. Ler Excel do Pipefy
    excel_file = 'Backup/integrao_a_28-01-2026.xlsx'
    
    if not os.path.exists(excel_file):
        print(f"❌ Arquivo Excel não encontrado: {excel_file}")
        return
    
    print(f"\n📊 Lendo Excel do Pipefy...")
    df = pd.read_excel(excel_file)
    print(f"✅ {len(df)} linhas carregadas do Excel")
    
    # Extrair valores da coluna de tempo na Cotação
    coluna_tempo = 'Tempo total na fase Cotação (dias)'
    valores_excel = []
    
    for valor in df[coluna_tempo]:
        if pd.notna(valor):
            try:
                valores_excel.append(float(valor))
            except:
                pass
    
    print(f"✅ {len(valores_excel)} valores válidos no Excel")
    media_excel = sum(valores_excel) / len(valores_excel) if valores_excel else 0
    print(f"📈 Média no Excel: {media_excel:.6f} dias")
    
    # 2. Ler JSON da API
    json_file = 'pipefy_compras_servicos.json'
    
    if not os.path.exists(json_file):
        print(f"\n❌ Arquivo JSON não encontrado: {json_file}")
        print("Execute o programa principal primeiro!")
        return
    
    print(f"\n📊 Lendo dados da API...")
    with open(json_file, 'r', encoding='utf-8') as f:
        cards = json.load(f)
    print(f"✅ {len(cards)} cards carregados da API")
    
    # Calcular tempos pela API
    valores_api = []
    for card in cards:
        tempo = calcular_tempo_fase_api(card, 'Cotação')
        if tempo > 0:
            valores_api.append(tempo)
    
    print(f"✅ {len(valores_api)} valores calculados pela API")
    media_api = sum(valores_api) / len(valores_api) if valores_api else 0
    print(f"📈 Média calculada pela API: {media_api:.6f} dias")
    
    # 3. Comparação
    print("\n" + "=" * 80)
    print("📊 RESULTADO DA COMPARAÇÃO")
    print("=" * 80)
    print(f"Valores no Excel:       {len(valores_excel)}")
    print(f"Valores na API:         {len(valores_api)}")
    print(f"Média Excel:            {media_excel:.6f} dias")
    print(f"Média API:              {media_api:.6f} dias")
    print(f"Diferença:              {abs(media_excel - media_api):.6f} dias")
    print(f"Diferença percentual:   {abs(media_excel - media_api)/media_excel*100:.2f}%")
    
    if abs(media_excel - media_api) < 0.01:
        print("\n✅ VALORES IDÊNTICOS!")
    elif abs(media_excel - media_api) < 0.1:
        print("\n✅ VALORES MUITO PRÓXIMOS!")
    else:
        print("\n⚠️  VALORES DIFERENTES - Analisando...")
        
        # Análise detalhada
        print("\n" + "=" * 80)
        print("🔍 ANÁLISE DETALHADA DAS DIFERENÇAS")
        print("=" * 80)
        
        # Comparar distribuição
        print(f"\nExcel - Min: {min(valores_excel):.6f}, Max: {max(valores_excel):.6f}")
        print(f"API   - Min: {min(valores_api):.6f}, Max: {max(valores_api):.6f}")
        
        # Verificar se há cards com múltiplas passagens pela Cotação
        multiplas_passagens = []
        for card in cards:
            count = 0
            for ph in card.get('phases_history', []):
                if ph.get('phase', {}).get('name', '').lower() == 'cotação':
                    count += 1
            if count > 1:
                multiplas_passagens.append({
                    'titulo': card.get('title', '')[:50],
                    'passagens': count
                })
        
        if multiplas_passagens:
            print(f"\n⚠️  {len(multiplas_passagens)} cards passaram pela Cotação múltiplas vezes:")
            for i, c in enumerate(multiplas_passagens[:10], 1):
                print(f"   {i}. {c['titulo']} - {c['passagens']} vezes")
            print("\n💡 Isso pode explicar diferenças se o Excel soma ou calcula diferente.")
    
    # 4. Exportar comparação
    print("\n" + "=" * 80)
    print("💾 EXPORTANDO COMPARAÇÃO")
    print("=" * 80)
    
    with open('comparacao_excel_vs_api.csv', 'w', encoding='utf-8') as f:
        f.write("Fonte,Quantidade,Média (dias),Mínimo,Máximo\n")
        f.write(f"Excel,{len(valores_excel)},{media_excel:.6f},{min(valores_excel):.6f},{max(valores_excel):.6f}\n")
        f.write(f"API,{len(valores_api)},{media_api:.6f},{min(valores_api):.6f},{max(valores_api):.6f}\n")
    
    print("✅ Comparação exportada: comparacao_excel_vs_api.csv")
    
    # Exportar valores lado a lado
    with open('valores_excel_vs_api.csv', 'w', encoding='utf-8') as f:
        f.write("Excel (dias),API (dias)\n")
        max_len = max(len(valores_excel), len(valores_api))
        for i in range(max_len):
            excel_val = f"{valores_excel[i]:.6f}" if i < len(valores_excel) else ""
            api_val = f"{valores_api[i]:.6f}" if i < len(valores_api) else ""
            f.write(f"{excel_val},{api_val}\n")
    
    print("✅ Valores exportados: valores_excel_vs_api.csv")

if __name__ == '__main__':
    main()
    print("\n✅ Análise concluída!")
    input("\nPressione ENTER para fechar...")
