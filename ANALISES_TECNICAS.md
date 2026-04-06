# 📊 Documentação Técnica - Análises Implementadas

## 1. KPIs Principais

### 1.1 Total Pago
- **Definição**: Soma de todos os valores pagos no período
- **Fórmula**: SUM(Valor Pago)
- **Variação**: Comparativo com período anterior
- **Uso**: Acompanhamento do gasto total

### 1.2 Fornecedores Ativos
- **Definição**: Quantidade de fornecedores únicos
- **Fórmula**: COUNT(DISTINCT Fornecedor)
- **Inclui**: Novos fornecedores do mês
- **Uso**: Diversificação de supply chain

### 1.3 Ticket Médio
- **Definição**: Valor médio por nota fiscal
- **Fórmula**: SUM(Valor) / COUNT(Notas Fiscais)
- **Variação**: Comparativo mensal
- **Uso**: Análise de concentração de volume

### 1.4 Maior Pagamento
- **Definição**: Máximo valor pago em uma transação
- **Fórmula**: MAX(Valor Pago)
- **Acompanhado**: Nome do fornecedor responsável
- **Uso**: Identificar transações especiais

### 1.5 Valor em Estoque
- **Definição**: Soma de (Quantidade × Valor Unitário)
- **Fórmula**: SUM(Qtd × Valor Unit)
- **Atualização**: Tempo real com movimentações
- **Uso**: Capital de giro imobilizado

### 1.6 Itens Críticos
- **Definição**: Quantidade de itens abaixo do mínimo
- **Fórmula**: COUNT WHERE Qtd < Qtd Mín
- **Alerta**: Vermelho no KPI
- **Ação**: Emitir ordem de compra

---

## 2. Curva ABC - Análise Pareto

### 2.1 Princípio
Baseado na Regra de Pareto (80-20):
- **Classe A**: 80% do valor com poucos fornecedores
- **Classe B**: 15% do valor com quantidade intermediária
- **Classe C**: 5% do valor com muitos fornecedores

### 2.2 Cálculo
```
1. Agrupar por fornecedor
2. Somar valores pagos
3. Ordenar decrescente
4. Calcular percentual acumulado
5. Classificar conforme % acumulado
```

### 2.3 Aplicações
- **Classe A**: Negociações prioritárias
- **Classe B**: Acompanhamento regular
- **Classe C**: Padronização

### 2.4 Dados do Dashboard
```
Classe A: 
  - Qtd: Poucos fornecedores
  - Valor: 80% total
  - Ação: Contrato estratégico

Classe B:
  - Qtd: Intermediária
  - Valor: 15% total
  - Ação: Acompanhamento

Classe C:
  - Qtd: Muitos fornecedores
  - Valor: 5% total
  - Ação: Consolidação
```

---

## 3. Tempo de Entrega por Fornecedor

### 3.1 Métricas
- **Prazo Médio**: Dias prometidos em média
- **Atraso Médio**: Dias de atraso em média
- **Qtd Pedidos**: Quantidade de ordens analisadas

### 3.2 Score de Confiabilidade
```
Score = 100 - (Atraso Médio / Prazo Médio × 100)

Interpretação:
- Score ≥ 80: Excelente (🟢)
- 60 ≤ Score < 80: Bom (🟡)
- Score < 60: Inadequado (🔴)
```

### 3.3 Importância
- Afeta cronograma de produção
- Impacta estoque de segurança
- Indica confiabilidade do fornecedor

---

## 4. Análise de Concentração de Risco

### 4.1 Princípio
Mede dependência crítica de fornecedores específicos

### 4.2 Classificação de Risco
```
CRÍTICO: > 30% do gasto
  → Risco de desabastecimento total
  → Necessário plano B urgente

ALTO: 15-30% do gasto
  → Risco significativo
  → Diversificar gradualmente

MÉDIO: 8-15% do gasto
  → Risco moderado
  → Monitorar performance

BAIXO: < 8% do gasto
  → Risco aceitável
  → Manutenção normal
```

### 4.3 Alertas Automáticos
- Notificação quando fornecedor atinge crítico
- Recomendação de ação
- Rastreamento de mitigação

### 4.4 Concentração Top 5
- Percentual que os 5 maiores representam
- Quanto menor, melhor distribuído
- Ideal: < 80%

---

## 5. Evolução Temporal de Gastos

### 5.1 Dados Disponíveis
- Gasto mensal total
- Evolução por fornecedor
- Comparativos período a período

### 5.2 Análise de Sazonalidade
- Picos de gasto
- Períodos de queda
- Padrões recorrentes

### 5.3 Tendência
- Crescimento ou queda
- Taxa de variação
- Previsibilidade

### 5.4 Usos
- Planejamento de caixa
- Previsão orçamentária
- Identificação de anomalias

---

## 6. Ranking de Fornecedores

### 6.1 Campos Analisados
```
Posição: Ordem por volume gasto

Fornecedor: Nome do fornecedor

Total Pago: Soma período

% Total: Percentual do gasto

Qtd NFs: Quantidade de notas

Ticket Médio: Total / Qtd NFs

Classificação: Classe ABC
```

### 6.2 Interatividade
- Clique no fornecedor para detalhes
- Visualização da ficha completa
- Histórico de transações
- Gráficos específicos

---

## 7. Alertas Automáticos do Sistema

### 7.1 Categorias de Alertas

#### Estoque (🔴 CRÍTICO)
- Itens abaixo da quantidade mínima
- Necessidade imediata de compra
- Risco de parada de produção

#### Ordens (🟡 ATENÇÃO)
- Atraso > 5 dias
- Contato necessário com fornecedor
- Risco de cronograma

#### Concentração (🔴 CRÍTICO)
- Fornecedor crítico identificado
- Necessário plano de contingência
- Risco de negócio

### 7.2 Ações Recomendadas
- Alerta inclui ação sugerida
- Clique para detalhes
- Rastreamento de resolução

---

## 8. Gestão de Estoque

### 8.1 Campos
```
Código: Identificação única
Material: Descrição
Qtd Atual: Quantidade presente
Qtd Mínima: Nível de reposição
Qtd Máxima: Limite superior
Valor Unitário: Preço
Valor Total: Qtd × Valor Unit
Localização: Endereço físico
Última Movimentação: Data
Classificação ABC: Categoria
```

### 8.2 Status de Item
```
Normal: Qtd entre Mín e Máx
Atenção: Qtd > Mín mas < 1.5×Mín
Crítico: Qtd < Mín
```

### 8.3 Curva ABC de Estoque
- Mesma lógica de fornecedores
- Aplicado a materiais
- Prioridade de reposição

---

## 9. Ordens de Compra

### 9.1 Status Possíveis
- **Emitida**: Inicial, aguardando aprovação
- **Aprovada**: Confirmada, aguardando fornecedor
- **Em Trânsito**: Saiu do fornecedor
- **Recebida**: Entregue e conferido
- **Cancelada**: Cancelada por algum motivo

### 9.2 Acompanhamento
- Prazo de entrega
- Dias em atraso
- Valor total
- Fornecedor responsável

### 9.3 Métricas
- Total de ordens
- Ordens abertas (ação necessária)
- Ordens atrasadas (urgência)

---

## 10. Movimentações de Estoque

### 10.1 Tipos
- **Entrada**: Recebimento de nota fiscal
- **Saída**: Baixa para produção/consumo

### 10.2 Registro
```
Data: Quando ocorreu
Tipo: Entrada ou Saída
Material: Qual item
Quantidade: Quantas unidades
Documento: NF ou REQ
Responsável: Quem fez
Motivo: Produção, Manutenção, etc
```

### 10.3 Rastreabilidade
- Histórico completo
- Responsabilização
- Auditoria

---

## 11. Análise de Oportunidades

### 11.1 Consolidação de Fornecedores
- Identificar itens com múltiplos fornecedores
- Calcular economia potencial
- Redução de custos administrativos

### 11.2 Variação de Preços
- Detectar flutuações
- Identificar anomalias
- Oportunidade de negociação

### 11.3 Previsão
- Gasto esperado próximo período
- Sazonalidade
- Tendências

---

## 12. Formato de Dados de Entrada

### 12.1 Colunas Esperadas
```
Data do Pagamento: Date (DD/MM/YYYY)
CNPJ do Fornecedor: String (XX.XXX.XXX/XXXX-XX)
Nome do Fornecedor: String
Valor Pago: Decimal (0.00)
Número da Nota Fiscal: String
Centro de Custo: String
Categoria de Material: String
Descrição do Pagamento: String
Forma de Pagamento: String
```

### 12.2 Validações
- Datas válidas
- CNPJ formato correto
- Valores numéricos positivos
- Sem valores nulos críticos

### 12.3 Tratamento
- Duplicatas ignoradas
- Valores negativos (devoluções) inclusos
- Caracteres especiais normalizados

---

## 13. Saídas e Exportações

### 13.1 Formato Excel
```
Abas:
- Pagamentos: Dados brutos + KPIs
- Estoque: Inventário completo
- Ordens: Status de compras
- Movimentações: Histórico
- Resumo: KPIs principais
```

### 13.2 Gráficos Inclusos
- Curva ABC
- Evolução temporal
- Top fornecedores
- Performance entrega

### 13.3 Relatório Executivo
- KPIs principais
- Principais fornecedores
- Alertas críticos
- Oportunidades identificadas

---

## 14. Integração de Dados

### 14.1 Carregamento
- Detecta automaticamente colunas
- Solicita mapeamento se necessário
- Valida dados
- Carrega em memória

### 14.2 Atualização
- Substitui dados anteriores
- Mantém histórico (exportar antes)
- Recalcula todas análises

### 14.3 Performance
- Até 50.000 registros recomendado
- Gráficos renderizam em <1s
- Interface responsiva

---

## 15. Algoritmos Principais

### 15.1 Curva ABC
```python
dados = agrupa_por_fornecedor(pagamentos)
dados = ordena_decrescente(dados)
dados = calcula_percentual_acumulado(dados)
dados = classifica(A se % ≤ 80, B se % ≤ 95, C)
```

### 15.2 Score Entrega
```python
score = 100 - (atraso_médio / prazo_médio * 100)
score = max(0, min(100, score))
```

### 15.3 Risco
```python
risco = CRÍTICO se % > 30
risco = ALTO se % > 15
risco = MÉDIO se % > 8
risco = BAIXO se % ≤ 8
```

---

## 📈 Exemplos de Uso

### Caso 1: Análise ABC
1. Importar dados de pagamentos
2. Visualizar Curva ABC
3. Identificar Classe A (crítica)
4. Negociar termos melhores

### Caso 2: Gestão de Risco
1. Analisar Concentração
2. Encontrar fornecedor crítico (>30%)
3. Desenvolver plano B
4. Monitorar redução de risco

### Caso 3: Otimização de Entrega
1. Analisar Score por fornecedor
2. Identificar baixo desempenho
3. Contatar fornecedor
4. Acompanhar melhoria

### Caso 4: Gestão de Estoque
1. Visualizar itens críticos
2. Emitir ordem de compra
3. Rastrear entrega
4. Registrar recebimento

---

## 🔐 Segurança de Dados

- Dados armazenados em memória durante sessão
- Exportar para Excel para persistência
- Sem transmissão externa
- Uso local

---

**Última Atualização**: Janeiro 2026  
**Versão**: 2.0.0
