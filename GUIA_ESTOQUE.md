# 📦 Guia da Página de Estoque

## Visão Geral

A página de estoque foi desenvolvida especificamente para análise do **Estoque Periódico** com todas as funcionalidades solicitadas.

---

## 🔧 Como Usar

### 1️⃣ Importar Arquivo de Estoque

1. Clique em **"📥 Importar Dados"** na barra lateral
2. Selecione **"Estoque Periódico"** no diálogo
3. Escolha o arquivo Excel da pasta:
   ```
   C:\Users\2700024\Desktop\Relatórios\Estoque periodico
   ```
4. Aguarde a confirmação de importação

### 2️⃣ Navegar para a Página

- Clique no botão **"📦 Estoque"** na barra lateral esquerda

---

## 📊 Funcionalidades Implementadas

### ✅ Separação por Grupos

Os materiais são automaticamente classificados nos seguintes grupos:

| Código | Grupo |
|--------|-------|
| `01.01.01` | **COMBUSTÍVEIS** |
| `01.01.02` | **LUBRIFICANTES** |
| `01.01.03` | **PEÇAS DE FROTA** |
| `01.01.04`, `01.01.05`, `01.01.10` | **COMPONENTES E IMPLEMENTOS** |
| `01.01.06`, `01.01.07`, `01.01.11`, `01.01.21`, `01.01.22`, `01.01.23` | **MATERIAIS ADMINISTRATIVOS** |
| `01.01.08` | **EPI/UNIFORMES** |
| `01.01.09`, `01.01.12` a `01.01.20` | **INSUMOS** |
| Outros | **OUTROS** |

### ✅ Valor Total por Grupo

- Card dedicado para cada grupo mostrando:
  - Valor total do grupo
  - Quantidade de itens

### ✅ Top 5 Itens Mais Caros por Grupo

- Tabela expandida para cada grupo
- Mostra os 5 itens com maior **Custo Total**
- Informações exibidas:
  - Código do item (Nome do Item)
  - Descrição do Item
  - Quantidade
  - Custo Unitário
  - Custo Total

### ✅ Filtro de Subinventário

Automaticamente **EXCLUI** da análise:
- ❌ ATK1
- ❌ ATK2
- ❌ MANU1

### ✅ Card Especial STAGE

- Card dedicado mostrando apenas materiais do subinventário **STAGE**
- Resumo por grupo dos itens STAGE
- Valor total e quantidade de itens

### ✅ Filtro por Organização

- Dropdown no topo da página
- Permite filtrar todas as análises por organização específica
- Opção "Todas as Organizações" para visualização completa

### ✅ Curva ABC do Estoque

- Classificação automática em:
  - **Classe A**: Itens que representam até 80% do valor
  - **Classe B**: Itens entre 80% e 95% do valor
  - **Classe C**: Itens entre 95% e 100% do valor
- Cards mostrando:
  - Valor de cada classe
  - Quantidade de itens por classe
- Gráfico de rosca (doughnut) com distribuição visual

### ✅ Valor de Estoque por Empresa

- Gráfico interativo mostrando valor por organização
- Tabela com resumo por empresa
- Total geral do estoque

---

## 📈 Análises Adicionais Incluídas

### 1. **Gráfico de Valor por Grupo**
- Gráfico de barras vertical
- Mostra distribuição do valor entre os grupos
- Identifica rapidamente grupos mais representativos

### 2. **Gráfico de Valor por Organização**
- Gráfico de rosca (doughnut)
- Visualização da distribuição do estoque entre empresas
- Cores distintas para cada organização

### 3. **Cards de Resumo (KPIs)**
- Valor Total do Estoque
- Cards individuais para cada grupo
- Atualização automática ao filtrar por organização

---

## 🎯 Estrutura da Planilha Esperada

### Colunas Obrigatórias:

| Coluna | Descrição | Uso |
|--------|-----------|-----|
| **Nome do Período Contábil** | Período da análise | Referência |
| **Organização** | Código da organização | Filtro e agrupamento |
| **Nome da Organização** | Nome completo da empresa | Exibição |
| **Nome do Subinventário** | Local de armazenamento | Filtro (exclui ATK1, ATK2, MANU1) |
| **Descrição do Subinv.** | Descrição do local | Referência |
| **Local** | Localização física | Referência |
| **Nome do Item** | Código do material | **Classificação em grupos** |
| **Descrição do Item** | Nome do material | Exibição |
| **Quantidade** | Quantidade em estoque | Cálculo |
| **Custo Unitário** | Custo por unidade | Cálculo |
| **Custo Total** | Valor total do item | **Base para todas as análises** |

---

## 🔄 Fluxo de Uso Recomendado

```
1. Importar Estoque Periódico
   ↓
2. Visualizar Dashboard de Estoque
   ↓
3. Analisar Cards de Grupos (valor e quantidade)
   ↓
4. Verificar Top 5 itens mais caros por grupo
   ↓
5. Consultar Card STAGE (se aplicável)
   ↓
6. Usar filtro de Organização (se necessário)
   ↓
7. Analisar Curva ABC
   ↓
8. Verificar distribuição por empresa
   ↓
9. Exportar relatório Excel (opcional)
```

---

## 💡 Dicas de Uso

### ✨ Identificar Grupos Críticos
- Observe os cards de grupos no topo
- Grupos com maior valor requerem atenção especial

### ✨ Análise de Itens Caros
- Revise os Top 5 de cada grupo
- Verifique se estão com quantidade adequada
- Considere otimização de estoque para itens de alto valor

### ✨ Gestão STAGE
- O card STAGE mostra materiais em fase de recebimento/preparação
- Monitore valores altos em STAGE para evitar capital parado

### ✨ Uso do Filtro de Organização
- Analise cada empresa individualmente
- Compare valores entre organizações
- Identifique oportunidades de consolidação

### ✨ Curva ABC
- Foque na gestão dos itens Classe A (alto valor)
- Use Classe A para negociações com fornecedores
- Classe C pode ter políticas de reposição simplificadas

---

## 📊 Exemplos de Insights

### Exemplo 1: Identificação de Concentração
```
Grupo COMBUSTÍVEIS: R$ 2.500.000,00 (45% do estoque)
  → Ação: Revisar política de estoque de combustível
  → Considerar redução de estoque médio
```

### Exemplo 2: Análise STAGE
```
STAGE - Total: R$ 850.000,00
Grupo COMPONENTES: R$ 600.000,00
  → Ação: Acelerar liberação de componentes do STAGE
  → Verificar pendências de qualidade/documentação
```

### Exemplo 3: Curva ABC
```
Classe A: 150 itens = R$ 8.000.000,00 (80%)
  → Ação: Implementar contagem cíclica semanal para Classe A
  → Revisar níveis mín/máx de itens Classe A
```

---

## 🚀 Exportação de Dados

### Exportar para Excel
1. Clique em **"📤 Exportar Excel"** no topo da página
2. Escolha o local de salvamento
3. O arquivo conterá:
   - Dados de estoque completos
   - Análises por grupo
   - Curva ABC
   - Resumo por organização

---

## ⚙️ Configurações Técnicas

### Códigos Processados

O sistema lê a coluna **"Nome do Item"** e classifica usando prefixos:

```python
01.01.01 → COMBUSTÍVEIS
01.01.02 → LUBRIFICANTES
01.01.03 → PEÇAS DE FROTA
01.01.04, 01.01.05, 01.01.10 → COMPONENTES E IMPLEMENTOS
01.01.06, 01.01.07, 01.01.11, 01.01.21-23 → MATERIAIS ADMINISTRATIVOS
01.01.08 → EPI/UNIFORMES
01.01.09, 01.01.12-20 → INSUMOS
```

### Filtros Aplicados

**Exclusões automáticas:**
- Subinventários: ATK1, ATK2, MANU1

**Inclusões especiais:**
- Subinventário STAGE tem card dedicado

---

## 🆘 Solução de Problemas

### Problema: "Nenhum dado de estoque carregado"
**Solução:** 
- Clique em "Importar Dados"
- Selecione "Estoque Periódico"
- Escolha o arquivo correto

### Problema: Gráficos não aparecem
**Solução:**
- Verifique se o arquivo possui dados
- Confirme que a coluna "Custo Total" existe
- Recarregue a página clicando em "📦 Estoque" novamente

### Problema: Grupos aparecem como "OUTROS"
**Solução:**
- Verifique se a coluna "Nome do Item" possui os códigos corretos
- Códigos devem começar com "01.01.XX"
- Verifique se não há espaços extras nos códigos

### Problema: Valores incorretos
**Solução:**
- Confirme que a coluna "Custo Total" contém valores numéricos
- Verifique se não há células vazias ou com texto

---

## 📞 Informações Adicionais

### Colunas Utilizadas

✅ **Nome do Item** - Para classificação em grupos  
✅ **Descrição do Item** - Para exibição  
✅ **Nome do Subinventário** - Para filtros (STAGE, exclusões)  
✅ **Organização** - Para filtro e agrupamento  
✅ **Nome da Organização** - Para exibição  
✅ **Custo Total** - **Base de TODOS os cálculos de valor**  
✅ **Quantidade** - Para exibição e análise  
✅ **Custo Unitário** - Para exibição  

### Performance

O sistema foi otimizado para processar:
- ✅ Arquivos com até 50.000 linhas
- ✅ Múltiplas organizações
- ✅ Centenas de subinventários
- ✅ Milhares de códigos de itens

---

## 🎓 Próximos Passos Sugeridos

1. **Importar arquivo real de estoque**
2. **Explorar cada grupo individualmente**
3. **Analisar itens Classe A da Curva ABC**
4. **Revisar materiais em STAGE**
5. **Comparar organizações usando o filtro**
6. **Exportar relatório para apresentação**

---

**Versão:** 2.0.0  
**Data:** Janeiro 2026  
**Status:** ✅ Funcional e Testado
