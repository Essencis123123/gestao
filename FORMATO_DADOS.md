# 📋 Estrutura de Dados para Importação

## Formato Esperado para Arquivo Excel

Seu arquivo Excel deve conter as seguintes colunas. O nome das colunas pode variar ligeiramente, mas a ordem recomendada é:

### Colunas Obrigatórias

| Coluna | Tipo | Formato | Exemplo | Descrição |
|--------|------|---------|---------|-----------|
| Data do Pagamento | Data | DD/MM/YYYY | 15/01/2026 | Quando o pagamento foi feito |
| CNPJ do Fornecedor | Texto | XX.XXX.XXX/XXXX-XX | 12.345.678/0001-90 | CNPJ do fornecedor |
| Nome do Fornecedor | Texto | - | ABC Indústria LTDA | Razão social do fornecedor |
| Valor Pago | Número | 0.00 | 15.500,00 | Valor do pagamento |
| Número da Nota Fiscal | Texto | - | NF-12345 | Identificação da NF |

### Colunas Opcionais (Recomendadas)

| Coluna | Tipo | Formato | Exemplo | Descrição |
|--------|------|---------|---------|-----------|
| Centro de Custo | Texto | - | Produção | Alocação de custo |
| Categoria de Material | Texto | - | Matéria-Prima | Tipo de compra |
| Descrição do Pagamento | Texto | - | Material para produção | Detalhes da compra |
| Forma de Pagamento | Texto | - | Boleto | Método pagamento |

---

## Exemplo de Arquivo CSV

```csv
Data do Pagamento,CNPJ do Fornecedor,Nome do Fornecedor,Valor Pago,Número da Nota Fiscal,Centro de Custo,Categoria de Material,Descrição do Pagamento,Forma de Pagamento
15/01/2026,12.345.678/0001-90,ABC Indústria LTDA,15500.00,NF-001,Produção,Matéria-Prima,Chapas de aço 5mm,Transferência
14/01/2026,23.456.789/0001-01,XYZ Comércio,8750.50,NF-002,Manutenção,Componentes,Parafusos sortidos,Boleto
13/01/2026,34.567.890/0001-12,Fornecedor Nacional,22300.00,NF-003,Produção,Equipamentos,Motor elétrico 3hp,PIX
12/01/2026,12.345.678/0001-90,ABC Indústria LTDA,5600.00,NF-004,Administrativo,Consumíveis,Óleo hidráulico,Cheque
```

---

## Exemplo de Arquivo Excel (Formato)

### Aba: Pagamentos

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| Data do Pagamento | CNPJ do Fornecedor | Nome do Fornecedor | Valor Pago | Número da Nota Fiscal | Centro de Custo | Categoria de Material | Descrição | Forma Pagamento |
| 15/01/2026 | 12.345.678/0001-90 | ABC Indústria LTDA | 15500.00 | NF-001 | Produção | Matéria-Prima | Chapas aço | Transferência |
| 14/01/2026 | 23.456.789/0001-01 | XYZ Comércio | 8750.50 | NF-002 | Manutenção | Componentes | Parafusos | Boleto |

---

## Validações Importantes

### Data do Pagamento
- ✅ Aceito: 15/01/2026, 2026-01-15, 01/15/2026
- ❌ Rejeito: 15-01-2026, 2026/01/15
- **Tipo**: Deve ser reconhecível como data

### CNPJ
- ✅ Aceito: 12.345.678/0001-90, 12345678000190
- ❌ Rejeito: 123456780001-90 (formato inválido)
- **Tipo**: 14 dígitos totais

### Valor
- ✅ Aceito: 15500.00, 15500,00, 15500
- ❌ Rejeito: R$ 15500.00, 15.500,00 com símbolo
- **Tipo**: Número positivo

### Fornecedor e CNPJ
- **Combinação**: Mesmo fornecedor deve ter CNPJ consistente
- **Duplicatas**: Mesmos pagamentos NÃO devem ser repetidos

---

## Alternativas de Formato

### Excel com Aba Nomeada
A aplicação detecta automaticamente e importa da primeira aba,
mas você pode renomear para:
- Pagamentos
- Dados
- Transações
- Movimentações

### CSV com Diferentes Delimitadores
- Ponto-vírgula: `Data;CNPJ;Valor`
- Vírgula: `Data,CNPJ,Valor`
- Tabulação: `Data    CNPJ    Valor`

---

## Mapeamento de Colunas Automático

A aplicação reconhece automaticamente variações de nomes:

| Nome Detectado | Mapeado Para | Exemplos |
|----------------|--------------|----------|
| Data, data_pagamento, data_compra, Data do Pagamento | Data do Pagamento | Data, data, DATA |
| CNPJ, cnpj_fornecedor, cnpj, Fornecedor CNPJ | CNPJ do Fornecedor | CNPJ, cnpj |
| Fornecedor, nome_fornecedor, Razão Social | Nome do Fornecedor | Fornecedor, fornecedor |
| Valor, valor_pago, Valor Pago, total | Valor Pago | Valor, total |
| NF, nota_fiscal, NF, Número NF | Número da Nota Fiscal | NF, nf |

---

## Exemplo Completo para Teste

### Arquivo: `exemplo_dados.xlsx`

```
Data do Pagamento,CNPJ do Fornecedor,Nome do Fornecedor,Valor Pago,Número da Nota Fiscal,Centro de Custo,Categoria de Material
01/01/2025,12.345.678/0001-90,ABC Indústria LTDA,5000.00,NF-2501,Produção,Matéria-Prima
02/01/2025,23.456.789/0001-01,XYZ Comércio,3500.00,NF-2502,Manutenção,Componentes
03/01/2025,34.567.890/0001-12,Fornecedor Nacional,8000.00,NF-2503,Produção,Equipamentos
05/01/2025,12.345.678/0001-90,ABC Indústria LTDA,4500.00,NF-2504,Administrativo,Consumíveis
10/01/2025,45.678.901/0001-23,Tech Solutions,6500.00,NF-2505,Produção,Serviços
15/01/2025,23.456.789/0001-01,XYZ Comércio,7200.00,NF-2506,Manutenção,Matéria-Prima
20/01/2025,56.789.012/0001-34,Material Express,9500.00,NF-2507,Logística,Componentes
25/01/2025,12.345.678/0001-90,ABC Indústria LTDA,5800.00,NF-2508,Produção,Equipamentos
```

---

## Checklist para Preparar Seus Dados

- [ ] Arquivo no formato Excel (.xlsx) ou CSV
- [ ] Coluna de data (reconhecível como data)
- [ ] Coluna de CNPJ ou código de fornecedor
- [ ] Coluna de nome do fornecedor
- [ ] Coluna de valor (numérico)
- [ ] Coluna de nota fiscal ou documento
- [ ] Sem linhas vazias no meio dos dados
- [ ] Primeira linha com nomes de colunas
- [ ] Dados consistentes (sem mistura de formatos)
- [ ] Sem valores duplicados conhecidos
- [ ] Valores em reais (R$) sem símbolo

---

## Troubleshooting

### Erro: "Coluna não encontrada"
- Verifique se o nome da coluna está correto
- A aplicação faz busca case-insensitive
- Tente adicionar/remover espaços

### Erro: "Formato de data inválido"
- Use DD/MM/YYYY ou YYYY-MM-DD
- Não misture formatos diferentes
- Evite abreviações (jan, Jan, JAN)

### Erro: "Valor não é número"
- Remova símbolo de moeda (R$)
- Use . ou , como separador decimal (ambos aceitos)
- Sem separadores de milhar se usar decimal

### Dados não aparecem
- Verifique se o arquivo não está aberto em outro programa
- Confirme que há dados além do cabeçalho
- Tente reabrir a aplicação

---

## Performance

| Quantidade de Registros | Tempo de Carregamento | Recomendação |
|-------------------------|----------------------|--------------|
| < 1.000 | < 0.5s | Ideal |
| 1.000 - 10.000 | 0.5s - 2s | Bom |
| 10.000 - 50.000 | 2s - 10s | Aceitável |
| 50.000+ | > 10s | Considere filtros |

---

## Dicas de Boas Práticas

1. **Consistência**: Use mesmos nomes de fornecedores e centros de custo
2. **Formato**: Mantenha padrão de datas e valores em todo arquivo
3. **Limpeza**: Remova dados duplicados antes de importar
4. **Backup**: Sempre guarde cópia do arquivo original
5. **Versão**: Nomeie arquivos com data (dados_2026_01_15.xlsx)

---

## Exemplo de Estrutura de Dados de Estoque

Para análise completa, você também pode ter:

```csv
Código,Material,Quantidade Atual,Quantidade Mínima,Quantidade Máxima,Valor Unitário,Localização,Classificação ABC
MAT-001,Parafuso M8,250,50,500,5.00,Corredor 1,C
MAT-002,Motor Elétrico,45,10,100,1500.00,Corredor 3,A
MAT-003,Óleo Hidráulico,120,30,250,45.00,Corredor 2,B
```

---

## Estrutura de Ordens de Compra

```csv
Número OC,Data Emissão,Fornecedor,CNPJ,Valor Total,Status,Prazo Entrega,Dias em Atraso
OC-001,15/01/2026,ABC Indústria,12.345.678/0001-90,25000.00,Recebida,15,0
OC-002,14/01/2026,XYZ Comércio,23.456.789/0001-01,18500.00,Em Trânsito,10,5
OC-003,10/01/2026,Fornecedor Nacional,34.567.890/0001-12,35000.00,Emitida,30,0
```

---

**Pronto para importar seus dados? Use a estrutura acima e clique em "📥 Importar Dados"!**
