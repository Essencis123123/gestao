# 🚀 Guia de Início Rápido

## ⚡ 3 Passos para Começar

### 1️⃣ Instalar Dependências
```bash
pip install -r requirements.txt
```

Ou use o script automático (Windows):
```bash
run.bat
```

### 2️⃣ Executar a Aplicação
```bash
python main.py
```

### 3️⃣ Explorar o Dashboard
- A aplicação abre com dados de demonstração automáticos
- Todos os gráficos e análises já estão disponíveis
- Importe seus próprios dados para análises reais

---

## 📋 Funcionalidades Principais Visíveis no Dashboard

### Dashboard Principal
✅ **8 KPIs Importantes**
- Total Pago (com variação mensal)
- Fornecedores Ativos
- Ticket Médio
- Maior Pagamento
- Valor em Estoque
- Itens Críticos
- Ordens Abertas
- Ordens Atrasadas

✅ **4 Gráficos Interativos**
- Evolução Mensal (Linha)
- Curva ABC (Rosca)
- Top 10 Fornecedores (Barras Horizontais)
- Performance de Entrega (Barras)

✅ **Seções de Análise**
- Alertas do Sistema
- Concentração de Risco
- Ranking de Fornecedores (Tabela Interativa)

### Navegação Lateral
- 📈 Dashboard - Visão geral
- 📦 Estoque - Gestão de inventário
- 📋 Ordens - Acompanhamento
- 🔄 Movimentações - Entradas/Saídas
- 📄 Relatórios - Exportações

### Botões de Ação
- 📥 Importar Dados - Carrega arquivos Excel/CSV
- 📤 Exportar Excel - Salva relatório completo
- 🔄 Atualizar - Recarrega dados

---

## 📊 Dados de Demonstração Inclusos

A aplicação vem com **dados realistas de exemplo**:

- **500 Registros de Pagamento** (Jan/2024 - Jan/2026)
- **15 Fornecedores** com distribuição realista
- **17 Materiais** em estoque
- **100 Ordens de Compra** em vários status
- **200 Movimentações** de entrada/saída

Esses dados permitem:
- Explorar todas as funcionalidades
- Ver gráficos em ação
- Testar exportações
- Validar a estrutura

---

## 🎨 Destaques Visuais

### Design Moderno
- Gradientes profissionais
- Ícones Font Awesome
- Animações suaves
- Tipografia moderna (Inter)

### Cores Sofisticadas
- 🔵 Azul Primário: #667eea
- 🟣 Roxo Secundário: #764ba2
- 🟢 Verde Sucesso: #48bb78
- 🟠 Laranja Aviso: #ed8936
- 🔴 Vermelho Perigo: #f56565

### Responsividade
- Desktop (1920px+)
- Laptop (1400px)
- Tablet (768px)
- Adapta-se automaticamente

---

## 💡 Dicas Rápidas

### Clicar em Fornecedores
Na tabela de ranking, clique em qualquer fornecedor para ver:
- Histórico completo de pagamentos
- Análise de padrões
- Gráficos específicos
- Distribuição por categoria

### Usar Filtros
No topo do dashboard:
1. Selecione período (Mês, Trimestre, Semestre, Ano)
2. Filtre por Centro de Custo
3. Escolha Categoria de Material

### Exportar Dados
- Clique em "📤 Exportar Excel"
- Escolha local de salvamento
- Arquivo com múltiplas abas será gerado

---

## 🔧 Requisitos Mínimos

| Requisito | Mínimo | Recomendado |
|-----------|--------|-------------|
| Python | 3.8 | 3.10+ |
| RAM | 2GB | 4GB+ |
| Disco | 500MB | 1GB+ |
| Conexão | Nenhuma | Boa (CDN) |

---

## ✅ Checklist de Instalação

- [ ] Python 3.8+ instalado
- [ ] requirements.txt no diretório
- [ ] Ambiente virtual criado (opcional)
- [ ] Dependências instaladas
- [ ] main.py no diretório
- [ ] Aplicação inicia sem erros
- [ ] Dashboard carrega dados
- [ ] Gráficos exibem corretamente

---

## 🎯 Próximos Passos

1. **Explorar o Dashboard**
   - Clique nos diferentes módulos
   - Teste os filtros
   - Interaja com os gráficos

2. **Importar Seus Dados**
   - Prepare arquivo Excel com seus dados
   - Use as colunas conforme especificado
   - Clique em Importar Dados

3. **Exportar Análises**
   - Gere relatórios executivos
   - Compartilhe com stakeholders
   - Use em apresentações

4. **Personalizar** (Avançado)
   - Edite `main.py` para cores
   - Adicione novos campos
   - Crie novas análises

---

## ❓ Dúvidas Frequentes

**P: Posso usar meus dados reais?**
R: Sim! Importe seu arquivo Excel com os dados em colunas apropriadas.

**P: Preciso de conexão com internet?**
R: Apenas para carregar fontes e Chart.js (pode funcionar offline com ajustes).

**P: Como faço backup dos dados?**
R: Exporte para Excel regularmente usando o botão "Exportar Excel".

**P: Posso modificar o código?**
R: Sim! O código está bem documentado e modular.

**P: Qual navegador é usado?**
R: PyQt6 WebEngine (baseado em Chromium).

---

## 📞 Suporte

Para problemas:
1. Verifique se todas as dependências estão instaladas
2. Confirme que é Python 3.8+
3. Verifique espaço em disco
4. Tente reinstalar dependências

---

**Pronto para começar? Execute `python main.py` ou clique em `run.bat` (Windows)!**

🎉 Aproveite seu novo painel de gestão!
