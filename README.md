# 📊 Painel Inteligente de Gestão v2.0

Um sistema sofisticado e moderno para análise avançada de estoque, fornecedores e ordens de compra, desenvolvido com PyQt6 e HTML/CSS moderno.

## ✨ Características Principais

### 🎯 Dashboard Principal
- **KPIs em Tempo Real**: Total pago, fornecedores ativos, ticket médio, maior pagamento
- **Gráficos Interativos**: Evolução mensal, Curva ABC, ranking de fornecedores
- **Análise de Risco**: Concentração de dependência por fornecedor
- **Alertas Automáticos**: Notificações de estoque crítico e atrasos

### 📦 Gestão de Estoque
- Inventário completo com valores
- Classificação ABC automática
- Itens críticos em destaque
- Histórico de movimentações

### 📋 Ordens de Compra
- Acompanhamento de pedidos em tempo real
- Identificação de atrasos
- Análise de performance de entrega
- Score de confiabilidade por fornecedor

### 🔄 Movimentações
- Registro de entradas e saídas
- Rastreamento por material
- Responsáveis e motivos
- Relatório completo de movimentação

### 📊 Análises Avançadas
- **Curva ABC**: Classificação de fornecedores por volume de gasto
- **Tempo de Entrega**: Performance de cada fornecedor
- **Concentração de Risco**: Análise de dependência crítica
- **Comparativos**: Evolução mensal e anual

### 📥 Importação de Dados
- Suporte a arquivos Excel (.xlsx, .xls)
- Suporte a CSV
- Detecção automática de colunas
- Validação de dados

### 📤 Exportação e Relatórios
- Export para Excel com múltiplas abas
- Gráficos prontos para apresentações
- Relatórios executivos automáticos
- Dados estruturados para análise

## 🚀 Instalação

### Pré-requisitos
- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

### Passos de Instalação

1. **Clone ou Baixe o Repositório**
   ```bash
   cd "Ferramenta de gestão"
   ```

2. **Crie um Ambiente Virtual (Recomendado)**
   ```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # Linux/Mac
   source venv/bin/activate
   ```

3. **Instale as Dependências**
   ```bash
   pip install -r requirements.txt
   ```

## 📖 Como Usar

### Iniciar a Aplicação
```bash
python main.py
```

### Importar Dados
1. Clique em **📥 Importar Dados** na barra lateral
2. Selecione um arquivo Excel ou CSV
3. Os dados serão carregados e o dashboard será atualizado

### Formato de Arquivo Esperado

O arquivo deve conter as seguintes colunas (nomes podem variar):

```
Data do Pagamento | CNPJ | Fornecedor | Valor | Nota Fiscal | Categoria | Centro de Custo
```

### Navegar no Dashboard

- **📈 Dashboard**: Visão geral com KPIs e gráficos
- **📦 Estoque**: Gestão de inventário
- **📋 Ordens de Compra**: Acompanhamento de pedidos
- **🔄 Movimentações**: Histórico de entradas/saídas
- **📄 Relatórios**: Gerar relatórios executivos

### Exportar Dados
1. Clique em **📤 Exportar Excel** na barra lateral
2. Escolha a pasta e o nome do arquivo
3. Seu relatório será gerado com todas as abas

## 📊 Métricas Disponíveis

### KPIs Principais
- Total pago no período
- Quantidade de fornecedores ativos
- Ticket médio
- Maior pagamento
- Variação percentual mensal
- Valor total em estoque
- Itens críticos (abaixo do mínimo)
- Ordens abertas
- Ordens atrasadas

### Análises
- **Curva ABC**: Classifica fornecedores em A (80%), B (15%), C (5%)
- **Tempo de Entrega**: Score de confiabilidade baseado em prazos
- **Concentração de Risco**: Identifica dependências críticas
- **Evolução Temporal**: Tendências de gasto ao longo do tempo

## 🎨 Design e Interface

- **Tema Moderno**: Gradientes e cores sofisticadas
- **Interface Responsiva**: Adapta-se a diferentes resoluções
- **Gráficos Interativos**: Chart.js com animações suaves
- **Tipografia Moderna**: Fonte Inter para melhor legibilidade
- **Acessibilidade**: Contraste adequado e navegação intuitiva

## ⚙️ Configuração

### Customizar Cores
Edite o arquivo `main.py` e procure pela seção `COLORS` para personalizar as cores da aplicação.

### Adicionar Novos Campos
Modifique a classe `DataHandler` para adicionar novos campos de dados e análises.

### Estender Funcionalidades
A arquitetura modular permite adicionar:
- Novos gráficos
- Novas páginas
- Novas análises
- Integrações com bancos de dados

## 📚 Estrutura do Projeto

```
Ferramenta de gestão/
├── main.py                 # Arquivo principal da aplicação
├── requirements.txt        # Dependências do projeto
└── README.md              # Este arquivo
```

## 🔧 Dependências

- **PyQt6**: Framework GUI moderno
- **PyQt6-WebEngine**: Navegador integrado para HTML/CSS
- **pandas**: Processamento de dados
- **numpy**: Cálculos numéricos
- **openpyxl**: Manipulação de arquivos Excel

## 💡 Dicas de Uso

1. **Filtros Globais**: Use os filtros no topo do dashboard para refinar análises
2. **Click em Fornecedores**: Clique em qualquer fornecedor na tabela para ver detalhes
3. **Exportar Regularmente**: Exporte seus dados regularmente para manter backups
4. **Monitorar Alertas**: Preste atenção aos alertas automáticos do sistema

## 🐛 Solução de Problemas

### Erro ao Importar Dados
- Verifique se o arquivo está no formato correto (.xlsx ou .csv)
- Confirme que as colunas existem no arquivo
- Tente reabrir a aplicação

### Gráficos não aparecem
- Verifique sua conexão com a internet (Chart.js é carregado de CDN)
- Limpe o cache do navegador
- Reinicie a aplicação

### Performance lenta
- Reduza a quantidade de dados importados
- Feche outras aplicações pesadas
- Aumente a memória virtual do sistema

## 📝 Notas Importantes

- Os dados são armazenados em memória durante a sessão
- Para persistência, exporte regularmente para Excel
- Recomenda-se backup dos arquivos de origem
- A aplicação gera automaticamente dados de amostra para demonstração

## 🎯 Roadmap Futuro

- [ ] Integração com banco de dados
- [ ] Autenticação de usuários
- [ ] Dashboards personalizáveis
- [ ] Agendamento de relatórios
- [ ] API REST
- [ ] Aplicativo Mobile
- [ ] Previsões com ML

## 📧 Suporte

Para reportar bugs ou sugerir melhorias, entre em contato com a equipe de desenvolvimento.

## 📄 Licença

Desenvolvido para uso empresarial interno.

---

**Desenvolvido com ❤️ para otimizar sua gestão de estoque e fornecedores**

**Versão**: 2.0.0  
**Última Atualização**: Janeiro 2026
