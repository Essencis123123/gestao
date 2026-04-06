# 🌐 Versão Web - Ferramenta de Gestão

Seu sistema agora está disponível como **site web** usando Flask!

## 🚀 Como Usar

### 1. **Instalar Dependências**
```bash
pip install flask pandas openpyxl
```

### 2. **Rodar o Servidor**
```bash
python app.py
```

### 3. **Acessar no Navegador**
```
http://localhost:5000
```

## 📊 Funcionalidades

### Dashboard
- **KPIs em tempo real** com estatísticas gerais
- Vision completa do sistema
- Informações de carga dos dados

### Seções de Dados
- **📦 Materiais** - Consulte o catálogo com busca
- **📈 Estoque** - Visualize estado do estoque
- **🔧 Serviços** - Filtrado por tipo (Direto, Indireto, Despesas)

### Recursos Avançados
- ✅ Busca em tempo real (live search)
- ✅ Filtros por tipo
- ✅ Design responsivo
- ✅ API REST para integração
- ✅ Relógio em tempo real
- ✅ Layout moderno com tema profissional

## 🔗 API Endpoints

```javascript
GET  /api/stats            → Estatísticas gerais
GET  /api/materiais        → Lista de materiais (com ?search=term)
GET  /api/estoque          → Dados de estoque (com ?search=term)
GET  /api/servicos         → Serviços (com ?tipo=direto|indireto|despesas)
```

**Exemplo de busca:**
```
http://localhost:5000/api/materiais?search=combustivel&limit=50
```

## 📁 Arquivos Necessários

A aplicação carrega dados dos seguintes arquivos Excel:
- `codigos.xlsx` - Materiais
- `Solvi - Estoque Periodico.xlsx` - Estoque
- `Lista de Códigos - Itens direto.xlsx` - Serviços Direto
- `Lista de Códigos - Itens indireto.xlsx` - Serviços Indireto
- `Lista de Códigos - Itens despesas.xlsx` - Serviços Despesas

## 🎨 Tema Visual

Design baseado no `anywhere.py` com:
- **Paleta de cores profissional** (verde escuro, dourado, tons terrosos)
- **Sidebar navegação** com ícones
- **Tabelas inteligentes** com hover
- **Cards informativos** com KPIs
- **Animações suaves** de transição

## 🌍 Deploy Online (Próximos Passos)

Para colocar online, você pode usar:

1. **Heroku** (grátis na tier gratuita, agora com limitações)
2. **PythonAnywhere** (host Python dedicado)
3. **AWS/GCP/Azure** (cloud providers)
4. **Railway** ou **Render** (alternativas modernas)

### Deploy no PythonAnywhere (Recomendado):
1. Crie conta em https://www.pythonanywhere.com
2. Faça upload dos arquivos
3. Configure a aplicação Flask
4. Seu site estará online em minutos!

## ⚙️ Configurações

### Modo Debug
Edite a última linha do `app.py`:
```python
app.run(debug=True, host='0.0.0.0', port=5000)
```

- `debug=True` → Recarrega automático (desenvolvimento)
- `debug=False` → Modo produção (seguro)
- `host='0.0.0.0'` → Acessível de qualquer máquina

### Porta Customizada
Se a porta 5000 estiver em uso:
```python
app.run(debug=False, host='0.0.0.0', port=8080)
```

## 🔒 Segurança

Para produção, adicione:
```python
app.config['ENV'] = 'production'
app.config['DEBUG'] = False
app.secret_key = 'sua-chave-secreta-aqui'
```

## 📞 Suporte

Se encontrar erros:
1. Verifique se os arquivos Excel estão na mesma pasta
2. Confira os nomes exatos dos arquivos
3. Verifique a versão do Python (3.8+)
4. Reinstale as dependências: `pip install --upgrade flask pandas`

---

**Sua ferramenta está pronta para ser usada online!** 🎉
