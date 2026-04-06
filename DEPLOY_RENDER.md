# 🚀 Guia de Deploy no RENDER

Sua aplicação será automaticamente deployada do GitHub para o Render!

## 📋 Pré-requisitos
- ✅ Repositório GitHub (você já tem em `https://github.com/Essencis123123/gestao`)
- ✅ Conta Render (criar em https://render.com)
- ✅ Arquivos de configuração (já criados!)

## 🎯 Passo 1: Criar Conta no Render

1. Vá para: **https://render.com**
2. Clique em **"Sign Up"**
3. Escolha **"Continue with GitHub"**
4. Autorize o Render a acessar seu GitHub
5. Pronto! Conta criada ✅

## 🔗 Passo 2: Conectar Repositório GitHub

1. Na dashboard do Render, clique em **"New +"** 
2. Selecione **"Web Service"**
3. Em "Connect a repository", escolha **"GitHub"**
4. Busque por **"gestao"**
5. Clique em **"Connect"**

**Configurações automáticas (já estão no arquivo `render.yaml`):**
- ✅ Name: `gestao-web`
- ✅ Environment: `Python`
- ✅ Build Command: `pip install -r requirements-web.txt`
- ✅ Start Command: `gunicorn app:app --bind 0.0.0.0:10000`
- ✅ Python Version: 3.11

## ⚙️ Passo 3: Variáveis de Ambiente (Opcional)

Se precisar de variáveis (API keys, etc):

1. Na página do serviço, vá para **"Environment"**
2. Clique em **"Add Environment Variable"**
3. Exemplo:
   ```
   FLASK_ENV = production
   PORT = 10000
   ```

## 🚀 Passo 4: Deploy Automático

**Pronto!** O Render fará deploy automático quando você fizer `git push`:

```bash
git add .
git commit -m "Atualizações"
git push origin main
```

O Render automaticamente:
1. Detecta o push
2. Clona seu repositório
3. Instala dependências (`requirements-web.txt`)
4. Executa `gunicorn app:app`
5. Sobe seu site online em ~2 minutos ✅

## 📊 Acompanhar Deploy

1. Acesse sua dashboard do Render
2. Clique no seu serviço **"gestao-web"**
3. Vá para **"Logs"** para ver o progresso
4. Quando terminar, você recebe uma URL pública:

```
https://gestao-web-xxxx.onrender.com
```

## 🔄 Atualizar Seu Site

A partir de agora, toda vez que você fazer push no GitHub, o site atualiza automaticamente! 🎉

```bash
# Fazer uma mudança no app.py
# Depois...
git add app.py
git commit -m "Melhoria X"
git push origin main

# Pronto! Site atualizado em 2 minutos
```

## 📁 Arquivos de Configuração Criados

- **`Procfile`** - Define como rodar a aplicação
- **`render.yaml`** - Configurações do Render
- **`requirements-web.txt`** - Dependências (Python packages)
- **`app.py`** (atualizado) - Suporta porta dinâmica

## ❓ Troubleshooting

### "Build failed"
- Verifique logs no Render
- Certifique-se que `requirements-web.txt` tem todas as dependências
- Rode localmente: `pip install -r requirements-web.txt`

### "ModuleNotFoundError"
- Adicione o package em `requirements-web.txt`
- Faça push novamente

### Site não abre
- Espere 2-3 minutos após deploy
- Verifique se há erros em "Logs"
- Teste localmente: `python app.py`

### Dados não carregam
- Verifique se os arquivos Excel estão no repositório
- Ou configure para carregar de uma URL (GitHub Raw)

## 🌍 URL do Seu Site

Após deploy bem-sucedido:
```
https://gestao-web-[seu-codigo].onrender.com
```

Compartilhe essa URL com outros usuários!

## 💰 Preços do Render

| Plano | Custo | Limite |
|-------|-------|--------|
| **Free** | $0/mês | 0.5GB RAM, auto-pause após inatividade |
| **Starter** | $7/mês | 1GB RAM, sempre ativo |
| **Standard** | $12/mês | 2GB RAM, melhor performance |

**Recomendação:** Comece com **Free**, depois upgrade para **Starter** se precisar de mais performance.

---

## ✅ Checklist Final

- [ ] Conta Render criada
- [ ] Repositório GitHub conectado
- [ ] `Procfile` presente
- [ ] `render.yaml` presente
- [ ] `requirements-web.txt` com dependências
- [ ] Deploy iniciado (verde em "Status")
- [ ] URL pública obtida
- [ ] Site acessível

---

**Seu site estará online em minutos!** 🎉

Próximos passos opcionais:
- Adicionar um domínio próprio
- Configurar HTTPS (automático)
- Adicionar suporte a banco de dados (PostgreSQL, MongoDB)

Dúvidas? Consulte: https://render.com/docs
