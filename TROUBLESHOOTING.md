# 🔧 Troubleshooting - Erros Comuns de Deploy

## ❌ Erro: "pandas 2.1.4 - metadata-generation-failed"

### Problema
```
× Encountered error while generating package metadata.
╰─> pandas
error: metadata-generation-failed
```

### Causa
- Render usa Python 3.14 (versão mais recente)
- pandas 2.1.4 não tem wheels (versão compilada) para Python 3.14
- Pip tenta compilar do source, mas falha com incompatibilidade da API interna

### ✅ Solução (JÁ IMPLEMENTADA)

**Atualizar `requirements-web.txt`:**
```
flask==3.0.2           # ✅ Compatível com Python 3.14
pandas==2.2.3          # ✅ Tem wheels para Python 3.14
openpyxl==3.1.2        # ✅ Atualizado
gunicorn==23.0.0       # ✅ Versão estável
requests==2.32.0       # ✅ Atualizado
```

### 🚀 Como Atualizar

Se você teve este erro:

1. **Abra `requirements-web.txt`** no seu editor
2. **Atualize as versões** (cópia acima)
3. **Faça commit:**
   ```bash
   git add requirements-web.txt
   git commit -m "Atualiza dependências para Python 3.14"
   git push origin main
   ```
4. **No Render:**
   - Vá para seu serviço
   - Clique em **"Manual Deploy"** (ou espere próximo push)
   - Deploy vai rodar novamente com as deps certas ✅

---

## ❌ Erro: "ModuleNotFoundError: No module named 'flask'"

### Problema
App inicia mas mostra erro de import

### Solução
```bash
1. Certifique-se que requirements-web.txt tem "flask==3.0.2"
2. Faça git push
3. Render reinstala dependências automaticamente
```

---

## ❌ Erro: "Build failed - pip command not found"

### Solução
Render usa Python 3.11+ automaticamente. Nenhuma ação necessária.

---

## ✅ Verificar Depois do Deploy

No Render dashboard, você deve ver:

```
✓ Build succeeded
✓ Deploy succeeded
✓ Available at: https://gestao-web-xxxxx.onrender.com
```

Se vir verde em tudo, seu site está online! 🎉

---

## 📊 Versões Compatíveis

| Pacote | Versão | Python 3.14 | Observação |
|--------|--------|-----------|-----------|
| Flask | 3.0.2 | ✅ | Estável |
| Pandas | 2.2.3 | ✅ | Tem wheels |
| Openpyxl | 3.1.2 | ✅ | Compatível |
| Gunicorn | 23.0.0 | ✅ | Testado |
| Requests | 2.32.0 | ✅ | Seguro |

---

## 🆘 Ainda Não Funciona?

Se o deploy ainda falhar:

1. **Verifique logs no Render:**
   - Dashboard → Seu service → "Logs"
   - Procure pela linha de erro específica

2. **Tente reduzir para o mínimo:**
   ```
   flask==3.0.2
   pandas==2.2.3
   openpyxl==3.1.2
   gunicorn==23.0.0
   ```

3. **Se for erro de compilação C:**
   - Use versões pré-compiladas (wheels)
   - Evite versões muito antigas de pandas

4. **Teste localmente primeiro:**
   ```bash
   pip install -r requirements-web.txt
   python app.py
   ```

---

## 📚 Referências

- Render Docs: https://render.com/docs
- PyPI - Pandas: https://pypi.org/project/pandas/
- Verificar wheels: https://pypi.org/project/pandas/#files

---

**Seu deploy deve funcionar agora! 🚀**

Se tiver outro erro, me avisa o código que aparece nos logs!
