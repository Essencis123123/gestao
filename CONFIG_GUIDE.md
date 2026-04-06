# Sistema de Configurações

## Visão Geral

O sistema agora possui um gerenciador de configurações que permite alterar credenciais e parâmetros sem precisar modificar o código-fonte.

## Como Usar

### 1. Acessar as Configurações

- Abra o sistema
- Clique no botão **⚙️ Configurações** na barra lateral
- Uma janela de configurações será aberta

### 2. Editar Credenciais do Pipefy

Na janela de configurações, você pode:

- **Editar Client ID**: Digite o novo Client ID fornecido pelo Pipefy
- **Editar Client Secret**: Digite o novo Client Secret (o campo é ocultado por segurança)
- **Mostrar/Ocultar Secret**: Clique no ícone 👁️ para alternar a visibilidade
- **Testar Conexão**: Clique em "🔍 Testar Conexão" para validar as credenciais antes de salvar
- **Salvar**: Clique em "💾 Salvar" para gravar as alterações

### 3. Aplicar Alterações

Após salvar as configurações:
- O sistema exibirá uma mensagem de confirmação
- **Reinicie o sistema** para que as novas credenciais sejam carregadas

## Arquivo de Configuração

As configurações são armazenadas no arquivo `config.json` na pasta raiz do sistema.

### Estrutura do config.json

```json
{
    "pipefy": {
        "client_id": "seu_client_id",
        "client_secret": "seu_client_secret",
        "token_url": "https://app.pipefy.com/oauth/token",
        "api_url": "https://api.pipefy.com/graphql"
    }
}
```

### Segurança

⚠️ **IMPORTANTE**: 
- O arquivo `config.json` contém credenciais sensíveis
- Ele está incluído no `.gitignore` para não ser versionado
- Nunca compartilhe este arquivo publicamente
- Use `config.example.json` como modelo para novos ambientes

## Primeira Configuração

Se é a primeira vez que está usando o sistema:

1. Copie o arquivo `config.example.json` para `config.json`
2. Edite `config.json` com suas credenciais reais
3. OU use a interface gráfica para configurar

## Backup

Recomenda-se fazer backup do arquivo `config.json` em local seguro:
- Antes de atualizar o sistema
- Periodicamente, para recuperação em caso de perda

## Solução de Problemas

### "Erro de autenticação"
- Verifique se as credenciais estão corretas
- Use a opção "Testar Conexão" antes de salvar
- Certifique-se de que o arquivo `config.json` não está corrompido

### "Configurações não aplicadas"
- Reinicie completamente o sistema após salvar
- Verifique as permissões de escrita na pasta

### "Arquivo não encontrado"
- O arquivo `config.json` será criado automaticamente na primeira execução
- Ou copie manualmente de `config.example.json`

## Adicionando Novas Configurações

Para desenvolvedores que desejam adicionar novas configurações:

1. Edite a classe `ConfigManager` em `main.py`
2. Adicione novos campos no método `_load_config()`
3. Crie métodos getter/setter se necessário
4. Atualize a interface `ConfigDialog` para permitir edição
5. Atualize este README com a nova configuração

---

**Versão**: 2.0.0  
**Última atualização**: Fevereiro 2026
