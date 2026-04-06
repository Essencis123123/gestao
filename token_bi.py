"""
QUERIES GRAPHQL ÚTEIS PARA PIPEFY NO POWER BI

Após executar o script de teste, use estas queries no Power BI
para extrair diferentes tipos de dados do Pipefy.
"""

# ============================================================
# 1. LISTAR TODAS AS ORGANIZAÇÕES
# ============================================================
query_organizacoes = """
{
    organizations {
        id
        name
    }
}
"""

# ============================================================
# 2. LISTAR PIPES DE UMA ORGANIZAÇÃO
# Substitua ORG_ID pelo ID da sua organização
# ============================================================
query_pipes_organizacao = """
{
    organization(id: "ORG_ID") {
        id
        name
        pipes {
            id
            name
        }
    }
}
"""

# ============================================================
# 2. BUSCAR CARDS DE UM PIPE ESPECÍFICO
# Substitua PIPE_ID pelo ID do seu pipe
# ============================================================
query_cards_pipe = """
{
    allCards(pipeId: "PIPE_ID", first: 50) {
        edges {
            node {
                id
                title
                createdAt
                current_phase {
                    name
                }
                assignees {
                    id
                    name
                }
                fields {
                    name
                    value
                }
            }
        }
        pageInfo {
            endCursor
            hasNextPage
        }
    }
}
"""

# ============================================================
# 3. BUSCAR CARDS COM PAGINAÇÃO (PARA PIPES GRANDES)
# ============================================================
query_cards_paginacao = """
{
    pipe(id: "PIPE_ID") {
        id
        name
        cards(first: 50, after: "CURSOR_AQUI") {
            edges {
                node {
                    id
                    title
                    createdAt
                    finished_at
                }
            }
            pageInfo {
                endCursor
                hasNextPage
            }
        }
    }
}
"""

# ============================================================
# 4. BUSCAR INFORMAÇÕES DETALHADAS DE UM CARD
# ============================================================
query_card_detalhado = """
{
    card(id: "CARD_ID") {
        id
        title
        createdAt
        updated_at
        finished_at
        current_phase {
            id
            name
        }
        assignees {
            id
            name
            email
        }
        fields {
            name
            value
            filled_at
        }
        comments {
            text
            author {
                name
            }
            created_at
        }
    }
}
"""

# ============================================================
# 5. BUSCAR ESTRUTURA DO PIPE (FASES E CAMPOS)
# ============================================================
query_estrutura_pipe = """
{
    pipe(id: "PIPE_ID") {
        id
        name
        phases {
            id
            name
            fields {
                id
                label
                type
                description
                required
                options
            }
        }
    }
}
"""

# ============================================================
# 6. BUSCAR DADOS DE TABELA (DATABASE)
# Substitua TABLE_ID pelo ID da sua tabela
# ============================================================
query_table_records = """
{
    table(id: "TABLE_ID") {
        id
        name
        table_records(first: 50) {
            edges {
                node {
                    id
                    title
                    created_at
                    record_fields {
                        name
                        value
                    }
                }
            }
            pageInfo {
                endCursor
                hasNextPage
            }
        }
    }
}
"""

# ============================================================
# 7. BUSCAR RELATÓRIO DE CARDS POR FASE
# ============================================================
query_relatorio_fases = """
{
    pipe(id: "PIPE_ID") {
        id
        name
        phases {
            id
            name
            cards_count
        }
    }
}
"""

# ============================================================
# 8. BUSCAR CARDS CRIADOS EM UM PERÍODO ESPECÍFICO
# ============================================================
query_cards_periodo = """
{
    pipe(id: "PIPE_ID") {
        cards(
            first: 100
            search: {
                created_at: {
                    from: "2024-01-01"
                    to: "2024-12-31"
                }
            }
        ) {
            edges {
                node {
                    id
                    title
                    createdAt
                    current_phase {
                        name
                    }
                }
            }
        }
    }
}
"""

# ============================================================
# EXEMPLO DE CÓDIGO POWER QUERY M COMPLETO
# ============================================================
powerquery_exemplo = """
let
    // ========== CONFIGURAÇÕES ==========
    client_id = "ofgUSnXFhXadEzrDd_ZtUzXsV8-Crv-0NFboRn0CbrU",
    client_secret = "DyLPVER8t6SIeVpO7lQiDqTzoquM3UqLDOUgOFtHFpw",
    token_url = "https://app.pipefy.com/oauth/token",
    api_url = "https://api.pipefy.com/graphql",
    pipe_id = "SEU_PIPE_ID_AQUI",  // Substitua pelo ID do seu pipe
    
    // ========== OBTER TOKEN ==========
    TokenResponse = Json.Document(
        Web.Contents(token_url, [
            Headers = [#"Content-Type"="application/x-www-form-urlencoded"],
            Content = Text.ToBinary(
                "grant_type=client_credentials" &
                "&client_id=" & client_id &
                "&client_secret=" & client_secret
            )
        ])
    ),
    access_token = TokenResponse[access_token],
    
    // ========== QUERY GRAPHQL ==========
    GraphQLQuery = "{
        pipe(id: \"" & pipe_id & "\") {
            id
            name
            cards(first: 100) {
                edges {
                    node {
                        id
                        title
                        createdAt
                        current_phase {
                            name
                        }
                        fields {
                            name
                            value
                        }
                    }
                }
            }
        }
    }",
    
    // ========== FAZER REQUISIÇÃO ==========
    Response = Json.Document(
        Web.Contents(api_url, [
            Headers = [
                #"Authorization" = "Bearer " & access_token,
                #"Content-Type" = "application/json"
            ],
            Content = Text.ToBinary(
                "{\"query\": \"" & Text.Replace(Text.Replace(GraphQLQuery, "#(lf)", " "), "#(cr)", "") & "\"}"
            )
        ])
    ),
    
    // ========== PROCESSAR DADOS ==========
    Data = Response[data][pipe][cards][edges],
    
    // Converter para tabela
    ConvertToTable = Table.FromList(Data, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ExpandColumn = Table.ExpandRecordColumn(ConvertToTable, "Column1", {"node"}, {"node"}),
    ExpandNode = Table.ExpandRecordColumn(ExpandColumn, "node", 
        {"id", "title", "createdAt", "current_phase", "fields"}, 
        {"Card_ID", "Título", "Data_Criação", "Fase_Atual", "Campos"}
    ),
    ExpandPhase = Table.ExpandRecordColumn(ExpandNode, "Fase_Atual", {"name"}, {"Fase"}),
    
    // Expandir campos customizados
    ExpandFields = Table.ExpandListColumn(ExpandPhase, "Campos"),
    ExpandFieldsRecord = Table.ExpandRecordColumn(ExpandFields, "Campos", 
        {"name", "value"}, 
        {"Campo_Nome", "Campo_Valor"}
    ),
    
    // Criar colunas dinâmicas para cada campo
    PivotFields = Table.Pivot(
        ExpandFieldsRecord, 
        List.Distinct(ExpandFieldsRecord[Campo_Nome]), 
        "Campo_Nome", 
        "Campo_Valor"
    )
in
    PivotFields
"""

print("=" * 70)
print("📚 QUERIES GRAPHQL DISPONÍVEIS PARA O PIPEFY")
print("=" * 70)
print("\n✅ Script atualizado com query corrigida!")
print("\n📝 Execute novamente o script token_bi.py para ver seus pipes.")
print("\n💡 Após identificar o ID do pipe que você deseja conectar, copie a query")
print("   'query_cards_pipe' ou use o código 'powerquery_exemplo' abaixo.")
print("-" * 70)
print("🔽 CÓDIGO POWER QUERY PARA COPIAR E COLAR NO POWER BI 🔽")
print("-" * 70)
print(powerquery_exemplo)
print("-" * 70)