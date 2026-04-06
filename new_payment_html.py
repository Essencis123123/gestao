def generate_payment_report_html(payment_data: Dict) -> str:
    """Gera HTML da página de relatório de pagamentos com abas."""
    
    def fmt_currency(val):
        try:
            if val is None or pd.isna(val):
                return "R$ 0,00"
            val_float = float(val)
            return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"
    
    def fmt_date(val):
        try:
            if pd.isna(val):
                return "-"
            return pd.to_datetime(val).strftime('%d/%m/%Y')
        except:
            return "-"
    
    # Dados fornecidos
    top_fornecedores = payment_data.get('top_fornecedores', pd.DataFrame())
    formas_pagamento = payment_data.get('formas_pagamento', {})
    curva_abc = payment_data.get('curva_abc', {})
    impostos = payment_data.get('impostos', {})
    condicao_menor_28 = payment_data.get('condicao_menor_28', pd.DataFrame())
    valor_total_pago = payment_data.get('valor_total_pago', 0)
    qtd_total_nffs = payment_data.get('qtd_total_nffs', 0)
    
    # Resumo formas de pagamento
    resumo_formas = formas_pagamento.get('resumo', pd.DataFrame())
    fornecedores_boleto = formas_pagamento.get('boleto', [])
    fornecedores_deposito = formas_pagamento.get('deposito', [])
    
    # Curva ABC
    tabela_abc = curva_abc.get('tabela', pd.DataFrame())
    classe_a = curva_abc.get('classe_a', [])
    classe_b = curva_abc.get('classe_b', [])
    classe_c = curva_abc.get('classe_c', [])
    
    # Top Fornecedores
    top_fornecedores_html = ""
    if not top_fornecedores.empty:
        for idx, row in top_fornecedores.iterrows():
            top_fornecedores_html += f"""
            <tr>
                <td>{idx + 1}</td>
                <td><strong>{row['Fornecedor']}</strong></td>
                <td style='text-align: right;'>{fmt_currency(row['Valor Total Pago'])}</td>
                <td style='text-align: center;'>{int(row['Quantidade NFFs'])}</td>
            </tr>
            """
    else:
        top_fornecedores_html = "<tr><td colspan='4' style='text-align: center; color: #999;'>Nenhum dado disponível</td></tr>"
    
    # Formas de Pagamento
    formas_pagamento_html = ""
    if not resumo_formas.empty:
        for idx, row in resumo_formas.iterrows():
            formas_pagamento_html += f"""
            <tr>
                <td><strong>{row['Método']}</strong></td>
                <td style='text-align: center;'>{int(row['Quantidade'])}</td>
                <td style='text-align: right;'>{fmt_currency(row['Valor Total'])}</td>
            </tr>
            """
    else:
        formas_pagamento_html = "<tr><td colspan='3' style='text-align: center; color: #999;'>Nenhum dado disponível</td></tr>"
    
    # Curva ABC
    curva_abc_html = ""
    if not tabela_abc.empty:
        for idx, row in tabela_abc.head(20).iterrows():
            classe_badge = f"<span class='badge-{row['Classe'].lower()}'>{row['Classe']}</span>"
            curva_abc_html += f"""
            <tr>
                <td>{classe_badge}</td>
                <td><strong>{row['Fornecedor']}</strong></td>
                <td style='text-align: right;'>{fmt_currency(row['Valor Total'])}</td>
                <td style='text-align: right;'>{row['Percentual']:.2f}%</td>
                <td style='text-align: right;'>{row['Percentual Acumulado']:.2f}%</td>
            </tr>
            """
    else:
        curva_abc_html = "<tr><td colspan='5' style='text-align: center; color: #999;'>Nenhum dado disponível</td></tr>"
    
    # Condição < 28 dias
    condicao_html = ""
    if not condicao_menor_28.empty:
        for idx, row in condicao_menor_28.head(15).iterrows():
            condicao_html += f"""
            <tr>
                <td><strong>{row['Fornecedor']}</strong></td>
                <td style='text-align: center;'><span class='badge-info'>{int(row['Dias'])} dias</span></td>
                <td>{row['Condição Pagamento']}</td>
                <td style='text-align: right;'>{fmt_currency(row['Valor Total'])}</td>
                <td style='text-align: center;'>{int(row['Qtd NFFs'])}</td>
            </tr>
            """
    else:
        condicao_html = "<tr><td colspan='5' style='text-align: center; color: #999;'>Nenhum fornecedor com condição menor que 28 dias</td></tr>"
    
    # HTML da página COM ABAS
    html = get_base_html()
    html += f"""
    <div class="dashboard-container">
        <div class="header">
            <h1><i class="fas fa-money-bill-wave"></i> Relatório de Pagamentos</h1>
            <p class="subtitle">Análise detalhada de pagamentos, fornecedores e impostos</p>
        </div>
        
        <!-- KPIs -->
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                    <i class="fas fa-dollar-sign"></i>
                </div>
                <div class="kpi-label">Valor Total Pago</div>
                <div class="kpi-value">{fmt_currency(valor_total_pago)}</div>
            </div>
            
            <div class="kpi-card success">
                <div class="kpi-icon">
                    <i class="fas fa-file-invoice"></i>
                </div>
                <div class="kpi-label">Total de NFFs</div>
                <div class="kpi-value">{qtd_total_nffs:,}</div>
            </div>
            
            <div class="kpi-card warning">
                <div class="kpi-icon">
                    <i class="fas fa-receipt"></i>
                </div>
                <div class="kpi-label">Total de Impostos</div>
                <div class="kpi-value">{fmt_currency(impostos.get('Total', 0))}</div>
            </div>
            
            <div class="kpi-card danger">
                <div class="kpi-icon">
                    <i class="fas fa-users"></i>
                </div>
                <div class="kpi-label">Fornecedores Únicos</div>
                <div class="kpi-value">{len(top_fornecedores)}</div>
            </div>
        </div>
        
        <!-- MENU DE ABAS -->
        <div class="tabs-container">
            <div class="tabs-menu">
                <button class="tab-button active" onclick="openTab('buscar')">
                    <i class="fas fa-search"></i> Buscar Fornecedor
                </button>
                <button class="tab-button" onclick="openTab('top-fornecedores')">
                    <i class="fas fa-trophy"></i> Top Fornecedores
                </button>
                <button class="tab-button" onclick="openTab('formas-pagamento')">
                    <i class="fas fa-credit-card"></i> Formas de Pagamento
                </button>
                <button class="tab-button" onclick="openTab('curva-abc')">
                    <i class="fas fa-chart-line"></i> Curva ABC
                </button>
                <button class="tab-button" onclick="openTab('impostos')">
                    <i class="fas fa-receipt"></i> Impostos
                </button>
                <button class="tab-button" onclick="openTab('condicao-pagamento')">
                    <i class="fas fa-clock"></i> Condição &lt; 28 dias
                </button>
            </div>
            
            <!-- TAB: BUSCAR FORNECEDOR -->
            <div id="buscar" class="tab-content active">
                <div class="search-container">
                    <h2><i class="fas fa-search"></i> Buscar Fornecedor</h2>
                    <p style="color: #666; margin-bottom: 20px;">Pesquise por nome ou CNPJ do fornecedor</p>
                    
                    <div class="search-box">
                        <input type="text" id="search-input" placeholder="Digite o nome ou CNPJ do fornecedor..." 
                               class="search-input" onkeypress="if(event.key==='Enter') buscarFornecedor()">
                        <button onclick="buscarFornecedor()" class="search-button">
                            <i class="fas fa-search"></i> Buscar
                        </button>
                    </div>
                    
                    <div id="search-results" class="search-results"></div>
                </div>
            </div>
            
            <!-- TAB: TOP FORNECEDORES -->
            <div id="top-fornecedores" class="tab-content">
                <h2><i class="fas fa-trophy"></i> Top 10 Fornecedores por Valor Pago</h2>
                <p style="color: #666; margin-bottom: 20px;">Fornecedores que mais receberam pagamentos (baseado na coluna Valor Pago)</p>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th style="width: 50px;">#</th>
                            <th>Fornecedor</th>
                            <th style="text-align: right;">Valor Total Pago</th>
                            <th style="text-align: center;">Qtd NFFs</th>
                        </tr>
                    </thead>
                    <tbody>
                        {top_fornecedores_html}
                    </tbody>
                </table>
            </div>
            
            <!-- TAB: FORMAS DE PAGAMENTO -->
            <div id="formas-pagamento" class="tab-content">
                <h2><i class="fas fa-credit-card"></i> Análise de Formas de Pagamento</h2>
                <p style="color: #666; margin-bottom: 20px;">Classificação: Boleto ou Depósito (tudo que não é boleto)</p>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Método de Pagamento</th>
                            <th style="text-align: center;">Quantidade de Pagamentos</th>
                            <th style="text-align: right;">Valor Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        {formas_pagamento_html}
                    </tbody>
                </table>
                
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 24px;">
                    <div>
                        <h3 style="margin-bottom: 12px; color: var(--primary);">
                            <i class="fas fa-barcode"></i> Fornecedores com Boleto ({len(fornecedores_boleto)})
                        </h3>
                        <button onclick="toggleList('boleto')" class="btn-primary" style="margin-bottom: 12px;">
                            <i class="fas fa-eye"></i> Mostrar/Ocultar Fornecedores
                        </button>
                        <div id="boleto-list" style="display: none; max-height: 300px; overflow-y: auto; background: #f7fafc; padding: 16px; border-radius: 8px;">
                            {'<br>'.join(['• ' + f for f in fornecedores_boleto]) if fornecedores_boleto else 'Nenhum fornecedor'}
                        </div>
                    </div>
                    
                    <div>
                        <h3 style="margin-bottom: 12px; color: var(--success);">
                            <i class="fas fa-university"></i> Fornecedores com Depósito ({len(fornecedores_deposito)})
                        </h3>
                        <button onclick="toggleList('deposito')" class="btn-success" style="margin-bottom: 12px;">
                            <i class="fas fa-eye"></i> Mostrar/Ocultar Fornecedores
                        </button>
                        <div id="deposito-list" style="display: none; max-height: 300px; overflow-y: auto; background: #f7fafc; padding: 16px; border-radius: 8px;">
                            {'<br>'.join(['• ' + f for f in fornecedores_deposito]) if fornecedores_deposito else 'Nenhum fornecedor'}
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- TAB: CURVA ABC -->
            <div id="curva-abc" class="tab-content">
                <h2><i class="fas fa-chart-line"></i> Curva ABC - Fornecedores por Valor Financeiro</h2>
                <div class="info-box" style="margin-bottom: 20px;">
                    <strong>Legenda:</strong> 
                    <span class="badge-a">A</span> 80% do valor | 
                    <span class="badge-b">B</span> 15% do valor | 
                    <span class="badge-c">C</span> 5% do valor
                </div>
                
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 24px;">
                    <div class="stat-box" style="border-left: 4px solid #48bb78;">
                        <div class="stat-label">Classe A</div>
                        <div class="stat-value">{len(classe_a)} fornecedores</div>
                        <div class="stat-detail">80% do valor total</div>
                    </div>
                    <div class="stat-box" style="border-left: 4px solid #ed8936;">
                        <div class="stat-label">Classe B</div>
                        <div class="stat-value">{len(classe_b)} fornecedores</div>
                        <div class="stat-detail">15% do valor total</div>
                    </div>
                    <div class="stat-box" style="border-left: 4px solid #f56565;">
                        <div class="stat-label">Classe C</div>
                        <div class="stat-value">{len(classe_c)} fornecedores</div>
                        <div class="stat-detail">5% do valor total</div>
                    </div>
                </div>
                
                <table class="data-table">
                    <thead>
                        <tr>
                            <th style="width: 80px;">Classe</th>
                            <th>Fornecedor</th>
                            <th style="text-align: right;">Valor Total Pago</th>
                            <th style="text-align: right;">% Individual</th>
                            <th style="text-align: right;">% Acumulado</th>
                        </tr>
                    </thead>
                    <tbody>
                        {curva_abc_html}
                    </tbody>
                </table>
            </div>
            
            <!-- TAB: IMPOSTOS -->
            <div id="impostos" class="tab-content">
                <h2><i class="fas fa-file-invoice-dollar"></i> Detalhamento de Impostos</h2>
                <p style="color: #666; margin-bottom: 20px;">Soma de todos os impostos retidos nos pagamentos</p>
                
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px;">
                    <div class="stat-box">
                        <div class="stat-label">ISS</div>
                        <div class="stat-value" style="font-size: 1.5rem;">{fmt_currency(impostos.get('ISS', 0))}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">IR</div>
                        <div class="stat-value" style="font-size: 1.5rem;">{fmt_currency(impostos.get('IR', 0))}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">INSS</div>
                        <div class="stat-value" style="font-size: 1.5rem;">{fmt_currency(impostos.get('INSS', 0))}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">PIS</div>
                        <div class="stat-value" style="font-size: 1.5rem;">{fmt_currency(impostos.get('PIS', 0))}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">COFINS</div>
                        <div class="stat-value" style="font-size: 1.5rem;">{fmt_currency(impostos.get('COFINS', 0))}</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-label">CSLL</div>
                        <div class="stat-value" style="font-size: 1.5rem;">{fmt_currency(impostos.get('CSLL', 0))}</div>
                    </div>
                </div>
                
                <div class="stat-box" style="margin-top: 24px; background: linear-gradient(135deg, #f56565 0%, #e53e3e 100%); color: white; border-left: none;">
                    <div class="stat-label" style="color: rgba(255,255,255,0.9);">TOTAL DE IMPOSTOS</div>
                    <div class="stat-value" style="color: white; font-size: 2.5rem;">{fmt_currency(impostos.get('Total', 0))}</div>
                </div>
            </div>
            
            <!-- TAB: CONDIÇÃO DE PAGAMENTO -->
            <div id="condicao-pagamento" class="tab-content">
                <h2><i class="fas fa-clock"></i> Fornecedores com Condição de Pagamento Menor que 28 Dias</h2>
                <p style="color: #666; margin-bottom: 20px;">Fornecedores com prazo de pagamento inferior a 28 dias</p>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Fornecedor</th>
                            <th style="text-align: center;">Dias</th>
                            <th>Condição de Pagamento</th>
                            <th style="text-align: right;">Valor Total</th>
                            <th style="text-align: center;">Qtd NFFs</th>
                        </tr>
                    </thead>
                    <tbody>
                        {condicao_html}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <style>
        /* Estilos das Abas */
        .tabs-container {{
            background: white;
            border-radius: var(--radius);
            overflow: hidden;
            box-shadow: var(--shadow-md);
            margin-top: 24px;
        }}
        
        .tabs-menu {{
            display: flex;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
        }}
        
        .tab-button {{
            flex: 1;
            min-width: 150px;
            padding: 16px 20px;
            border: none;
            background: transparent;
            color: rgba(255, 255, 255, 0.8);
            font-weight: 600;
            font-size: 0.9rem;
            cursor: pointer;
            transition: all 0.3s ease;
            border-bottom: 3px solid transparent;
            white-space: nowrap;
        }}
        
        .tab-button:hover {{
            background: rgba(255, 255, 255, 0.1);
            color: white;
        }}
        
        .tab-button.active {{
            background: rgba(255, 255, 255, 0.2);
            color: white;
            border-bottom-color: #48bb78;
        }}
        
        .tab-content {{
            display: none;
            padding: 32px;
            animation: fadeIn 0.3s ease;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        /* Estilos do Campo de Busca */
        .search-container {{
            max-width: 800px;
            margin: 0 auto;
        }}
        
        .search-container h2 {{
            font-size: 1.8rem;
            color: var(--dark);
            margin-bottom: 8px;
        }}
        
        .search-box {{
            display: flex;
            gap: 12px;
            margin-bottom: 24px;
        }}
        
        .search-input {{
            flex: 1;
            padding: 14px 20px;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            font-size: 1rem;
            transition: all 0.2s ease;
        }}
        
        .search-input:focus {{
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }}
        
        .search-button {{
            padding: 14px 32px;
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            font-size: 1rem;
            cursor: pointer;
            transition: all 0.2s ease;
        }}
        
        .search-button:hover {{
            transform: translateY(-2px);
            box-shadow: var(--shadow-lg);
        }}
        
        .search-results {{
            min-height: 100px;
        }}
        
        .result-card {{
            background: #f7fafc;
            border-left: 4px solid var(--primary);
            padding: 24px;
            border-radius: 8px;
            margin-bottom: 16px;
        }}
        
        .result-header {{
            display: flex;
            justify-content: space-between;
            align-items: start;
            margin-bottom: 16px;
        }}
        
        .result-title {{
            font-size: 1.3rem;
            font-weight: 700;
            color: var(--dark);
        }}
        
        .result-badge {{
            background: var(--success);
            color: white;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }}
        
        .result-info {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 20px;
        }}
        
        .info-item {{
            background: white;
            padding: 12px;
            border-radius: 6px;
        }}
        
        .info-label {{
            font-size: 0.75rem;
            color: #718096;
            text-transform: uppercase;
            font-weight: 600;
            margin-bottom: 4px;
        }}
        
        .info-value {{
            font-size: 1.1rem;
            font-weight: 700;
            color: var(--dark);
        }}
        
        .historico-table {{
            margin-top: 20px;
        }}
        
        .historico-title {{
            font-size: 1rem;
            font-weight: 600;
            color: var(--dark);
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        
        .no-results {{
            text-align: center;
            padding: 40px;
            color: #999;
        }}
        
        .no-results i {{
            font-size: 64px;
            margin-bottom: 16px;
            display: block;
        }}
        
        /* Estilos Gerais */
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 16px;
        }}
        
        .data-table thead {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }}
        
        .data-table th {{
            padding: 14px;
            text-align: left;
            font-weight: 600;
            font-size: 0.9rem;
        }}
        
        .data-table tbody tr {{
            border-bottom: 1px solid #e2e8f0;
            transition: background 0.2s ease;
        }}
        
        .data-table tbody tr:hover {{
            background: #f7fafc;
        }}
        
        .data-table td {{
            padding: 14px;
            font-size: 0.95rem;
        }}
        
        .badge-a {{
            background: #48bb78;
            color: white;
            padding: 4px 12px;
            border-radius: 12px;
            font-weight: 700;
            font-size: 0.85rem;
        }}
        
        .badge-b {{
            background: #ed8936;
            color: white;
            padding: 4px 12px;
            border-radius: 12px;
            font-weight: 700;
            font-size: 0.85rem;
        }}
        
        .badge-c {{
            background: #f56565;
            color: white;
            padding: 4px 12px;
            border-radius: 12px;
            font-weight: 700;
            font-size: 0.85rem;
        }}
        
        .badge-info {{
            background: #4299e1;
            color: white;
            padding: 4px 12px;
            border-radius: 12px;
            font-weight: 600;
            font-size: 0.85rem;
        }}
        
        .stat-box {{
            background: #f7fafc;
            padding: 20px;
            border-radius: 12px;
            border-left: 4px solid var(--primary);
        }}
        
        .stat-label {{
            font-size: 0.85rem;
            color: #718096;
            font-weight: 600;
            text-transform: uppercase;
            margin-bottom: 8px;
        }}
        
        .stat-value {{
            font-size: 2rem;
            font-weight: 800;
            color: var(--dark);
            line-height: 1.2;
        }}
        
        .stat-detail {{
            font-size: 0.85rem;
            color: #999;
            margin-top: 4px;
        }}
        
        .info-box {{
            background: #edf2f7;
            padding: 12px 16px;
            border-radius: 8px;
            font-size: 0.9rem;
        }}
        
        .btn-primary, .btn-success {{
            padding: 10px 20px;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            font-size: 0.9rem;
        }}
        
        .btn-primary {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            color: white;
        }}
        
        .btn-success {{
            background: linear-gradient(135deg, var(--success) 0%, #38a169 100%);
            color: white;
        }}
        
        .btn-primary:hover, .btn-success:hover {{
            transform: translateY(-2px);
            box-shadow: var(--shadow-md);
        }}
    </style>
    
    <script>
        // Dados para busca (passados do Python)
        const paymentDataJSON = {json.dumps({
            'top_fornecedores': top_fornecedores.to_dict('records') if not top_fornecedores.empty else [],
            'valor_total_pago': valor_total_pago,
            'qtd_total_nffs': qtd_total_nffs
        })};
        
        function openTab(tabName) {{
            // Esconder todos os conteúdos
            const tabContents = document.getElementsByClassName('tab-content');
            for (let i = 0; i < tabContents.length; i++) {{
                tabContents[i].classList.remove('active');
            }}
            
            // Remover ativo de todos os botões
            const tabButtons = document.getElementsByClassName('tab-button');
            for (let i = 0; i < tabButtons.length; i++) {{
                tabButtons[i].classList.remove('active');
            }}
            
            // Mostrar conteúdo da aba selecionada
            document.getElementById(tabName).classList.add('active');
            event.target.classList.add('active');
        }}
        
        function buscarFornecedor() {{
            const termo = document.getElementById('search-input').value.trim();
            const resultsDiv = document.getElementById('search-results');
            
            if (!termo) {{
                resultsDiv.innerHTML = '<div class="no-results"><i class="fas fa-info-circle"></i><p>Digite um nome ou CNPJ para buscar</p></div>';
                return;
            }}
            
            resultsDiv.innerHTML = '<div style="text-align: center; padding: 40px;"><i class="fas fa-spinner fa-spin" style="font-size: 48px; color: var(--primary);"></i><p style="margin-top: 16px;">Buscando...</p></div>';
            
            // Simular busca (em produção, fazer chamada para o backend)
            setTimeout(() => {{
                const fornecedores = paymentDataJSON.top_fornecedores;
                const resultados = fornecedores.filter(f => 
                    f.Fornecedor.toUpperCase().includes(termo.toUpperCase())
                );
                
                if (resultados.length === 0) {{
                    resultsDiv.innerHTML = `
                        <div class="no-results">
                            <i class="fas fa-search"></i>
                            <p><strong>Nenhum fornecedor encontrado</strong></p>
                            <p style="font-size: 0.9rem; color: #999;">Não encontramos fornecedores com o termo "${{termo}}"</p>
                        </div>
                    `;
                }} else {{
                    let html = '';
                    resultados.forEach(fornecedor => {{
                        html += `
                            <div class="result-card">
                                <div class="result-header">
                                    <div class="result-title">${{fornecedor.Fornecedor}}</div>
                                    <div class="result-badge"><i class="fas fa-check-circle"></i> Pago</div>
                                </div>
                                
                                <div class="result-info">
                                    <div class="info-item">
                                        <div class="info-label">Valor Total Pago</div>
                                        <div class="info-value" style="color: var(--success);">R$ ${{fornecedor['Valor Total Pago'].toLocaleString('pt-BR', {{minimumFractionDigits: 2}})}}</div>
                                    </div>
                                    <div class="info-item">
                                        <div class="info-label">Quantidade de Faturamentos</div>
                                        <div class="info-value" style="color: var(--primary);">${{fornecedor['Quantidade NFFs']}} NFFs</div>
                                    </div>
                                </div>
                                
                                <p style="margin-top: 12px; color: #666; font-size: 0.9rem;">
                                    <i class="fas fa-info-circle"></i> Este fornecedor recebeu ${{fornecedor['Quantidade NFFs']}} pagamento(s) totalizando R$ ${{fornecedor['Valor Total Pago'].toLocaleString('pt-BR', {{minimumFractionDigits: 2}})}}
                                </p>
                            </div>
                        `;
                    }});
                    resultsDiv.innerHTML = html;
                }}
            }}, 500);
        }}
        
        function toggleList(type) {{
            const list = document.getElementById(type + '-list');
            list.style.display = list.style.display === 'none' ? 'block' : 'none';
        }}
    </script>
    </body>
    </html>
    """
    
    return html
