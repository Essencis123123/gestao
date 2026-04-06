"""
Ferramenta de Gestão - Versão Web
Converte a ferramenta de gestão em um site web usando Flask
com o tema visual do anywhere.py
"""

import os
import sys
import json
import pandas as pd
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify
import traceback

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================================
# COLORES E CONFIGURAÇÕES
# ============================================================================
THEME_COLORS = {
    'dark': '#1a3a2a',
    'light': '#f4f3ef',
    'accent': '#2d6348',
    'gold': '#b8903a',
    'text': '#1a1916',
}

# ============================================================================
# CARREGAMENTO DE DADOS
# ============================================================================

def load_data_safe(filename):
    """Carrega dados de um arquivo Excel com tratamento de erro."""
    try:
        path = os.path.join(BASE_DIR, filename)
        if not os.path.exists(path):
            print(f"[!] Arquivo não encontrado: {filename}")
            return []
        
        df = pd.read_excel(path, engine='openpyxl')
        df.columns = df.columns.str.strip()
        print(f"[OK] {filename} carregado! Colunas: {list(df.columns)} | Linhas: {len(df)}")
        return df.fillna('').to_dict('records')
    except Exception as e:
        print(f"[ERRO] Erro ao ler {filename}: {e}")
        traceback.print_exc()
        return []

print("\n" + "="*60)
print("CARREGANDO APLICAÇÃO WEB - FERRAMENTA DE GESTÃO")
print("="*60 + "\n")

# Carregamento de dados principais
try:
    materiais_data = load_data_safe('codigos.xlsx')
    estoque_data = load_data_safe('Solvi - Estoque Periodico.xlsx')
    servicos_direto = load_data_safe('Lista de Códigos - Itens direto.xlsx')
    servicos_indireto = load_data_safe('Lista de Códigos - Itens indireto.xlsx')
    servicos_despesas = load_data_safe('Lista de Códigos - Itens despesas.xlsx')
except Exception as e:
    print(f"[ERRO] Falha ao carregar dados: {e}")
    materiais_data = []
    estoque_data = []
    servicos_direto = []
    servicos_indireto = []
    servicos_despesas = []

total_records = (len(materiais_data) + len(estoque_data) + 
                len(servicos_direto) + len(servicos_indireto) + len(servicos_despesas))

print(f"[OK] TOTAL DE REGISTROS CARREGADOS: {total_records}\n" + "="*60 + "\n")

# ============================================================================
# API JSON ENDPOINTS
# ============================================================================

@app.route('/api/stats', methods=['GET'])
def get_stats():
    """Retorna estatísticas gerais da aplicação."""
    return jsonify({
        'total_materiais': len(materiais_data),
        'total_estoque': len(estoque_data),
        'total_servicos_direto': len(servicos_direto),
        'total_servicos_indireto': len(servicos_indireto),
        'total_servicos_despesas': len(servicos_despesas),
        'total_geral': total_records,
        'ultima_atualizacao': datetime.now().isoformat()
    })

@app.route('/api/materiais', methods=['GET'])
def get_materiais():
    """Retorna lista de materiais com busca opcional."""
    search = request.args.get('search', '').lower()
    limit = int(request.args.get('limit', 50))
    
    results = materiais_data
    if search:
        results = [
            m for m in materiais_data 
            if search in str(m).lower()
        ]
    
    return jsonify(results[:limit])

@app.route('/api/estoque', methods=['GET'])
def get_estoque():
    """Retorna dados de estoque com busca opcional."""
    search = request.args.get('search', '').lower()
    limit = int(request.args.get('limit', 50))
    
    results = estoque_data
    if search:
        results = [
            e for e in estoque_data 
            if search in str(e).lower()
        ]
    
    return jsonify(results[:limit])

@app.route('/api/servicos', methods=['GET'])
def get_servicos():
    """Retorna dados de serviços."""
    tipo = request.args.get('tipo', 'todos')
    limit = int(request.args.get('limit', 50))
    
    if tipo == 'direto':
        results = servicos_direto
    elif tipo == 'indireto':
        results = servicos_indireto
    elif tipo == 'despesas':
        results = servicos_despesas
    else:  # todos
        results = servicos_direto + servicos_indireto + servicos_despesas
    
    return jsonify(results[:limit])

# ============================================================================
# PÁGINAS
# ============================================================================

HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Ferramenta de Gestão</title>
  <style>
    :root {
      --bg:           #f4f3ef;
      --bg2:          #eceae4;
      --surface:      #ffffff;
      --surface2:     #f9f8f5;
      --border:       #dedad2;
      --border2:      #c8c3b8;
      --text:         #1a1916;
      --text2:        #4a4740;
      --muted:        #8c897f;
      --accent:       #1a3a2a;
      --accent-light: #2d6348;
      --accent-pale:  #e8f0eb;
      --gold:         #b8903a;
      --gold-pale:    #f5edd8;
      --red:          #8b2e2e;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
      min-height: 100vh;
    }

    .layout { display: flex; min-height: 100vh; }

    /* Sidebar */
    .sidebar {
      width: 260px; background: var(--accent);
      display: flex; flex-direction: column;
      position: fixed; top: 0; left: 0; bottom: 0; z-index: 200;
      color: white;
    }

    .sidebar-logo {
      padding: 32px 28px 28px;
      border-bottom: 1px solid rgba(255,255,255,0.1);
    }

    .logo-mark {
      font-size: 1.2rem; font-weight: 700;
      margin-bottom: 2px;
    }

    .logo-sub {
      font-size: 0.65rem; font-weight: 500;
      text-transform: uppercase; color: rgba(255,255,255,0.45);
    }

    .sidebar-nav {
      padding: 20px 16px;
      flex: 1;
    }

    .nav-item {
      display: flex; align-items: center; gap: 12px;
      padding: 11px 12px; border-radius: 8px;
      color: rgba(255,255,255,0.65); font-size: 0.875rem;
      cursor: pointer; transition: all 0.18s;
      border: none; background: none; width: 100%; text-align: left;
    }

    .nav-item:hover { background: rgba(255,255,255,0.08); color: #fff; }
    .nav-item.active { background: rgba(255,255,255,0.14); color: #fff; }

    .sidebar-footer {
      padding: 20px 28px;
      border-top: 1px solid rgba(255,255,255,0.08);
      font-size: 0.72rem; color: rgba(255,255,255,0.3);
    }

    /* Main Content */
    .main { margin-left: 260px; flex: 1; display: flex; flex-direction: column; }

    .topbar {
      height: 60px; background: var(--surface); border-bottom: 1px solid var(--border);
      display: flex; align-items: center; justify-content: space-between;
      padding: 0 36px; position: sticky; top: 0; z-index: 100;
    }

    .topbar-title {
      font-size: 1.05rem; font-weight: 600; color: var(--text);
    }

    .topbar-date {
      font-size: 0.75rem; color: var(--muted);
    }

    .content {
      padding: 36px; flex: 1;
    }

    .kpi-row {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
      gap: 16px; margin-bottom: 32px;
    }

    .kpi-card {
      background: var(--surface); border: 1px solid var(--border);
      border-radius: 8px; padding: 22px 24px;
      position: relative; overflow: hidden;
    }

    .kpi-card::before {
      content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
      background: var(--accent-light);
    }

    .kpi-label {
      font-size: 0.68rem; font-weight: 600; letter-spacing: 0.14em;
      text-transform: uppercase; color: var(--muted); margin-bottom: 12px;
    }

    .kpi-value {
      font-size: 2.2rem; font-weight: 700; color: var(--text);
      margin-bottom: 6px;
    }

    .kpi-sub {
      font-size: 0.75rem; color: var(--muted);
    }

    .page { display: none; }
    .page.active { display: block; animation: fadeIn 0.3s ease; }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .page-header { margin-bottom: 28px; }
    .page-header h2 {
      font-size: 1.6rem; font-weight: 700; margin-bottom: 4px;
    }
    .page-header p { font-size: 0.82rem; color: var(--muted); }

    .search-input {
      width: 100%; background: var(--surface);
      border: 1px solid var(--border); border-radius: 4px;
      padding: 11px 14px;
      color: var(--text); font-family: Arial, sans-serif; font-size: 0.875rem;
      margin-bottom: 20px;
    }

    .search-input:focus { outline: none; border-color: var(--accent); }

    .table-card {
      background: var(--surface); border: 1px solid var(--border);
      border-radius: 8px; overflow: hidden;
    }

    table {
      width: 100%; border-collapse: collapse;
    }

    thead {
      background: var(--surface2);
      border-bottom: 1px solid var(--border);
    }

    th {
      padding: 12px 16px; text-align: left;
      font-size: 0.65rem; font-weight: 600;
      text-transform: uppercase; color: var(--muted);
    }

    td {
      padding: 12px 16px; border-bottom: 1px solid rgba(26,25,22,0.04);
      font-size: 0.875rem;
    }

    tbody tr:hover {
      background: var(--surface2);
    }

    .loading { text-align: center; color: var(--muted); padding: 20px; }
    .loading::after {
      content: '...'; animation: dots 1.5s steps(4, end) infinite;
    }

    @keyframes dots {
      0%, 20% { content: '.'; }
      40% { content: '..'; }
      60%, 100% { content: '...'; }
    }

    .stat-count {
      background: var(--accent-pale); color: var(--accent);
      padding: 3px 11px; border-radius: 20px; font-size: 0.7rem; font-weight: 500;
    }
  </style>
</head>
<body>
  <div class="layout">
    <!-- Sidebar -->
    <div class="sidebar">
      <div class="sidebar-logo">
        <div class="logo-mark">📊 Gestão</div>
        <div class="logo-sub">Sistema de Controle</div>
      </div>

      <nav class="sidebar-nav">
        <button class="nav-item active" data-page="home">
          🏠 Dashboard
        </button>
        <button class="nav-item" data-page="materiais">
          📦 Materiais
          <span class="stat-count" id="mat-count">{{ total_materiais }}</span>
        </button>
        <button class="nav-item" data-page="estoque">
          📈 Estoque
          <span class="stat-count" id="est-count">{{ total_estoque }}</span>
        </button>
        <button class="nav-item" data-page="servicos">
          🔧 Serviços
          <span class="stat-count" id="srv-count">{{ total_servicos }}</span>
        </button>
      </nav>

      <div class="sidebar-footer">
        Ferramenta de Gestão v2.0<br>
        <span id="server-time">--:--</span>
      </div>
    </div>

    <!-- Main Content -->
    <div class="main">
      <div class="topbar">
        <div class="topbar-title" id="page-title">Dashboard</div>
        <div class="topbar-date" id="topbar-date">--/--/-- --:--</div>
      </div>

      <div class="content">
        <!-- HOME PAGE -->
        <div class="page active" id="home">
          <div style="margin-bottom: 28px;">
            <h1 style="font-size: 2rem; margin-bottom: 8px;">Dashboard</h1>
            <p style="color: var(--muted); font-size: 0.9rem;">Visão geral do sistema de gestão</p>
          </div>

          <div class="kpi-row">
            <div class="kpi-card">
              <div class="kpi-label">Total de Materiais</div>
              <div class="kpi-value" id="kpi-materiais">{{ total_materiais }}</div>
              <div class="kpi-sub">registros</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-label">Total de Estoque</div>
              <div class="kpi-value" id="kpi-estoque">{{ total_estoque }}</div>
              <div class="kpi-sub">itens</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-label">Serviços Direto</div>
              <div class="kpi-value" id="kpi-srv-direto">{{ total_servicos_direto }}</div>
              <div class="kpi-sub">registros</div>
            </div>
            <div class="kpi-card">
              <div class="kpi-label">Total Geral</div>
              <div class="kpi-value" id="kpi-total">{{ total_geral }}</div>
              <div class="kpi-sub">todos os registros</div>
            </div>
          </div>

          <h3 style="margin: 32px 0 16px; font-size: 1.3rem;">Informações do Sistema</h3>
          <div class="table-card">
            <table>
              <thead>
                <tr>
                  <th>Componente</th>
                  <th>Quantidade</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>Materiais</td>
                  <td><strong id="dash-mat">{{ total_materiais }}</strong></td>
                  <td style="color: var(--accent-light);">✓ Carregado</td>
                </tr>
                <tr>
                  <td>Estoque</td>
                  <td><strong id="dash-est">{{ total_estoque }}</strong></td>
                  <td style="color: var(--accent-light);">✓ Carregado</td>
                </tr>
                <tr>
                  <td>Serviços Direto</td>
                  <td><strong id="dash-sd">{{ total_servicos_direto }}</strong></td>
                  <td style="color: var(--accent-light);">✓ Carregado</td>
                </tr>
                <tr>
                  <td>Serviços Indireto</td>
                  <td><strong id="dash-si">{{ total_servicos_indireto }}</strong></td>
                  <td style="color: var(--accent-light);">✓ Carregado</td>
                </tr>
                <tr>
                  <td>Serviços Despesas</td>
                  <td><strong id="dash-sp">{{ total_servicos_despesas }}</strong></td>
                  <td style="color: var(--accent-light);">✓ Carregado</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <!-- MATERIAIS PAGE -->
        <div class="page" id="materiais">
          <div class="page-header">
            <h2>Materiais</h2>
            <p>Consulte o catálogo de materiais disponíveis</p>
          </div>
          <input type="text" class="search-input" placeholder="Buscar materiais..." id="search-materiais">
          <div class="table-card">
            <div id="materiais-container" class="loading">Carregando dados...</div>
          </div>
        </div>

        <!-- ESTOQUE PAGE -->
        <div class="page" id="estoque">
          <div class="page-header">
            <h2>Estoque</h2>
            <p>Visualize o estado atual do estoque</p>
          </div>
          <input type="text" class="search-input" placeholder="Buscar no estoque..." id="search-estoque">
          <div class="table-card">
            <div id="estoque-container" class="loading">Carregando dados...</div>
          </div>
        </div>

        <!-- SERVIÇOS PAGE -->
        <div class="page" id="servicos">
          <div class="page-header">
            <h2>Serviços</h2>
            <p>Consulte o catálogo de serviços</p>
          </div>

          <div style="margin-bottom: 20px;">
            <select id="servicos-tipo" style="padding: 10px; border: 1px solid var(--border); border-radius: 4px;">
              <option value="todos">Todos os Serviços</option>
              <option value="direto">Serviços Direto</option>
              <option value="indireto">Serviços Indireto</option>
              <option value="despesas">Serviços Despesas</option>
            </select>
          </div>

          <input type="text" class="search-input" placeholder="Buscar serviços..." id="search-servicos">
          <div class="table-card">
            <div id="servicos-container" class="loading">Carregando dados...</div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <script>
    const stats = {{ stats | tojson }};
    
    // Navegação entre páginas
    document.querySelectorAll('.nav-item').forEach(btn => {
      btn.addEventListener('click', () => {
        const page = btn.getAttribute('data-page');
        if (!page) return;
        
        document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        
        btn.classList.add('active');
        document.getElementById(page).classList.add('active');
        
        const titles = {
          'home': 'Dashboard',
          'materiais': 'Materiais',
          'estoque': 'Estoque',
          'servicos': 'Serviços'
        };
        document.getElementById('page-title').innerText = titles[page] || 'Dashboard';
        
        if (page === 'materiais') loadMateriais();
        if (page === 'estoque') loadEstoque();
        if (page === 'servicos') loadServicos();
      });
    });

    // Atualizar hora
    function updateTime() {
      const now = new Date();
      document.getElementById('topbar-date').innerText = 
        now.toLocaleDateString('pt-BR') + ' ' + now.toLocaleTimeString('pt-BR');
      document.getElementById('server-time').innerText = now.toLocaleTimeString('pt-BR');
    }
    updateTime();
    setInterval(updateTime, 1000);

    // Carregar Materiais
    async function loadMateriais() {
      const container = document.getElementById('materiais-container');
      const search = document.getElementById('search-materiais').value;
      container.innerHTML = '<div class="loading">Carregando...</div>';
      
      try {
        const res = await fetch(`/api/materiais?search=${encodeURIComponent(search)}&limit=100`);
        const data = await res.json();
        
        if (data.length === 0) {
          container.innerHTML = '<p style="padding: 20px; color: var(--muted);">Nenhum resultado encontrado.</p>';
          return;
        }

        let html = '<table><thead><tr>';
        if (data.length > 0) {
          Object.keys(data[0]).forEach(key => {
            html += `<th>${key}</th>`;
          });
        }
        html += '</tr></thead><tbody>';
        
        data.forEach(row => {
          html += '<tr>';
          Object.values(row).forEach(val => {
            html += `<td>${String(val).substring(0, 100)}</td>`;
          });
          html += '</tr>';
        });
        
        html += '</tbody></table>';
        container.innerHTML = html;
      } catch (e) {
        container.innerHTML = `<p style="color: var(--red); padding: 20px;">Erro ao carregar: ${e.message}</p>`;
      }
    }

    // Carregar Estoque
    async function loadEstoque() {
      const container = document.getElementById('estoque-container');
      const search = document.getElementById('search-estoque').value;
      container.innerHTML = '<div class="loading">Carregando...</div>';
      
      try {
        const res = await fetch(`/api/estoque?search=${encodeURIComponent(search)}&limit=100`);
        const data = await res.json();
        
        if (data.length === 0) {
          container.innerHTML = '<p style="padding: 20px; color: var(--muted);">Nenhum resultado encontrado.</p>';
          return;
        }

        let html = '<table><thead><tr>';
        if (data.length > 0) {
          Object.keys(data[0]).forEach(key => {
            html += `<th>${key}</th>`;
          });
        }
        html += '</tr></thead><tbody>';
        
        data.forEach(row => {
          html += '<tr>';
          Object.values(row).forEach(val => {
            html += `<td>${String(val).substring(0, 100)}</td>`;
          });
          html += '</tr>';
        });
        
        html += '</tbody></table>';
        container.innerHTML = html;
      } catch (e) {
        container.innerHTML = `<p style="color: var(--red); padding: 20px;">Erro ao carregar: ${e.message}</p>`;
      }
    }

    // Carregar Serviços
    async function loadServicos() {
      const container = document.getElementById('servicos-container');
      const tipo = document.getElementById('servicos-tipo').value;
      const search = document.getElementById('search-servicos').value;
      container.innerHTML = '<div class="loading">Carregando...</div>';
      
      try {
        const res = await fetch(`/api/servicos?tipo=${tipo}&search=${encodeURIComponent(search)}&limit=100`);
        const data = await res.json();
        
        if (data.length === 0) {
          container.innerHTML = '<p style="padding: 20px; color: var(--muted);">Nenhum resultado encontrado.</p>';
          return;
        }

        let html = '<table><thead><tr>';
        if (data.length > 0) {
          Object.keys(data[0]).forEach(key => {
            html += `<th>${key}</th>`;
          });
        }
        html += '</tr></thead><tbody>';
        
        data.forEach(row => {
          html += '<tr>';
          Object.values(row).forEach(val => {
            html += `<td>${String(val).substring(0, 100)}</td>`;
          });
          html += '</tr>';
        });
        
        html += '</tbody></table>';
        container.innerHTML = html;
      } catch (e) {
        container.innerHTML = `<p style="color: var(--red); padding: 20px;">Erro ao carregar: ${e.message}</p>`;
      }
    }

    // Event listeners para busca
    document.getElementById('search-materiais')?.addEventListener('keyup', loadMateriais);
    document.getElementById('search-estoque')?.addEventListener('keyup', loadEstoque);
    document.getElementById('search-servicos')?.addEventListener('keyup', loadServicos);
    document.getElementById('servicos-tipo')?.addEventListener('change', loadServicos);
  </script>
</body>
</html>
'''

@app.route('/')
def index():
    """Página principal."""
    return render_template_string(
        HTML_TEMPLATE,
        total_materiais=len(materiais_data),
        total_estoque=len(estoque_data),
        total_servicos_direto=len(servicos_direto),
        total_servicos_indireto=len(servicos_indireto),
        total_servicos_despesas=len(servicos_despesas),
        total_servicos=len(servicos_direto) + len(servicos_indireto) + len(servicos_despesas),
        total_geral=total_records,
        stats={
            'total_materiais': len(materiais_data),
            'total_estoque': len(estoque_data),
            'total_servicos_direto': len(servicos_direto),
            'total_servicos_indireto': len(servicos_indireto),
            'total_servicos_despesas': len(servicos_despesas),
            'total_geral': total_records,
        }
    )

# ============================================================================
# INICIALIZAÇÃO
# ============================================================================

if __name__ == '__main__':
    print("\n" + "="*60)
    print("INICIANDO SERVIDOR WEB")
    print("="*60)
    print(f"Acesse em: http://localhost:5000")
    print(f"Presione CTRL+C para parar o servidor")
    print("="*60 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)
