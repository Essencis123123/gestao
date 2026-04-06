import os
import sys
import datetime
import pandas as pd
import json
import glob
import requests
from functools import wraps
from flask import Flask, render_template_string, session, redirect, url_for, request, jsonify

app = Flask(__name__)
app.secret_key = 'gestao_oracle_secret_2026_xK9mPqR'  # chave segura para sessões

# ─── CONSTANTES DE AUTENTICAÇÃO ───────────────────────────────────────────────
ADMIN_LOGIN = 'admin'
ADMIN_SENHA = 'admin'

# ─── PIPEFY CONFIG ────────────────────────────────────────────────────────────
BASE_GESTAO = r"C:\Users\2700024\Desktop\Ferramenta de gestão"
PIPEFY_TOKEN_URL = "https://app.pipefy.com/oauth/token"
PIPEFY_API_URL   = "https://api.pipefy.com/graphql"

_pipefy_access_token = None
_pipefy_token_expiry = None

def _get_pipefy_config():
    cfg_path = os.path.join(BASE_GESTAO, 'config.json')
    try:
        with open(cfg_path, 'r', encoding='utf-8') as f:
            return json.load(f).get('pipefy', {})
    except Exception:
        return {}

def _get_pipefy_token():
    global _pipefy_access_token, _pipefy_token_expiry
    now = datetime.datetime.now()
    if _pipefy_access_token and _pipefy_token_expiry and now < _pipefy_token_expiry:
        return _pipefy_access_token
    cfg = _get_pipefy_config()
    client_id     = cfg.get('client_id', '')
    client_secret = cfg.get('client_secret', '')
    if not client_id or not client_secret:
        return None
    try:
        resp = requests.post(PIPEFY_TOKEN_URL, data={
            'grant_type': 'client_credentials',
            'client_id': client_id,
            'client_secret': client_secret,
        }, timeout=15)
        data = resp.json()
        _pipefy_access_token = data.get('access_token')
        expires_in = data.get('expires_in', 3600)
        _pipefy_token_expiry = now + datetime.timedelta(seconds=expires_in - 60)
        return _pipefy_access_token
    except Exception as e:
        print(f"[PIPEFY] Falha no token: {e}")
        return None

def _pipefy_query(query, variables=None):
    token = _get_pipefy_token()
    if not token:
        return {'errors': [{'message': 'Token não disponível'}]}
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    payload = {'query': query}
    if variables:
        payload['variables'] = variables
    try:
        resp = requests.post(PIPEFY_API_URL, json=payload, headers=headers, timeout=30)
        return resp.json()
    except Exception as e:
        return {'errors': [{'message': str(e)}]}

def _pipefy_get_all_cards(pipe_id):
    """Busca todos os cards de um pipe com paginação."""
    all_cards = []
    cursor = None
    while True:
        after_clause = f', after: "{cursor}"' if cursor else ''
        q = f"""query {{
          pipe(id: {pipe_id}) {{
            cards(first: 50{after_clause}) {{
              pageInfo {{ hasNextPage endCursor }}
              edges {{
                node {{
                  id title url
                  created_at finished_at due_date
                  current_phase {{ name }}
                  assignees {{ name email }}
                  labels {{ name }}
                  fields {{ name value }}
                  phases_history {{ phase {{ name }} duration lastTimeOut }}
                }}
              }}
            }}
          }}
        }}"""
        result = _pipefy_query(q)
        data = result.get('data', {}).get('pipe', {}).get('cards', {})
        edges = data.get('edges', [])
        for e in edges:
            if e.get('node'):
                all_cards.append(e['node'])
        page_info = data.get('pageInfo', {})
        if not page_info.get('hasNextPage'):
            break
        cursor = page_info.get('endCursor')
    return all_cards

# ─── PIPE IDs mapeados por categoria ──────────────────────────────────────────
PIPEFY_PIPES = {
    'compras_csc':        '302422372',
    'compras_locais':     '302422373',
    'compras_servicos':   '302422374',
    'envio_nfe':          '302422375',
    'reserva_materiais':  '302422376',
}

# ─── HELPER: Proteção de rotas ────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def load_data(filename):
    path = os.path.join(BASE_DIR, filename)
    if not os.path.exists(path):
        print(f"[!] ARQUIVO NAO ENCONTRADO: {filename}")
        return []
    try:
        df = pd.read_excel(path, engine='openpyxl')
        df.columns = df.columns.str.strip()
        print(f"[OK] {filename} carregado! Colunas: {list(df.columns)} | Linhas: {len(df)}")
        return df.fillna('').to_dict('records')
    except Exception as e:
        print(f"[ERRO] Erro ao ler {filename}: {e}")
        return []

CATEGORIA_MAP = {
    '01.01.01': 'COMBUSTÍVEIS',
    '01.01.02': 'LUBRIFICANTES',
    '01.01.03': 'PEÇAS DE FROTA',
    '01.01.04': 'COMPONENTES E IMPLEMENTOS',
    '01.01.05': 'COMPONENTES E IMPLEMENTOS',
    '01.01.06': 'MATERIAIS ADMINISTRATIVOS',
    '01.01.07': 'MATERIAIS ADMINISTRATIVOS',
    '01.01.08': 'EPIS/UNIFORMES',
    '01.01.09': 'INSUMOS',
    '01.01.10': 'COMPONENTES E IMPLEMENTOS',
    '01.01.11': 'MATERIAIS ADMINISTRATIVOS',
    '01.01.12': 'INSUMOS',
    '01.01.13': 'INSUMOS',
    '01.01.14': 'INSUMOS',
    '01.01.15': 'INSUMOS',
    '01.01.16': 'INSUMOS',
    '01.01.17': 'INSUMOS',
    '01.01.18': 'INSUMOS',
    '01.01.19': 'INSUMOS',
    '01.01.20': 'INSUMOS',
    '01.01.21': 'MATERIAIS ADMINISTRATIVOS',
    '01.01.22': 'MATERIAIS ADMINISTRATIVOS',
    '01.01.23': 'MATERIAIS ADMINISTRATIVOS',
}

def classify_item(nome_item):
    s = str(nome_item).strip()
    for prefix in sorted(CATEGORIA_MAP.keys(), key=len, reverse=True):
        if s.startswith(prefix):
            return CATEGORIA_MAP[prefix]
    return 'OUTROS'

def find_file(filename):
    """Find file by name, tolerating extra spaces."""
    path = os.path.join(BASE_DIR, filename)
    if os.path.exists(path):
        return path
    # Try finding by glob pattern (handles extra spaces)
    import glob
    base = os.path.splitext(filename)[0].replace(' ', '*')
    ext = os.path.splitext(filename)[1]
    pattern = os.path.join(BASE_DIR, base + ext)
    matches = glob.glob(pattern)
    if matches:
        print(f"[FOUND] Arquivo encontrado via glob: {os.path.basename(matches[0])}")
        return matches[0]
    # List all xlsx files for debugging
    xlsx_files = [f for f in os.listdir(BASE_DIR) if f.endswith('.xlsx')]
    print(f"[DIR] Arquivos .xlsx na pasta: {xlsx_files}")
    return None

def load_estoque(filename):
    path = find_file(filename)
    if not path:
        print(f"[!] ARQUIVO NAO ENCONTRADO: {filename}")
        return []
    try:
        df = pd.read_excel(path, engine='openpyxl')
        df.columns = df.columns.str.strip()
        print(f"[OK] {filename} carregado! Colunas: {list(df.columns)} | Linhas: {len(df)}")
        # Tenta coluna 'Organizacao' primeiro (codigo), senao 'Nome da Organizacao' (nome)
        org_code_col = next((c for c in df.columns if c.strip() == 'Organização'), None)
        org_name_col = next((c for c in df.columns if c.strip() == 'Nome da Organização'), None)
        sub_col = next((c for c in df.columns if 'Subinvent' in c and 'Nome' in c), None)
        item_col = next((c for c in df.columns if c.strip() == 'Nome do Item'), None)
        if org_code_col:
            df = df[df[org_code_col].astype(str).str.strip() == '00301_EMG_BETIM']
            print(f"  Filtro por '{org_code_col}' = '00301_EMG_BETIM' -> {len(df)} registros")
        elif org_name_col:
            df = df[df[org_name_col].astype(str).str.strip() == 'EMG BETIM']
            print(f"  Filtro por '{org_name_col}' = 'EMG BETIM' -> {len(df)} registros")
        if sub_col:
            df = df[df[sub_col].astype(str).str.strip().isin(['MAT1', 'MAT3','TNQ1'])]
        if item_col:
            df['Categoria'] = df[item_col].apply(classify_item)
        else:
            df['Categoria'] = 'OUTROS'
        df = df.fillna('')
        print(f"[OK] {filename} filtrado: {len(df)} registros (Org=00301_EMG_BETIM, SubInv=MAT1/MAT3/TNQ1)")
        return df.to_dict('records')
    except Exception as e:
        print(f"[ERRO] Erro ao ler {filename}: {e}")
        return []

print("\n" + "="*60)
print("CARREGANDO DADOS DA APLICACAO")
print("="*60)

materiais_data    = load_data('codigos.xlsx')
# Normalizar nome da coluna: 'item' -> 'Item'
for rec in materiais_data:
    if 'item' in rec and 'Item' not in rec:
        rec['Item'] = rec.pop('item')
servicos_direto   = load_data('Lista de Códigos - Itens direto.xlsx')
servicos_indireto = load_data('Lista de Códigos - Itens indireto.xlsx')
servicos_despesas = load_data('Lista de Códigos - Itens despesas.xlsx')
estoque_data      = load_estoque('Solvi - Estoque Periodico.xlsx')

total = len(materiais_data) + len(servicos_direto) + len(servicos_indireto) + len(servicos_despesas) + len(estoque_data)
print(f"\n[OK] TOTAL DE REGISTROS: {total}\n" + "="*60 + "\n")

def get_file_date(filename):
    path = os.path.join(BASE_DIR, filename)
    if os.path.exists(path):
        ts = os.path.getmtime(path)
        return datetime.datetime.fromtimestamp(ts).strftime('%d/%m/%Y %H:%M')
    # Try glob for files with extra spaces in name
    import glob
    pattern = os.path.join(BASE_DIR, filename.replace(' ', '*'))
    matches = glob.glob(pattern)
    if matches:
        ts = os.path.getmtime(matches[0])
        return datetime.datetime.fromtimestamp(ts).strftime('%d/%m/%Y %H:%M')
    return 'N/A'

file_dates = {
    'materiais': get_file_date('codigos.xlsx'),
    'svc_direto': get_file_date('Lista de Códigos - Itens direto.xlsx'),
    'svc_indireto': get_file_date('Lista de Códigos - Itens indireto.xlsx'),
    'svc_despesas': get_file_date('Lista de Códigos - Itens despesas.xlsx'),
    'estoque': get_file_date('Solvi - Estoque Periodico.xlsx'),
}

HTML = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Regional MG/GO · Extensão Oracle Cloud</title>
  <script defer src="https://cdn.jsdelivr.net/npm/fuse.js@7.0.0/dist/fuse.min.js"></script>
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
      --row-hover:    #f0ede6;
      --shadow-sm:    0 1px 3px rgba(26,25,22,0.08);
      --shadow-md:    0 4px 16px rgba(26,25,22,0.10);
      --shadow-lg:    0 12px 40px rgba(26,25,22,0.14);
      --radius:       4px;
      --radius-lg:    8px;
    }

    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    html { scroll-behavior: smooth; }

    body {
      font-family: Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
      min-height: 100vh;
      overflow-x: hidden;
    }

    /* ── PAGE TRANSITIONS ── */
    .page { display: none; min-height: 100vh; flex-direction: column; }
    .page.active {
      display: flex;
      animation: fadeIn 0.3s ease both;
    }
    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(10px); }
      to   { opacity: 1; transform: translateY(0); }
    }

    /* ── INLINE PANEL TRANSITION ── */
    .inline-panel {
      overflow: hidden;
      max-height: 0;
      opacity: 0;
      transition:
        max-height 0.55s cubic-bezier(0.4, 0, 0.2, 1),
        opacity    0.35s ease,
        transform  0.45s cubic-bezier(0.4, 0, 0.2, 1);
      transform: translateY(28px);
      pointer-events: none;
    }
    .inline-panel.open {
      max-height: 2000px;
      opacity: 1;
      transform: translateY(0);
      pointer-events: auto;
    }

    .panel-header-row {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      margin-bottom: 20px;
      gap: 16px;
    }
    .btn-close-panel {
      flex-shrink: 0;
      display: flex; align-items: center; gap: 6px;
      background: var(--surface); border: 1px solid var(--border2);
      border-radius: var(--radius); padding: 8px 14px;
      font-family: Arial, sans-serif;
      font-size: 0.78rem; font-weight: 500; color: var(--muted);
      cursor: pointer; transition: all 0.18s;
    }
    .btn-close-panel:hover {
      border-color: var(--red); color: var(--red); background: #fdf5f5;
    }

    /* ── SIDEBAR / SIDENAV ── */
    .layout { display: flex; min-height: 100vh; }

    .sidebar {
      width: 260px; flex-shrink: 0;
      background: var(--accent);
      display: flex; flex-direction: column;
      position: fixed; top: 0; left: 0; bottom: 0; z-index: 200;
    }

    .sidebar-logo {
      padding: 32px 28px 28px;
      border-bottom: 1px solid rgba(255,255,255,0.1);
    }
    .logo-mark {
      font-family: Arial, sans-serif;
      font-size: 1.5rem; font-weight: 700;
      color: #fff; letter-spacing: 0.02em;
      margin-bottom: 2px;
    }
    .logo-sub {
      font-size: 0.65rem; font-weight: 500; letter-spacing: 0.18em;
      text-transform: uppercase; color: rgba(255,255,255,0.45);
    }

    .sidebar-divider {
      margin: 0 28px;
      border: none; border-top: 1px solid rgba(255,255,255,0.08);
    }

    .sidebar-nav { padding: 20px 16px; flex: 1; }
    .nav-section-label {
      font-size: 0.6rem; font-weight: 600; letter-spacing: 0.2em;
      text-transform: uppercase; color: rgba(255,255,255,0.35);
      padding: 0 12px; margin-bottom: 8px; margin-top: 20px;
    }
    .nav-item {
      display: flex; align-items: center; gap: 12px;
      padding: 11px 12px; border-radius: var(--radius-lg);
      color: rgba(255,255,255,0.65); font-size: 0.875rem; font-weight: 400;
      cursor: pointer; transition: all 0.18s; border: none; background: none;
      width: 100%; text-align: left;
    }
    .nav-item:hover { background: rgba(255,255,255,0.08); color: #fff; }
    .nav-item.active {
      background: rgba(255,255,255,0.14); color: #fff; font-weight: 500;
    }
    .nav-item svg { width: 16px; height: 16px; stroke: currentColor; fill: none; stroke-width: 1.8; flex-shrink: 0; }
    .nav-count {
      margin-left: auto; font-family: Arial, sans-serif;
      font-size: 0.7rem; background: rgba(255,255,255,0.1);
      border-radius: 20px; padding: 2px 8px; color: rgba(255,255,255,0.55);
    }

    .sidebar-footer {
      padding: 20px 28px;
      border-top: 1px solid rgba(255,255,255,0.08);
      font-size: 0.72rem; color: rgba(255,255,255,0.3);
      line-height: 1.6;
    }

    /* ── MAIN CONTENT ── */
    .main { margin-left: 260px; flex: 1; display: flex; flex-direction: column; min-height: 100vh; }

    /* ── TOP BAR ── */
    .topbar {
      height: 60px; background: var(--surface); border-bottom: 1px solid var(--border);
      display: flex; align-items: center; justify-content: space-between;
      padding: 0 36px; position: sticky; top: 0; z-index: 100;
    }
    .topbar-title {
      font-family: Arial, sans-serif;
      font-size: 1.05rem; font-weight: 600; color: var(--text);
      letter-spacing: 0.01em;
    }
    .topbar-meta {
      display: flex; align-items: center; gap: 20px;
    }
    .topbar-date {
      font-size: 0.75rem; color: var(--muted); font-family: Arial, sans-serif;
    }
    .status-dot {
      display: flex; align-items: center; gap: 7px;
      font-size: 0.72rem; color: var(--muted); letter-spacing: 0.04em;
    }
    .dot-live {
      width: 7px; height: 7px; border-radius: 50%;
      background: #2d9c5a; box-shadow: 0 0 0 2px #d1f0dc;
      animation: blink 2.5s ease-in-out infinite;
    }
    @keyframes blink { 0%,100%{ opacity:1; } 50%{ opacity:0.4; } }

    /* ── PAGE CONTENT ── */
    .content { padding: 36px; flex: 1; }

    /* ── HOME PAGE ── */
    .home-header { margin-bottom: 36px; }
    .home-header h1 {
      font-family: Arial, sans-serif;
      font-size: 2rem; font-weight: 700; color: var(--text);
      letter-spacing: -0.01em; margin-bottom: 8px;
    }
    .home-header p { color: var(--muted); font-size: 0.9rem; font-weight: 300; line-height: 1.7; }

    /* KPI Row */
    .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 32px; }
    .kpi-card {
      background: var(--surface); border: 1px solid var(--border);
      border-radius: var(--radius-lg); padding: 22px 24px;
      box-shadow: var(--shadow-sm);
      position: relative; overflow: hidden;
    }
    .kpi-card::before {
      content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
    }
    .kpi-card.k-green::before { background: var(--accent-light); }
    .kpi-card.k-gold::before  { background: var(--gold); }
    .kpi-card.k-gray::before  { background: var(--muted); }
    .kpi-label {
      font-size: 0.68rem; font-weight: 600; letter-spacing: 0.14em;
      text-transform: uppercase; color: var(--muted); margin-bottom: 12px;
    }
    .kpi-value {
      font-family: Arial, sans-serif;
      font-size: 2.2rem; font-weight: 700; color: var(--text); line-height: 1;
      margin-bottom: 6px;
    }
    .kpi-sub { font-size: 0.75rem; color: var(--muted); }

    /* Action cards */
    .action-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 32px; }
    .action-card {
      background: var(--surface); border: 1px solid var(--border);
      border-radius: var(--radius-lg); padding: 28px 32px;
      cursor: pointer; transition: all 0.2s; box-shadow: var(--shadow-sm);
      display: flex; align-items: flex-start; gap: 20px;
    }
    .action-card:hover {
      box-shadow: var(--shadow-md); border-color: var(--border2);
      transform: translateY(-2px);
    }
    .action-card.panel-active {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(26,58,42,0.08);
    }
    .action-icon {
      width: 48px; height: 48px; border-radius: var(--radius-lg);
      display: grid; place-items: center; flex-shrink: 0;
    }
    .action-icon.green { background: var(--accent-pale); }
    .action-icon.gold  { background: var(--gold-pale); }
    .action-icon svg { width: 22px; height: 22px; fill: none; stroke-width: 1.8; }
    .action-icon.green svg { stroke: var(--accent-light); }
    .action-icon.gold  svg { stroke: var(--gold); }
    .action-body { flex: 1; }
    .action-title {
      font-family: Arial, sans-serif;
      font-size: 1.15rem; font-weight: 600; margin-bottom: 6px;
    }
    .action-desc { font-size: 0.82rem; color: var(--muted); line-height: 1.6; margin-bottom: 14px; }
    .action-cta {
      display: inline-flex; align-items: center; gap: 6px;
      font-size: 0.78rem; font-weight: 600; letter-spacing: 0.06em;
      text-transform: uppercase; transition: gap 0.2s;
    }
    .action-cta.green { color: var(--accent-light); }
    .action-cta.gold  { color: var(--gold); }
    .action-cta svg { width: 13px; height: 13px; stroke: currentColor; fill: none; stroke-width: 2.5; }
    .action-card:hover .action-cta { gap: 10px; }

    /* ── SECTION PAGES ── */
    .page-header { margin-bottom: 28px; }
    .page-header h2 {
      font-family: Arial, sans-serif;
      font-size: 1.6rem; font-weight: 700; color: var(--text); margin-bottom: 4px;
    }
    .page-header p { font-size: 0.82rem; color: var(--muted); }

    /* Search */
    .search-row { display: flex; gap: 12px; margin-bottom: 20px; align-items: center; }
    .search-wrap { position: relative; flex: 1; }
    .search-wrap svg.ico {
      position: absolute; left: 14px; top: 50%; transform: translateY(-50%);
      width: 16px; height: 16px; stroke: var(--muted); fill: none; stroke-width: 2; pointer-events: none;
    }
    .search-input {
      width: 100%; background: var(--surface);
      border: 1px solid var(--border); border-radius: var(--radius);
      padding: 11px 14px 11px 42px;
      color: var(--text); font-family: Arial, sans-serif; font-size: 0.875rem;
      transition: all 0.2s; outline: none; box-shadow: var(--shadow-sm);
    }
    .search-input::placeholder { color: var(--muted); }
    .search-input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(26,58,42,0.08); }

    /* Tabs */
    .tab-bar {
      display: flex; gap: 0; margin-bottom: 20px;
      background: var(--surface); border: 1px solid var(--border);
      border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow-sm);
      width: fit-content;
    }
    .tab-btn {
      display: flex; align-items: center; gap: 8px;
      padding: 10px 20px; font-size: 0.8rem; font-weight: 500;
      color: var(--muted); background: none; border: none;
      border-right: 1px solid var(--border); cursor: pointer; transition: all 0.18s;
      white-space: nowrap;
    }
    .tab-btn:last-child { border-right: none; }
    .tab-btn:hover { background: var(--bg2); color: var(--text); }
    .tab-btn.active { background: var(--accent); color: #fff; }
    .tab-num {
      font-family: Arial, sans-serif; font-size: 0.68rem;
      background: rgba(255,255,255,0.18); border-radius: 20px; padding: 1px 7px;
    }
    .tab-btn:not(.active) .tab-num { background: rgba(26,25,22,0.07); }

    /* Table */
    .table-card {
      background: var(--surface); border: 1px solid var(--border);
      border-radius: var(--radius-lg); overflow: hidden; box-shadow: var(--shadow-sm);
    }
    .table-header {
      display: flex; align-items: center; justify-content: space-between;
      padding: 14px 20px; border-bottom: 1px solid var(--border);
      background: var(--surface2);
    }
    .table-header-left { display: flex; align-items: center; gap: 10px; }
    .file-badge {
      font-family: Arial, sans-serif; font-size: 0.7rem;
      background: var(--bg2); border: 1px solid var(--border);
      border-radius: var(--radius); padding: 3px 10px; color: var(--text2);
    }
    .count-pill {
      font-family: Arial, sans-serif; font-size: 0.72rem;
      background: var(--accent-pale); border: 1px solid rgba(26,58,42,0.15);
      color: var(--accent); border-radius: 20px; padding: 3px 11px; font-weight: 500;
    }

    .col-header {
      display: grid; padding: 10px 20px;
      border-bottom: 1px solid var(--border); background: var(--surface2);
    }
    .col-header.g4 { grid-template-columns: 2.5fr 2fr 2fr 1fr; }
    .col-head {
      font-size: 0.65rem; font-weight: 600; letter-spacing: 0.12em;
      text-transform: uppercase; color: var(--muted); cursor: pointer; user-select: none;
    }
    .col-head:hover { color: var(--text); }

    .list-scroll { overflow-y: auto; max-height: 540px; }
    .list-scroll::-webkit-scrollbar { width: 4px; }
    .list-scroll::-webkit-scrollbar-track { background: transparent; }
    .list-scroll::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 10px; }

    .list-row {
      display: grid; padding: 12px 20px; align-items: center;
      border-bottom: 1px solid rgba(26,25,22,0.04);
      transition: background 0.12s; cursor: default;
    }
    .list-row.g1 { grid-template-columns: 1fr; }
    .list-row.g4 { grid-template-columns: 2.5fr 2fr 2fr 1fr; }
    .col-header.g6 { grid-template-columns: 1.8fr 2.5fr 1.4fr 0.7fr 0.9fr 0.9fr; }
    .list-row.g6 { grid-template-columns: 1.8fr 2.5fr 1.4fr 0.7fr 0.9fr 0.9fr; }
    .list-row:last-child { border-bottom: none; }
    .list-row:hover { background: var(--row-hover); }

    .cell-idx {
      font-family: Arial, sans-serif; font-size: 0.68rem;
      color: var(--muted); min-width: 30px; margin-right: 14px;
      padding-right: 14px; border-right: 1px solid var(--border);
      text-align: right; display: inline-block;
    }
    .cell-code {
      display: flex; align-items: center; gap: 0;
      font-family: Arial, sans-serif; font-size: 0.85rem;
      color: var(--accent); font-weight: 500;
    }
    .cell-svc {
      font-family: Arial, sans-serif; font-size: 0.82rem;
      color: var(--accent-light); font-weight: 500;
    }
    .cell-cat { font-size: 0.83rem; color: var(--text2); }
    .cell-uso { font-size: 0.83rem; color: var(--muted); }
    .cell-ind {
      display: inline-flex; padding: 2px 9px; border-radius: var(--radius);
      font-size: 0.68rem; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase;
      background: var(--gold-pale); color: var(--gold);
      border: 1px solid rgba(184,144,58,0.22);
    }

    .cat-badge {
      display: inline-flex; padding: 2px 9px; border-radius: var(--radius);
      font-size: 0.6rem; font-weight: 600; letter-spacing: 0.04em;
      text-transform: uppercase; white-space: nowrap;
    }
    .cat-combustiveis     { background: #fde8e8; color: #8b2e2e; border: 1px solid rgba(139,46,46,0.18); }
    .cat-lubrificantes    { background: #fef3e0; color: #8b6b2e; border: 1px solid rgba(139,107,46,0.18); }
    .cat-pecas            { background: #e3ecf6; color: #2e558b; border: 1px solid rgba(46,85,139,0.18); }
    .cat-componentes      { background: #ece3f6; color: #5b2e8b; border: 1px solid rgba(91,46,139,0.18); }
    .cat-administrativos  { background: #eceae4; color: #4a4740; border: 1px solid rgba(74,71,64,0.18); }
    .cat-epis             { background: var(--accent-pale); color: var(--accent-light); border: 1px solid rgba(26,58,42,0.15); }
    .cat-insumos          { background: #e0f0ef; color: #1a6b63; border: 1px solid rgba(26,107,99,0.18); }
    .cat-outros           { background: #f0f0f0; color: #666; border: 1px solid rgba(0,0,0,0.08); }

    .cell-qty, .cell-cost {
      font-family: Arial, sans-serif; font-size: 0.82rem; color: var(--text2); text-align: right;
    }

    .filter-select {
      background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius);
      padding: 10px 14px; font-family: Arial, sans-serif; font-size: 0.84rem;
      color: var(--text); outline: none; cursor: pointer; box-shadow: var(--shadow-sm);
    }
    .filter-select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(26,58,42,0.08); }

    .btn-copy {
      opacity: 0; margin-left: 10px;
      background: var(--bg2); border: 1px solid var(--border);
      border-radius: var(--radius); padding: 2px 8px;
      font-size: 0.65rem; color: var(--muted); cursor: pointer;
      transition: all 0.15s; font-family: Arial, sans-serif;
    }
    .btn-copy:hover { border-color: var(--accent); color: var(--accent); }
    .list-row:hover .btn-copy { opacity: 1; }

    mark {
      background: rgba(184,144,58,0.2); color: var(--gold);
      border-radius: 2px; padding: 0 2px;
    }

    /* Table footer */
    .table-footer {
      display: flex; align-items: center; justify-content: space-between;
      padding: 12px 20px; border-top: 1px solid var(--border);
      background: var(--surface2); flex-wrap: wrap; gap: 8px;
    }
    .footer-info { font-family: Arial, sans-serif; font-size: 0.72rem; color: var(--muted); }

    .pagination { display: flex; align-items: center; gap: 3px; }
    .pg-btn {
      min-width: 30px; height: 30px; border-radius: var(--radius);
      background: var(--surface); border: 1px solid var(--border);
      color: var(--muted); font-size: 0.75rem; cursor: pointer;
      transition: all 0.15s; font-family: Arial, sans-serif;
      display: grid; place-items: center;
    }
    .pg-btn:hover:not(:disabled) { border-color: var(--accent); color: var(--accent); }
    .pg-btn.active { background: var(--accent); border-color: var(--accent); color: #fff; font-weight: 600; }
    .pg-btn:disabled { opacity: 0.3; cursor: not-allowed; }
    .pg-dots { color: var(--muted); font-size: 0.75rem; padding: 0 3px; }

    /* Empty state */
    .empty-state {
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      padding: 60px 20px; gap: 12px; color: var(--muted); text-align: center;
    }
    .empty-state svg { width: 40px; height: 40px; stroke: var(--border2); fill: none; stroke-width: 1.5; }
    .empty-state strong { font-size: 0.9rem; color: var(--text2); }
    .empty-state span { font-size: 0.8rem; }

    /* Horizontal rule accent */
    .hr-accent {
      border: none; border-top: 1px solid var(--border); margin: 28px 0;
    }

    @media (max-width: 900px) {
      .sidebar { display: none; }
      .main { margin-left: 0; }
      .kpi-row { grid-template-columns: repeat(2, 1fr); }
      .action-grid { grid-template-columns: 1fr; }
      .col-header.g4, .col-header.g6 { display: none; }
      .list-row.g4, .list-row.g6 { grid-template-columns: 1fr; gap: 5px; }
      .content { padding: 20px; }
      .panel-header-row { flex-direction: column; }
    }
  </style>
</head>
<body>

<!-- ══════════ HOME ══════════ -->
<div id="home-page" class="page active">
  <div class="layout">

    <!-- Sidebar -->
    <aside class="sidebar">
      <div class="sidebar-logo">
        <div class="logo-mark">EXTENSÃO ORACLE ERP</div>
        <div class="logo-sub">Sistema para consultas</div>
      </div>
      <nav class="sidebar-nav">
        <div class="nav-section-label">Módulos</div>
        <button class="nav-item active" onclick="goToHome()">
          <svg viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
          Visão Geral
        </button>
        <button class="nav-item" onclick="openPanelMateriais()">
          <svg viewBox="0 0 24 24"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/></svg>
          Materiais
          <span class="nav-count">{{ materiais_count }}</span>
        </button>
        <button class="nav-item" onclick="openPanelServicos()">
          <svg viewBox="0 0 24 24"><path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="2"/></svg>
          Serviços
          <span class="nav-count">{{ servicos_count }}</span>
        </button>
        <a class="nav-item" href="/estoque" style="text-decoration:none;">
          <svg viewBox="0 0 24 24"><path d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4"/></svg>
          Estoque
          <span class="nav-count">{{ estoque_count }}</span>
        </a>
        <div style="font-size:0.6rem;font-weight:600;letter-spacing:0.2em;text-transform:uppercase;color:rgba(255,255,255,0.35);padding:0 12px;margin:16px 0 8px;">Ferramenta de Gestão</div>
        <a class="nav-item" href="/login" style="text-decoration:none;">
          <svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
          Gestão (Login)
        </a>
      </nav>
      <div class="sidebar-footer">
        Base atualizada<br>
        v2.0 · Essencis Sistemas
      </div>
    </aside>

    <!-- Main -->
    <main class="main">
      <div class="topbar">
        <span class="topbar-title">Painel de Consulta</span>
        <div class="topbar-meta">
          <div class="status-dot">
            <div class="dot-live"></div>
            Sistema ativo
          </div>
          <div class="topbar-date" id="current-date"></div>
        </div>
      </div>

      <div class="content">
        <div class="home-header">
          <h1>Consulta de Cadastros</h1>
          <p>Acesse e pesquise materiais e serviços cadastrados no sistema Essencis.</p>
        </div>

        <div class="kpi-row">
          <div class="kpi-card k-green">
            <div class="kpi-label">Materiais</div>
            <div class="kpi-value">{{ materiais_count }}</div>
            <div class="kpi-sub">códigos cadastrados</div>
          </div>
          <div class="kpi-card k-gold">
            <div class="kpi-label">Serviços Diretos</div>
            <div class="kpi-value">{{ direto_count }}</div>
            <div class="kpi-sub">itens diretos</div>
          </div>
          <div class="kpi-card k-gray">
            <div class="kpi-label">Serviços Indiretos</div>
            <div class="kpi-value">{{ indireto_count }}</div>
            <div class="kpi-sub">itens indiretos</div>
          </div>
          <div class="kpi-card k-gray">
            <div class="kpi-label">Despesas</div>
            <div class="kpi-value">{{ despesas_count }}</div>
            <div class="kpi-sub">itens de despesa</div>
          </div>

        </div>

        <div class="action-grid">
          <div class="action-card" id="card-materiais" onclick="openPanelMateriais()">
            <div class="action-icon green">
              <svg viewBox="0 0 24 24"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/></svg>
            </div>
            <div class="action-body">
              <div class="action-title">Códigos de Materiais</div>
              <div class="action-desc">Consulte e pesquise todos os {{ materiais_count }} materiais cadastrados. Busca com correspondência aproximada.</div>
              <div class="action-cta green" id="cta-materiais">
                Acessar base
                <svg viewBox="0 0 24 24"><path d="M5 12h14M12 5l7 7-7 7"/></svg>
              </div>
            </div>
          </div>

          <div class="action-card" id="card-servicos" onclick="openPanelServicos()">
            <div class="action-icon gold">
              <svg viewBox="0 0 24 24"><path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="2"/></svg>
            </div>
            <div class="action-body">
              <div class="action-title">Códigos de Serviços</div>
              <div class="action-desc">Consulte serviços diretos, indiretos e despesas. Total de {{ servicos_count }} serviços disponíveis.</div>
              <div class="action-cta gold" id="cta-servicos">
                Consultar serviços
                <svg viewBox="0 0 24 24"><path d="M5 12h14M12 5l7 7-7 7"/></svg>
              </div>
            </div>
          </div>


        </div>

        <!-- ══ PAINEL INLINE: MATERIAIS ══ -->
        <div class="inline-panel" id="panel-materiais">
          <div class="panel-header-row">
            <div class="page-header" style="margin-bottom:0;">
              <h2>Códigos de Materiais</h2>
              <p>Pesquise por código ou descrição. A busca utiliza correspondência aproximada (fuzzy search).</p>
            </div>
            <button class="btn-close-panel" onclick="closePanel('materiais')">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              Fechar
            </button>
          </div>

          <div class="search-row">
            <div class="search-wrap">
              <svg class="ico" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              <input class="search-input" id="search-mat-inline" type="text"
                     placeholder="Buscar código ou descrição de material…"
                     oninput="onSearchMatInline(this.value)">
            </div>
          </div>

          <div class="table-card">
            <div class="table-header">
              <div class="table-header-left">
                <!-- file-badge hidden -->
              </div>
              <span class="count-pill" id="mat-inline-badge">0 itens</span>
            </div>
            <div class="col-header" style="grid-template-columns:1fr;">
              <span class="col-head" onclick="toggleSortMatInline()" style="cursor:pointer;">
                Item / Código ↕
              </span>
            </div>
            <div class="list-scroll" id="mat-inline-list"></div>
            <div class="table-footer">
              <span class="footer-info" id="mat-inline-footer"></span>
              <div class="pagination" id="mat-inline-pagination"></div>
            </div>
            <div style="text-align:right;padding:6px 16px;font-size:0.75rem;color:var(--muted);border-top:1px solid var(--border);">
              Última atualização da planilha: {{ file_dates['materiais'] }}
            </div>
          </div>
        </div>

        <!-- ══ PAINEL INLINE: SERVIÇOS ══ -->
        <div class="inline-panel" id="panel-servicos">
          <div class="panel-header-row">
            <div class="page-header" style="margin-bottom:0;">
              <h2>Códigos de Serviços</h2>
              <p>Consulte serviços diretos, indiretos e despesas cadastrados no sistema.</p>
            </div>
            <button class="btn-close-panel" onclick="closePanel('servicos')">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              Fechar
            </button>
          </div>

          <div class="tab-bar">
            <button class="tab-btn active" id="tab-direto" onclick="loadServicoInline('direto')">
              Itens Diretos
              <span class="tab-num" id="num-direto">0</span>
            </button>
            <button class="tab-btn" id="tab-indireto" onclick="loadServicoInline('indireto')">
              Itens Indiretos
              <span class="tab-num" id="num-indireto">0</span>
            </button>
            <button class="tab-btn" id="tab-despesa" onclick="loadServicoInline('despesa')">
              Despesas
              <span class="tab-num" id="num-despesa">0</span>
            </button>
          </div>

          <div class="search-row">
            <div class="search-wrap">
              <svg class="ico" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              <input class="search-input" id="search-svc-inline" type="text"
                     placeholder="Buscar por serviço, categoria ou uso pretendido…"
                     oninput="onSearchSvcInline(this.value)">
            </div>
          </div>

          <div class="table-card">
            <div class="table-header">
              <div class="table-header-left">
                <!-- file-badge hidden -->
              </div>
              <span class="count-pill" id="svc-inline-badge">0 itens</span>
            </div>
            <div class="col-header g4">
              <span class="col-head">Item de Serviço</span>
              <span class="col-head">Categoria</span>
              <span class="col-head">Uso Pretendido</span>
              <span class="col-head">Indireto</span>
            </div>
            <div class="list-scroll" id="svc-inline-list"></div>
            <div class="table-footer">
              <span class="footer-info" id="svc-inline-footer"></span>
              <div class="pagination" id="svc-inline-pagination"></div>
            </div>
            <div style="text-align:right;padding:6px 16px;font-size:0.75rem;color:var(--muted);border-top:1px solid var(--border);" id="svc-date-footer">
              Última atualização da planilha: <span id="svc-date-value">{{ file_dates['svc_direto'] }}</span>
            </div>
          </div>
        </div>

      </div><!-- /content -->
    </main>

  </div>
</div>

<script>
// ── GLOBAL ERROR HANDLER ──────────────────────────────────────────────────────
window.onerror = function(msg, url, line, col, err) {
  console.error('JS ERROR:', msg, 'at line', line, 'col', col, err);
  return false;
};

// ── DATA ──────────────────────────────────────────────────────────────────────
const MAT_DATA = {{ materiais_data | tojson | safe }};
const SVC_DATA = {
  direto:   {{ servicos_direto   | tojson | safe }},
  indireto: {{ servicos_indireto | tojson | safe }},
  despesa:  {{ servicos_despesas | tojson | safe }}
};
console.log('Dados carregados:', {
  materiais: MAT_DATA.length,
  direto:    SVC_DATA.direto.length,
  indireto:  SVC_DATA.indireto.length,
  despesa:   SVC_DATA.despesa.length
});

// ── STATE ─────────────────────────────────────────────────────────────────────
const PER_PAGE = 20;

let matInlineState = { filtered: [], page: 1, sort: 1, query: '', fuse: null };
let svcInlineState = { type: 'direto', list: [], filtered: [], page: 1, query: '', fuse: null };


const TAB_FILES = {
  direto:   'Lista de Códigos - Itens direto.xlsx',
  indireto: 'Lista de Códigos - Itens indireto.xlsx',
  despesa:  'Lista de Códigos - Itens despesas.xlsx'
};

// ── HELPERS ───────────────────────────────────────────────────────────────────
function esc(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/'/g,'&#39;').replace(/"/g,'&quot;');
}

function jsesc(str) {
  return String(str).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'\\"').replace(/</g,'\\x3C').replace(/>/g,'\\x3E');
}

function highlight(text, query) {
  const s = esc(text);
  if (!query || query.length < 2) return s;
  const re = new RegExp('('+query.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')', 'gi');
  return s.replace(re, '<mark>$1</mark>');
}

function emptyState(title, sub) {
  return '<div class="empty-state">' +
    '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>' +
    '<strong>' + esc(title) + '</strong>' +
    (sub ? '<span>' + sub + '</span>' : '') +
  '</div>';
}

function copyText(text, btn) {
  navigator.clipboard.writeText(text).then(function() {
    var orig = btn.textContent;
    btn.textContent = '✓ copiado';
    btn.style.color = 'var(--accent)';
    btn.style.borderColor = 'var(--accent)';
    setTimeout(function() {
      btn.textContent = orig;
      btn.style.color = '';
      btn.style.borderColor = '';
    }, 1500);
  }).catch(function() {});
}

function initFuse(data, keys) {
  if (typeof Fuse === 'undefined') return null;
  try {
    return new Fuse(data, { keys: keys, threshold: 0.4, ignoreLocation: true, includeScore: true });
  } catch(e) {
    console.warn('Fuse init error:', e);
    return null;
  }
}

// ── INIT ──────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', function() {
  var d = new Date();
  var fmt = d.toLocaleDateString('pt-BR', { weekday:'short', day:'2-digit', month:'short', year:'numeric' });
  document.querySelectorAll('[id$="-date"], #current-date').forEach(function(el) { if(el) el.textContent = fmt; });

  // Pre-init Fuse for materiais
  if (MAT_DATA.length > 0) {
    matInlineState.fuse = initFuse(MAT_DATA, ['Item']);
    matInlineState.filtered = MAT_DATA.slice();
  }

  // Tab counts
  document.getElementById('num-direto').textContent   = SVC_DATA.direto.length;
  document.getElementById('num-indireto').textContent = SVC_DATA.indireto.length;
  document.getElementById('num-despesa').textContent  = SVC_DATA.despesa.length;

  // Init svc inline state
  svcInlineState.list     = SVC_DATA.direto.slice();
  svcInlineState.filtered = SVC_DATA.direto.slice();
  svcInlineState.fuse     = initFuse(svcInlineState.list, ['Item de Serviço','Nome da Categoria','USO PRETENDIDO','Indireto']);

});

// ── PANEL OPEN / CLOSE ────────────────────────────────────────────────────────
function closeOtherPanels(skip) {
  ['materiais','servicos'].forEach(function(p) {
    if (p === skip) return;
    var el = document.getElementById('panel-' + p);
    if (el && el.classList.contains('open')) {
      el.classList.remove('open');
      var card = document.getElementById('card-' + p);
      if (card) card.classList.remove('panel-active');
      updateCTA(p, false);
    }
  });
}

function openPanelMateriais() {
  console.log('openPanelMateriais called, MAT_DATA.length:', MAT_DATA.length);
  try {
  closeOtherPanels('materiais');

  var pm = document.getElementById('panel-materiais');
  var already = pm.classList.contains('open');

  if (already) {
    pm.classList.remove('open');
    document.getElementById('card-materiais').classList.remove('panel-active');
    updateCTA('materiais', false);
    return;
  }

  // Reset state
  matInlineState.query   = '';
  matInlineState.page    = 1;
  matInlineState.filtered = MAT_DATA.slice();
  var inp = document.getElementById('search-mat-inline');
  if (inp) inp.value = '';

  pm.classList.add('open');
  document.getElementById('card-materiais').classList.add('panel-active');
  updateCTA('materiais', true);
  renderMatInline();

  // Smooth scroll to panel
  setTimeout(function() { pm.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 80);
  } catch(e) { console.error('ERROR in openPanelMateriais:', e); alert('Erro: ' + e.message); }
}

function openPanelServicos() {
  console.log('openPanelServicos called');
  try {
  closeOtherPanels('servicos');

  var ps = document.getElementById('panel-servicos');
  var already = ps.classList.contains('open');

  if (already) {
    ps.classList.remove('open');
    document.getElementById('card-servicos').classList.remove('panel-active');
    updateCTA('servicos', false);
    return;
  }

  ps.classList.add('open');
  document.getElementById('card-servicos').classList.add('panel-active');
  updateCTA('servicos', true);
  loadServicoInline('direto');

  setTimeout(function() { ps.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 80);
  } catch(e) { console.error('ERROR in openPanelServicos:', e); alert('Erro: ' + e.message); }
}

function closePanel(which) {
  var el = document.getElementById('panel-' + which);
  el.classList.remove('open');
  document.getElementById('card-' + which).classList.remove('panel-active');
  updateCTA(which, false);
}

function updateCTA(which, isOpen) {
  var cta = document.getElementById('cta-' + which);
  if (!cta) return;
  var closeIcon = '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
  var openIcon  = '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M5 12h14M12 5l7 7-7 7"/></svg>';
  var labels = {
    materiais: ['Acessar base', 'Fechar base'],
    servicos:  ['Consultar serviços', 'Fechar serviços']
  };
  var l = labels[which] || ['Abrir', 'Fechar'];
  cta.innerHTML = (isOpen ? l[1] + ' ' + closeIcon : l[0] + ' ' + openIcon);
}

// ── NAVIGATION ────────────────────────────────────────────────────────────────
function showPage(id) {
  document.querySelectorAll('.page').forEach(function(p) { p.classList.remove('active'); });
  document.getElementById(id).classList.add('active');
  window.scrollTo(0, 0);
}

function goToHome() { showPage('home-page'); }

// ── MATERIAIS INLINE ──────────────────────────────────────────────────────────
function renderMatInline() {
  var filtered = matInlineState.filtered;
  var page     = matInlineState.page;
  var query    = matInlineState.query;
  var total    = filtered.length;
  var start    = (page-1)*PER_PAGE;
  var end      = Math.min(start+PER_PAGE, total);
  var slice    = filtered.slice(start, end);
  var el       = document.getElementById('mat-inline-list');

  document.getElementById('mat-inline-badge').textContent = total + ' itens';

  if (!slice.length) {
    el.innerHTML = emptyState('Nenhum material encontrado', query ? 'para "' + esc(query) + '"' : '');
    document.getElementById('mat-inline-footer').textContent = '0 resultados';
    document.getElementById('mat-inline-pagination').innerHTML = '';
    return;
  }

  var html = '';
  for (var i = 0; i < slice.length; i++) {
    var item = slice[i];
    var val = String(item.Item || '').trim() || '—';
    var idx = start + i + 1;
    html += '<div class="list-row g1">' +
      '<div class="cell-code">' +
        '<span class="cell-idx">' + idx + '</span>' +
        '<span>' + highlight(val, query) + '</span>' +
        '<button class="btn-copy" onclick="copyText(\'' + jsesc(val) + '\',this)">copiar</button>' +
      '</div>' +
    '</div>';
  }
  el.innerHTML = html;

  document.getElementById('mat-inline-footer').textContent =
    'Exibindo ' + (start+1) + '–' + end + ' de ' + total + ' registros';
  renderPagination('mat-inline', total, page);
}

function onSearchMatInline(q) {
  matInlineState.query = q.trim();
  matInlineState.page  = 1;
  if (!q.trim()) {
    matInlineState.filtered = MAT_DATA.slice();
  } else if (matInlineState.fuse) {
    matInlineState.filtered = matInlineState.fuse.search(q.trim()).map(function(r) { return r.item; });
  } else {
    var lower = q.trim().toLowerCase();
    matInlineState.filtered = MAT_DATA.filter(function(r) { return String(r.Item||'').toLowerCase().indexOf(lower) !== -1; });
  }
  renderMatInline();
}

function toggleSortMatInline() {
  matInlineState.sort *= -1;
  var s = matInlineState.sort;
  matInlineState.filtered.sort(function(a, b) {
    return String(a.Item||'').localeCompare(String(b.Item||'')) * s;
  });
  matInlineState.page = 1;
  renderMatInline();
}

// ── SERVIÇOS INLINE ───────────────────────────────────────────────────────────
function loadServicoInline(tipo) {
  svcInlineState.type     = tipo;
  svcInlineState.list     = (SVC_DATA[tipo] || []).slice();
  svcInlineState.filtered = svcInlineState.list.slice();
  svcInlineState.page     = 1;
  svcInlineState.query    = '';
  var inp = document.getElementById('search-svc-inline');
  if (inp) inp.value = '';

  ['direto','indireto','despesa'].forEach(function(t) {
    var btn = document.getElementById('tab-'+t);
    if (btn) btn.className = 'tab-btn' + (t===tipo ? ' active' : '');
  });

  var SVC_DATES = {
    direto:   '{{ file_dates["svc_direto"] }}',
    indireto: '{{ file_dates["svc_indireto"] }}',
    despesa:  '{{ file_dates["svc_despesas"] }}'
  };
  var dv = document.getElementById('svc-date-value');
  if (dv) dv.textContent = SVC_DATES[tipo] || '';

  svcInlineState.fuse = initFuse(svcInlineState.list, ['Item de Serviço','Nome da Categoria','USO PRETENDIDO','Indireto']);

  renderSvcInline();
}

function renderSvcInline() {
  var filtered = svcInlineState.filtered;
  var page     = svcInlineState.page;
  var query    = svcInlineState.query;
  var total    = filtered.length;
  var start    = (page-1)*PER_PAGE;
  var end      = Math.min(start+PER_PAGE, total);
  var slice    = filtered.slice(start, end);
  var el       = document.getElementById('svc-inline-list');

  document.getElementById('svc-inline-badge').textContent = total + ' itens';

  if (!slice.length) {
    el.innerHTML = emptyState('Nenhum serviço encontrado', query ? 'para "' + esc(query) + '"' : '');
    document.getElementById('svc-inline-footer').textContent = '0 resultados';
    document.getElementById('svc-inline-pagination').innerHTML = '';
    return;
  }

  var html = '';
  for (var i = 0; i < slice.length; i++) {
    var item = slice[i];
    var svc = String(item['Item de Serviço']    || '').trim() || '—';
    var cat = String(item['Nome da Categoria']  || '').trim() || '—';
    var uso = String(item['USO PRETENDIDO']     || '').trim() || '—';
    var ind = String(item['Indireto']           || '').trim() || 'N/A';
    html += '<div class="list-row g4">' +
      '<div class="cell-svc">' +
        highlight(svc, query) +
        '<button class="btn-copy" onclick="copyText(\'' + jsesc(svc) + '\',this)">copiar</button>' +
      '</div>' +
      '<div class="cell-cat">' + highlight(cat, query) + '</div>' +
      '<div class="cell-uso">' + highlight(uso, query) + '</div>' +
      '<div><span class="cell-ind">' + esc(ind) + '</span></div>' +
    '</div>';
  }
  el.innerHTML = html;

  document.getElementById('svc-inline-footer').textContent =
    'Exibindo ' + (start+1) + '–' + end + ' de ' + total + ' registros';
  renderPagination('svc-inline', total, page);
}

function onSearchSvcInline(q) {
  svcInlineState.query = q.trim();
  svcInlineState.page  = 1;
  if (!q.trim()) {
    svcInlineState.filtered = svcInlineState.list.slice();
  } else if (svcInlineState.fuse) {
    svcInlineState.filtered = svcInlineState.fuse.search(q.trim()).map(function(r) { return r.item; });
  } else {
    var lower = q.trim().toLowerCase();
    svcInlineState.filtered = svcInlineState.list.filter(function(r) {
      return String(r['Item de Serviço']||'').toLowerCase().indexOf(lower) !== -1 ||
             String(r['Nome da Categoria']||'').toLowerCase().indexOf(lower) !== -1 ||
             String(r['USO PRETENDIDO']||'').toLowerCase().indexOf(lower) !== -1;
    });
  }
  renderSvcInline();
}

// ── PAGINATION ────────────────────────────────────────────────────────────────
function goMatPage(p) { matInlineState.page = p; renderMatInline(); }
function goSvcPage(p) { svcInlineState.page = p; renderSvcInline(); }

function renderPagination(ns, total, current) {
  var pages = Math.max(1, Math.ceil(total/PER_PAGE));
  var el = document.getElementById(ns+'-pagination');
  if (!el || pages <= 1) { if(el) el.innerHTML=''; return; }

  var goFn = (ns === 'mat-inline') ? 'goMatPage' : 'goSvcPage';

  var html = '<button class="pg-btn" onclick="' + goFn + '(' + (current-1) + ')" ' + (current===1?'disabled':'') + '>‹</button>';
  var DELTA = 2;
  var lo = Math.max(1, current-DELTA), hi = Math.min(pages, current+DELTA);
  if (lo > 1) { html += pgBtn(goFn, 1, current); if(lo>2) html += '<span class="pg-dots">…</span>'; }
  for (var i=lo; i<=hi; i++) html += pgBtn(goFn, i, current);
  if (hi < pages) { if(hi<pages-1) html += '<span class="pg-dots">…</span>'; html += pgBtn(goFn, pages, current); }
  html += '<button class="pg-btn" onclick="' + goFn + '(' + (current+1) + ')" ' + (current===pages?'disabled':'') + '>›</button>';
  el.innerHTML = html;
}

function pgBtn(goFn, p, current) {
  return '<button class="pg-btn' + (p===current?' active':'') + '" onclick="' + goFn + '(' + p + ')">' + p + '</button>';
}
</script>
</body>
</html>'''

@app.route('/')
def cadastros():
    stats = {
        'materiais_count': len(materiais_data),
        'direto_count':    len(servicos_direto),
        'indireto_count':  len(servicos_indireto),
        'despesas_count':  len(servicos_despesas),
        'servicos_count':  len(servicos_direto) + len(servicos_indireto) + len(servicos_despesas),
        'estoque_count':   len(estoque_data),
    }
    return render_template_string(HTML,
        materiais_data    = materiais_data,
        servicos_direto   = servicos_direto,
        servicos_indireto = servicos_indireto,
        servicos_despesas = servicos_despesas,
        file_dates        = file_dates,
        **stats)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGINA ESTOQUE - /estoque
# ═══════════════════════════════════════════════════════════════════════════════
HTML_ESTOQUE = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Estoque Almoxarifado · Extensão Oracle Cloud</title>
  <style>
    :root {
      --bg:#f4f3ef;--bg2:#eceae4;--surface:#ffffff;--surface2:#f9f8f5;
      --border:#dedad2;--border2:#c8c3b8;--text:#1a1916;--text2:#4a4740;
      --muted:#8c897f;--accent:#1a3a2a;--accent-light:#2d6348;--accent-pale:#e8f0eb;
      --gold:#b8903a;--gold-pale:#f5edd8;--red:#8b2e2e;--row-hover:#f0ede6;
      --shadow-sm:0 1px 3px rgba(26,25,22,0.08);--shadow-md:0 4px 16px rgba(26,25,22,0.10);
      --radius:4px;--radius-lg:8px;
    }
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
    html{scroll-behavior:smooth;}
    body{font-family:Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden;}

    .layout{display:flex;min-height:100vh;}
    .sidebar{width:260px;flex-shrink:0;background:var(--accent);display:flex;flex-direction:column;position:fixed;top:0;left:0;bottom:0;z-index:200;}
    .sidebar-logo{padding:32px 28px 28px;border-bottom:1px solid rgba(255,255,255,0.1);}
    .logo-mark{font-family:Arial,sans-serif;font-size:1.5rem;font-weight:700;color:#fff;letter-spacing:0.02em;margin-bottom:2px;}
    .logo-sub{font-size:0.65rem;font-weight:500;letter-spacing:0.18em;text-transform:uppercase;color:rgba(255,255,255,0.45);}
    .sidebar-nav{padding:20px 16px;flex:1;}
    .nav-section-label{font-size:0.6rem;font-weight:600;letter-spacing:0.2em;text-transform:uppercase;color:rgba(255,255,255,0.35);padding:0 12px;margin-bottom:8px;margin-top:20px;}
    .nav-item{display:flex;align-items:center;gap:12px;padding:11px 12px;border-radius:var(--radius-lg);color:rgba(255,255,255,0.65);font-size:0.875rem;font-weight:400;cursor:pointer;transition:all 0.18s;border:none;background:none;width:100%;text-align:left;text-decoration:none;}
    .nav-item:hover{background:rgba(255,255,255,0.08);color:#fff;}
    .nav-item.active{background:rgba(255,255,255,0.14);color:#fff;font-weight:500;}
    .nav-item svg{width:16px;height:16px;stroke:currentColor;fill:none;stroke-width:1.8;flex-shrink:0;}
    .nav-count{margin-left:auto;font-family:Arial,sans-serif;font-size:0.7rem;background:rgba(255,255,255,0.1);border-radius:20px;padding:2px 8px;color:rgba(255,255,255,0.55);}
    .sidebar-footer{padding:20px 28px;border-top:1px solid rgba(255,255,255,0.08);font-size:0.72rem;color:rgba(255,255,255,0.3);line-height:1.6;}

    .main{margin-left:260px;flex:1;display:flex;flex-direction:column;min-height:100vh;}
    .topbar{height:60px;background:var(--surface);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;padding:0 36px;position:sticky;top:0;z-index:100;}
    .topbar-title{font-family:Arial,sans-serif;font-size:1.05rem;font-weight:600;color:var(--text);}
    .topbar-meta{display:flex;align-items:center;gap:20px;}
    .topbar-date{font-size:0.75rem;color:var(--muted);font-family:Arial,sans-serif;}
    .status-dot{display:flex;align-items:center;gap:7px;font-size:0.72rem;color:var(--muted);}
    .dot-live{width:7px;height:7px;border-radius:50%;background:#2d9c5a;box-shadow:0 0 0 2px #d1f0dc;animation:blink 2.5s ease-in-out infinite;}
    @keyframes blink{0%,100%{opacity:1;}50%{opacity:0.4;}}

    .content{padding:36px;flex:1;}
    .page-header{margin-bottom:28px;}
    .page-header h1{font-family:Arial,sans-serif;font-size:2rem;font-weight:700;color:var(--text);margin-bottom:8px;}
    .page-header p{font-size:0.82rem;color:var(--muted);}

    .search-row{display:flex;gap:12px;margin-bottom:20px;align-items:center;}
    .search-wrap{position:relative;flex:1;}
    .search-wrap svg.ico{position:absolute;left:14px;top:50%;transform:translateY(-50%);width:16px;height:16px;stroke:var(--muted);fill:none;stroke-width:2;pointer-events:none;}
    .search-input{width:100%;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:11px 14px 11px 42px;color:var(--text);font-family:Arial,sans-serif;font-size:0.875rem;transition:all 0.2s;outline:none;box-shadow:var(--shadow-sm);}
    .search-input::placeholder{color:var(--muted);}
    .search-input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(26,58,42,0.08);}

    .tab-bar{display:flex;gap:0;margin-bottom:20px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow-sm);width:fit-content;flex-wrap:wrap;}
    .tab-btn{display:flex;align-items:center;gap:8px;padding:10px 20px;font-size:0.8rem;font-weight:500;color:var(--muted);background:none;border:none;border-right:1px solid var(--border);cursor:pointer;transition:all 0.18s;white-space:nowrap;}
    .tab-btn:last-child{border-right:none;}
    .tab-btn:hover{background:var(--bg2);color:var(--text);}
    .tab-btn.active{background:var(--accent);color:#fff;}
    .tab-num{font-family:Arial,sans-serif;font-size:0.68rem;background:rgba(255,255,255,0.18);border-radius:20px;padding:1px 7px;}
    .tab-btn:not(.active) .tab-num{background:rgba(26,25,22,0.07);}

    .table-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);overflow:hidden;box-shadow:var(--shadow-sm);}
    .table-header{display:flex;align-items:center;justify-content:space-between;padding:14px 20px;border-bottom:1px solid var(--border);background:var(--surface2);}
    .count-pill{font-family:Arial,sans-serif;font-size:0.72rem;background:var(--accent-pale);border:1px solid rgba(26,58,42,0.15);color:var(--accent);border-radius:20px;padding:3px 11px;font-weight:500;}
    .col-header{display:grid;padding:10px 20px;border-bottom:1px solid var(--border);background:var(--surface2);}
    .col-header.g6{grid-template-columns:1.8fr 2.5fr 1.4fr 0.7fr 0.9fr 0.9fr;}
    .col-head{font-size:0.65rem;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:var(--muted);}
    .list-scroll{overflow-y:auto;max-height:540px;}
    .list-scroll::-webkit-scrollbar{width:4px;}
    .list-scroll::-webkit-scrollbar-track{background:transparent;}
    .list-scroll::-webkit-scrollbar-thumb{background:var(--border2);border-radius:10px;}
    .list-row{display:grid;padding:12px 20px;align-items:center;border-bottom:1px solid rgba(26,25,22,0.04);transition:background 0.12s;cursor:default;}
    .list-row.g6{grid-template-columns:1.8fr 2.5fr 1.4fr 0.7fr 0.9fr 0.9fr;}
    .list-row:last-child{border-bottom:none;}
    .list-row:hover{background:var(--row-hover);}
    .cell-svc{font-family:Arial,sans-serif;font-size:0.82rem;color:var(--accent-light);font-weight:500;}
    .cell-cat{font-size:0.83rem;color:var(--text2);}
    .cell-qty,.cell-cost{font-family:Arial,sans-serif;font-size:0.82rem;color:var(--text2);text-align:right;}

    .cat-badge{display:inline-flex;padding:2px 9px;border-radius:var(--radius);font-size:0.6rem;font-weight:600;letter-spacing:0.04em;text-transform:uppercase;white-space:nowrap;}
    .cat-combustiveis{background:#fde8e8;color:#8b2e2e;border:1px solid rgba(139,46,46,0.18);}
    .cat-lubrificantes{background:#fef3e0;color:#8b6b2e;border:1px solid rgba(139,107,46,0.18);}
    .cat-pecas{background:#e3ecf6;color:#2e558b;border:1px solid rgba(46,85,139,0.18);}
    .cat-componentes{background:#ece3f6;color:#5b2e8b;border:1px solid rgba(91,46,139,0.18);}
    .cat-administrativos{background:#eceae4;color:#4a4740;border:1px solid rgba(74,71,64,0.18);}
    .cat-epis{background:var(--accent-pale);color:var(--accent-light);border:1px solid rgba(26,58,42,0.15);}
    .cat-insumos{background:#e0f0ef;color:#1a6b63;border:1px solid rgba(26,107,99,0.18);}
    .cat-outros{background:#f0f0f0;color:#666;border:1px solid rgba(0,0,0,0.08);}

    .btn-copy{opacity:0;margin-left:10px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:2px 8px;font-size:0.65rem;color:var(--muted);cursor:pointer;transition:all 0.15s;font-family:Arial,sans-serif;}
    .btn-copy:hover{border-color:var(--accent);color:var(--accent);}
    .list-row:hover .btn-copy{opacity:1;}

    mark{background:rgba(184,144,58,0.2);color:var(--gold);border-radius:2px;padding:0 2px;}

    .table-footer{display:flex;align-items:center;justify-content:space-between;padding:12px 20px;border-top:1px solid var(--border);background:var(--surface2);flex-wrap:wrap;gap:8px;}
    .footer-info{font-family:Arial,sans-serif;font-size:0.72rem;color:var(--muted);}
    .pagination{display:flex;align-items:center;gap:3px;}
    .pg-btn{min-width:30px;height:30px;border-radius:var(--radius);background:var(--surface);border:1px solid var(--border);color:var(--muted);font-size:0.75rem;cursor:pointer;transition:all 0.15s;font-family:Arial,sans-serif;display:grid;place-items:center;}
    .pg-btn:hover:not(:disabled){border-color:var(--accent);color:var(--accent);}
    .pg-btn.active{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600;}
    .pg-btn:disabled{opacity:0.3;cursor:not-allowed;}
    .pg-dots{color:var(--muted);font-size:0.75rem;padding:0 3px;}

    .empty-state{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:60px 20px;gap:12px;color:var(--muted);text-align:center;}
    .empty-state svg{width:40px;height:40px;stroke:var(--border2);fill:none;stroke-width:1.5;}
    .empty-state strong{font-size:0.9rem;color:var(--text2);}
    .empty-state span{font-size:0.8rem;}

    @media(max-width:900px){
      .sidebar{display:none;}.main{margin-left:0;}.col-header.g6{display:none;}
      .list-row.g6{grid-template-columns:1fr;gap:5px;}.content{padding:20px;}
    }
  </style>
</head>
<body>
<div class="layout">
  <aside class="sidebar">
    <div class="sidebar-logo">
      <div class="logo-mark">Extensão Oracle</div>
      <div class="logo-sub">Sistema para consultas</div>
    </div>
    <nav class="sidebar-nav">
      <div class="nav-section-label">Modulos</div>
      <a class="nav-item" href="/" style="text-decoration:none;">
        <svg viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
        Visao Geral
      </a>
      <a class="nav-item active" href="/estoque" style="text-decoration:none;">
        <svg viewBox="0 0 24 24"><path d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4"/></svg>
        Estoque
        <span class="nav-count">{{ estoque_count }}</span>
      </a>
      <div style="font-size:0.6rem;font-weight:600;letter-spacing:0.2em;text-transform:uppercase;color:rgba(255,255,255,0.35);padding:0 12px;margin:16px 0 8px;">Ferramenta de Gestão</div>
      <a class="nav-item" href="/login" style="text-decoration:none;">
        <svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
        Gestão (Login)
      </a>
    </nav>
    <div class="sidebar-footer">
      Base atualizada<br>
      v2.0 - Essencis Sistemas
    </div>
  </aside>

  <main class="main">
    <div class="topbar">
      <span class="topbar-title">Estoque Almoxarifado</span>
      <div class="topbar-meta">
        <div class="status-dot"><div class="dot-live"></div>Sistema ativo</div>
        <div class="topbar-date" id="current-date"></div>
      </div>
    </div>

    <div class="content">
      <div class="page-header">
        <h1>Estoque Almoxarifado</h1>
        <p>Estoque periodico - Organizacao 00301_EMG_BETIM</p>
      </div>

      <div class="tab-bar" id="est-tab-bar">
        <button class="tab-btn active" onclick="switchEstTab('')">Todos <span class="tab-num" id="num-est-todos">0</span></button>
        <button class="tab-btn" onclick="switchEstTab('LUBRIFICANTES')">Lubrificantes <span class="tab-num" id="num-est-lubrificantes">0</span></button>
        <button class="tab-btn" onclick="switchEstTab('PEÇAS DE FROTA')">Pecas de Frota <span class="tab-num" id="num-est-pecas">0</span></button>
        <button class="tab-btn" onclick="switchEstTab('COMPONENTES E IMPLEMENTOS')">Componentes <span class="tab-num" id="num-est-componentes">0</span></button>
        <button class="tab-btn" onclick="switchEstTab('MATERIAIS ADMINISTRATIVOS')">Mat. Administrativos <span class="tab-num" id="num-est-administrativos">0</span></button>
        <button class="tab-btn" onclick="switchEstTab('EPIS/UNIFORMES')">EPIs / Uniformes <span class="tab-num" id="num-est-epis">0</span></button>
        <button class="tab-btn" onclick="switchEstTab('INSUMOS')">Insumos <span class="tab-num" id="num-est-insumos">0</span></button>
        <button class="tab-btn" onclick="switchEstTab('OUTROS')">Outros <span class="tab-num" id="num-est-outros">0</span></button>
      </div>

      <div class="search-row">
        <div class="search-wrap">
          <svg class="ico" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
          <input class="search-input" id="search-est" type="text"
                 placeholder="Buscar por nome do item, descricao ou categoria..."
                 oninput="onSearchEst(this.value)">
        </div>
      </div>

      <div class="table-card">
        <div class="table-header">
          <div></div>
          <span class="count-pill" id="est-badge">0 itens</span>
        </div>
        <div class="col-header g6">
          <span class="col-head">Nome do Item</span>
          <span class="col-head">Descricao</span>
          <span class="col-head">Categoria</span>
          <span class="col-head" style="text-align:right;">Qtd</span>
          <span class="col-head" style="text-align:right;">Custo Unit.</span>
          <span class="col-head" style="text-align:right;">Custo Total</span>
        </div>
        <div class="list-scroll" id="est-list"></div>
        <div class="table-footer">
          <span class="footer-info" id="est-footer"></span>
          <div class="pagination" id="est-pagination"></div>
        </div>
        <div style="text-align:right;padding:6px 16px;font-size:0.75rem;color:var(--muted);border-top:1px solid var(--border);">
          Ultima atualizacao da planilha: {{ file_dates['estoque'] }}
        </div>
      </div>
    </div>
  </main>
</div>

<script>
var EST_DATA = ({{ estoque_data | tojson | safe }}).filter(function(r){return r.Categoria!=='COMBUSTÍVEIS';});
var PER_PAGE = 20;
var state = { filtered: EST_DATA.slice(), page: 1, query: '', catFilter: '' };

var CAT_CSS = {
  'LUBRIFICANTES':'cat-lubrificantes',
  'PEÇAS DE FROTA':'cat-pecas','COMPONENTES E IMPLEMENTOS':'cat-componentes',
  'MATERIAIS ADMINISTRATIVOS':'cat-administrativos','EPIS/UNIFORMES':'cat-epis',
  'INSUMOS':'cat-insumos','OUTROS':'cat-outros'
};

window.addEventListener('DOMContentLoaded', function() {
  var d = new Date();
  var el = document.getElementById('current-date');
  if (el) el.textContent = d.toLocaleDateString('pt-BR', {weekday:'short',day:'2-digit',month:'short',year:'numeric'});
  updateTabCounts();
  render();
});

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/'/g,'&#39;').replace(/"/g,'&quot;');}
function jsesc(s){return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'\\"').replace(/</g,'\\x3C').replace(/>/g,'\\x3E');}
function highlight(t,q){var s=esc(t);if(!q||q.length<2)return s;var re=new RegExp('('+q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi');return s.replace(re,'<mark>$1</mark>');}
function fmtNum(v){var n=parseFloat(v);if(isNaN(n))return '—';return n.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});}

function copyText(text,btn){
  navigator.clipboard.writeText(text).then(function(){
    var o=btn.textContent;btn.textContent='copiado';btn.style.color='var(--accent)';btn.style.borderColor='var(--accent)';
    setTimeout(function(){btn.textContent=o;btn.style.color='';btn.style.borderColor='';},1500);
  }).catch(function(){});
}

function filterData(){
  var q=state.query, cat=state.catFilter, base=EST_DATA;
  if(cat) base=base.filter(function(r){return r.Categoria===cat;});
  if(!q){state.filtered=base.slice();}
  else{var lo=q.toLowerCase();state.filtered=base.filter(function(r){
    return String(r['Nome do Item']||'').toLowerCase().indexOf(lo)!==-1||
           String(r['Descrição do Item']||'').toLowerCase().indexOf(lo)!==-1||
           String(r.Categoria||'').toLowerCase().indexOf(lo)!==-1;
  });}
  state.page=1;render();
}

function render(){
  var f=state.filtered,p=state.page,q=state.query;
  var total=f.length,start=(p-1)*PER_PAGE,end=Math.min(start+PER_PAGE,total);
  var slice=f.slice(start,end),el=document.getElementById('est-list');
  document.getElementById('est-badge').textContent=total+' itens';
  if(!slice.length){
    el.innerHTML='<div class="empty-state"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg><strong>Nenhum item encontrado</strong>'+(q?'<span>para "'+esc(q)+'"</span>':'')+'</div>';
    document.getElementById('est-footer').textContent='0 resultados';
    document.getElementById('est-pagination').innerHTML='';return;
  }
  var html='';
  for(var i=0;i<slice.length;i++){
    var it=slice[i];
    var nome=String(it['Nome do Item']||'').trim()||'—';
    var desc=String(it['Descrição do Item']||'').trim()||'—';
    var cat=String(it.Categoria||'').trim()||'OUTROS';
    var cls=CAT_CSS[cat]||'cat-outros';
    html+='<div class="list-row g6">'+
      '<div class="cell-svc">'+highlight(nome,q)+'<button class="btn-copy" onclick="copyText(\''+jsesc(nome)+'\',this)">copiar</button></div>'+
      '<div class="cell-cat" style="font-size:0.8rem;">'+highlight(desc,q)+'</div>'+
      '<div><span class="cat-badge '+cls+'">'+esc(cat)+'</span></div>'+
      '<div class="cell-qty">'+fmtNum(it['Quantidade'])+'</div>'+
      '<div class="cell-cost">'+fmtNum(it['Custo Unitário'])+'</div>'+
      '<div class="cell-cost">'+fmtNum(it['Custo Total'])+'</div>'+
    '</div>';
  }
  el.innerHTML=html;
  document.getElementById('est-footer').textContent='Exibindo '+(start+1)+'–'+end+' de '+total+' registros';
  renderPag(total,p);
}

function onSearchEst(q){state.query=q.trim();filterData();}

function switchEstTab(cat){
  state.catFilter=cat;state.page=1;
  var bar=document.getElementById('est-tab-bar');
  var btns=bar.querySelectorAll('.tab-btn');
  var cats=['','LUBRIFICANTES','PEÇAS DE FROTA','COMPONENTES E IMPLEMENTOS','MATERIAIS ADMINISTRATIVOS','EPIS/UNIFORMES','INSUMOS','OUTROS'];
  for(var i=0;i<btns.length;i++) btns[i].className='tab-btn'+(cats[i]===cat?' active':'');
  filterData();
}

function updateTabCounts(){
  var m={};for(var i=0;i<EST_DATA.length;i++){var c=EST_DATA[i].Categoria||'OUTROS';m[c]=(m[c]||0)+1;}
  var ids={'':'num-est-todos','LUBRIFICANTES':'num-est-lubrificantes','PEÇAS DE FROTA':'num-est-pecas','COMPONENTES E IMPLEMENTOS':'num-est-componentes','MATERIAIS ADMINISTRATIVOS':'num-est-administrativos','EPIS/UNIFORMES':'num-est-epis','INSUMOS':'num-est-insumos','OUTROS':'num-est-outros'};
  for(var k in ids){var e=document.getElementById(ids[k]);if(e) e.textContent=(k===''?EST_DATA.length:(m[k]||0));}
}

function goPage(p){state.page=p;render();}

function renderPag(total,cur){
  var pages=Math.max(1,Math.ceil(total/PER_PAGE)),el=document.getElementById('est-pagination');
  if(!el||pages<=1){if(el)el.innerHTML='';return;}
  var h='<button class="pg-btn" onclick="goPage('+(cur-1)+')" '+(cur===1?'disabled':'')+'>‹</button>';
  var D=2,lo=Math.max(1,cur-D),hi=Math.min(pages,cur+D);
  if(lo>1){h+=pb(1,cur);if(lo>2)h+='<span class="pg-dots">…</span>';}
  for(var i=lo;i<=hi;i++)h+=pb(i,cur);
  if(hi<pages){if(hi<pages-1)h+='<span class="pg-dots">…</span>';h+=pb(pages,cur);}
  h+='<button class="pg-btn" onclick="goPage('+(cur+1)+')" '+(cur===pages?'disabled':'')+'>›</button>';
  el.innerHTML=h;
}
function pb(p,c){return '<button class="pg-btn'+(p===c?' active':'')+'" onclick="goPage('+p+')">'+p+'</button>';}
</script>
</body>
</html>'''

@app.route('/estoque/')
@app.route('/estoque')
def estoque():
    return render_template_string(HTML_ESTOQUE,
        estoque_data  = estoque_data,
        estoque_count = len(estoque_data),
        file_dates    = file_dates)

# ═══════════════════════════════════════════════════════════════════
# SIDEBAR DA GESTÃO — helper compartilhado
# ═══════════════════════════════════════════════════════════════════
def _gestao_sidebar(active='dashboard'):
    nav_items = [
        ('dashboard',    '/',             'Visão Geral',           '<path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>'),
        ('materiais',    '/',             'Materiais',             '<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/>'),
        ('servicos',     '/',             'Serviços',              '<path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="2"/>'),
        ('estoque',      '/estoque',      'Estoque',               '<path d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4"/>'),
    ]
    gestao_items = [
        ('gestao',           '/gestao',           'Dashboard Gestão',   '<path d="M12 20V10"/><path d="M18 20V4"/><path d="M6 20v-4"/>'),
        ('gestao_estoque',   '/gestao/estoque',   'Estoque Análise',    '<path d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4"/>'),
        ('gestao_pagamentos','/gestao/pagamentos','Pagamentos',         '<path d="M12 1v22M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>'),
        ('gestao_ordens',    '/gestao/ordens',    'Ordens de Compra',   '<path d="M6 2 3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/>'),
        ('gestao_precos',    '/gestao/precos',    'Análise de Preços',  '<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>'),
        ('gestao_pipefy',    '/gestao/pipefy',    'Atividades Pipefy',  '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>'),
    ]
    html = '''<aside class="sidebar">
      <div class="sidebar-logo">
        <div class="logo-mark">EXTENSÃO ORACLE ERP</div>
        <div class="logo-sub">Sistema para consultas</div>
      </div>
      <nav class="sidebar-nav">
        <div class="nav-section-label">Módulos Públicos</div>'''
    for key, href, label, svg_path in nav_items:
        cls = 'nav-item active' if active == key else 'nav-item'
        html += f'<a class="{cls}" href="{href}" style="text-decoration:none;"><svg viewBox="0 0 24 24">{svg_path}</svg>{label}</a>'
    html += '<div class="nav-section-label" style="margin-top:20px;">Ferramenta de Gestão</div>'
    if session.get('logged_in'):
        for key, href, label, svg_path in gestao_items:
            cls = 'nav-item active' if active == key else 'nav-item'
            html += f'<a class="{cls}" href="{href}" style="text-decoration:none;"><svg viewBox="0 0 24 24">{svg_path}</svg>{label}</a>'
        html += '<a class="nav-item" href="/logout" style="text-decoration:none;color:rgba(255,100,100,0.8);"><svg viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>Sair</a>'
    else:
        html += '<a class="nav-item" href="/login" style="text-decoration:none;"><svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>Login Gestão</a>'
    html += '''
      </nav>
      <div class="sidebar-footer">Base atualizada<br>v3.0 · Essencis Sistemas</div>
    </aside>'''
    return html


# ─── SIDEBAR CSS (mesmo do flask_app original) ────────────────────────────────
_SIDEBAR_CSS = '''
    :root {
      --bg:#f4f3ef;--bg2:#eceae4;--surface:#ffffff;--surface2:#f9f8f5;
      --border:#dedad2;--border2:#c8c3b8;--text:#1a1916;--text2:#4a4740;
      --muted:#8c897f;--accent:#1a3a2a;--accent-light:#2d6348;--accent-pale:#e8f0eb;
      --gold:#b8903a;--gold-pale:#f5edd8;--red:#8b2e2e;--row-hover:#f0ede6;
      --shadow-sm:0 1px 3px rgba(26,25,22,0.08);--shadow-md:0 4px 16px rgba(26,25,22,0.10);
      --radius:4px;--radius-lg:8px;
    }
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
    html{scroll-behavior:smooth;}
    body{font-family:Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden;}
    .layout{display:flex;min-height:100vh;}
    .sidebar{width:260px;flex-shrink:0;background:var(--accent);display:flex;flex-direction:column;position:fixed;top:0;left:0;bottom:0;z-index:200;overflow-y:auto;}
    .sidebar-logo{padding:32px 28px 28px;border-bottom:1px solid rgba(255,255,255,0.1);}
    .logo-mark{font-family:Arial,sans-serif;font-size:1.5rem;font-weight:700;color:#fff;letter-spacing:0.02em;margin-bottom:2px;}
    .logo-sub{font-size:0.65rem;font-weight:500;letter-spacing:0.18em;text-transform:uppercase;color:rgba(255,255,255,0.45);}
    .sidebar-nav{padding:20px 16px;flex:1;}
    .nav-section-label{font-size:0.6rem;font-weight:600;letter-spacing:0.2em;text-transform:uppercase;color:rgba(255,255,255,0.35);padding:0 12px;margin-bottom:8px;margin-top:20px;}
    .nav-item{display:flex;align-items:center;gap:12px;padding:11px 12px;border-radius:var(--radius-lg);color:rgba(255,255,255,0.65);font-size:0.875rem;font-weight:400;cursor:pointer;transition:all 0.18s;border:none;background:none;width:100%;text-align:left;text-decoration:none;}
    .nav-item:hover{background:rgba(255,255,255,0.08);color:#fff;}
    .nav-item.active{background:rgba(255,255,255,0.14);color:#fff;font-weight:500;}
    .nav-item svg{width:16px;height:16px;stroke:currentColor;fill:none;stroke-width:1.8;flex-shrink:0;}
    .sidebar-footer{padding:20px 28px;border-top:1px solid rgba(255,255,255,0.08);font-size:0.72rem;color:rgba(255,255,255,0.3);line-height:1.6;}
    .main{margin-left:260px;flex:1;display:flex;flex-direction:column;min-height:100vh;}
    .topbar{height:60px;background:var(--surface);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;padding:0 36px;position:sticky;top:0;z-index:100;}
    .topbar-title{font-family:Arial,sans-serif;font-size:1.05rem;font-weight:600;color:var(--text);}
    .topbar-meta{display:flex;align-items:center;gap:20px;}
    .topbar-date{font-size:0.75rem;color:var(--muted);font-family:Arial,sans-serif;}
    .status-dot{display:flex;align-items:center;gap:7px;font-size:0.72rem;color:var(--muted);}
    .dot-live{width:7px;height:7px;border-radius:50%;background:#2d9c5a;box-shadow:0 0 0 2px #d1f0dc;animation:blink 2.5s ease-in-out infinite;}
    @keyframes blink{0%,100%{opacity:1;}50%{opacity:0.4;}}
    .content{padding:36px;flex:1;}
    @media(max-width:900px){.sidebar{display:none;}.main{margin-left:0;}.content{padding:20px;}}
'''

# ─── GESTÃO CSS adicional (Inter + Chart.js style) ────────────────────────────
_GESTAO_CSS = '''
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    .g-container{font-family:"Inter",Arial,sans-serif;}
    .g-header{background:linear-gradient(135deg,#1a3a2a 0%,#2d6348 100%);border-radius:8px;padding:28px 32px;margin-bottom:24px;color:white;box-shadow:0 4px 16px rgba(26,25,22,0.12);}
    .g-header h1{font-size:1.8rem;font-weight:800;margin-bottom:6px;}
    .g-header p{opacity:0.85;font-size:0.9rem;}
    .g-kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;margin-bottom:24px;}
    .g-kpi{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:20px 22px;box-shadow:var(--shadow-sm);border-left:4px solid var(--accent);}
    .g-kpi.success{border-left-color:#2d9c5a;}
    .g-kpi.warning{border-left-color:#b8903a;}
    .g-kpi.danger{border-left-color:#8b2e2e;}
    .g-kpi-label{font-size:0.68rem;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;color:var(--muted);margin-bottom:8px;}
    .g-kpi-value{font-size:1.9rem;font-weight:800;color:var(--text);line-height:1.1;}
    .g-kpi-sub{font-size:0.75rem;color:var(--muted);margin-top:4px;}
    .g-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:24px;box-shadow:var(--shadow-sm);margin-bottom:20px;}
    .g-card h3{font-size:1rem;font-weight:700;color:var(--text);margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid var(--border);}
    .g-grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:20px;margin-bottom:20px;}
    .g-grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:20px;}
    .g-chart-wrap{position:relative;height:280px;}
    .g-table{width:100%;border-collapse:collapse;font-size:0.84rem;}
    .g-table th{padding:10px 14px;text-align:left;font-size:0.7rem;font-weight:600;letter-spacing:0.1em;text-transform:uppercase;color:var(--muted);border-bottom:2px solid var(--border);background:var(--surface2);}
    .g-table td{padding:10px 14px;border-bottom:1px solid var(--border);}
    .g-table tbody tr:hover{background:var(--row-hover);}
    .g-badge{display:inline-flex;padding:2px 8px;border-radius:var(--radius);font-size:0.68rem;font-weight:600;text-transform:uppercase;}
    .g-badge.a{background:#fde8e8;color:#8b2e2e;}
    .g-badge.b{background:#fef3e0;color:#8b6b2e;}
    .g-badge.c{background:#e8f0eb;color:#2d6348;}
    .g-tab-bar{display:flex;gap:0;margin-bottom:20px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow-sm);width:fit-content;flex-wrap:wrap;}
    .g-tab{padding:10px 18px;font-size:0.8rem;font-weight:500;color:var(--muted);background:none;border:none;border-right:1px solid var(--border);cursor:pointer;transition:all 0.18s;white-space:nowrap;}
    .g-tab:last-child{border-right:none;}
    .g-tab:hover{background:var(--bg2);color:var(--text);}
    .g-tab.active{background:var(--accent);color:#fff;}
    .g-filter{padding:10px 14px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);font-size:0.84rem;color:var(--text);outline:none;}
    .g-filter:focus{border-color:var(--accent);}
    .g-no-data{text-align:center;padding:60px 20px;color:var(--muted);}
    .g-no-data .icon{font-size:3rem;margin-bottom:12px;}
    @media(max-width:900px){.g-grid2,.g-grid3{grid-template-columns:1fr;}}
'''

def _gestao_base(title, active, topbar_title, content_html):
    """Monta página completa no layout do flask_app."""
    import datetime as dt
    date_str = dt.datetime.now().strftime('%d/%m/%Y %H:%M')
    return render_template_string(f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>{title} · Oracle ERP</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
  <style>
    {_SIDEBAR_CSS}
    {_GESTAO_CSS}
  </style>
</head>
<body>
<div class="layout">
  {_gestao_sidebar(active)}
  <main class="main">
    <div class="topbar">
      <span class="topbar-title">{topbar_title}</span>
      <div class="topbar-meta">
        <div class="status-dot"><div class="dot-live"></div>Sistema ativo</div>
        <div class="topbar-date">{date_str}</div>
        {'<a href="/logout" style="font-size:0.78rem;color:var(--muted);text-decoration:none;padding:6px 12px;border:1px solid var(--border);border-radius:4px;">Sair</a>' if session.get('logged_in') else ''}
      </div>
    </div>
    <div class="content g-container">
{content_html}
    </div>
  </main>
</div>
</body>
</html>''')


# ═══════════════════════════════════════════════════════════════════
# ATUALIZAR sidebar nas rotas públicas existentes — injetar nav de gestão
# ═══════════════════════════════════════════════════════════════════

# ─── LOGIN / LOGOUT ──────────────────────────────────────────────────────────
LOGIN_HTML = '''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Login · Ferramenta de Gestão</title>
  <style>
    :root{--accent:#1a3a2a;--accent-light:#2d6348;}
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:Arial,sans-serif;background:#f4f3ef;min-height:100vh;display:flex;align-items:center;justify-content:center;}
    .login-wrap{background:#fff;border-radius:12px;padding:48px 40px;width:100%;max-width:420px;box-shadow:0 12px 40px rgba(26,25,22,0.14);}
    .login-logo{text-align:center;margin-bottom:32px;}
    .login-logo .mark{font-size:1.4rem;font-weight:700;color:var(--accent);}
    .login-logo .sub{font-size:0.7rem;letter-spacing:0.18em;text-transform:uppercase;color:#8c897f;}
    h2{font-size:1.3rem;font-weight:700;color:#1a1916;margin-bottom:6px;text-align:center;}
    .subtitle{text-align:center;font-size:0.82rem;color:#8c897f;margin-bottom:28px;}
    label{display:block;font-size:0.78rem;font-weight:600;color:#4a4740;margin-bottom:5px;letter-spacing:0.06em;text-transform:uppercase;}
    input{width:100%;padding:11px 14px;border:1px solid #dedad2;border-radius:4px;font-size:0.875rem;outline:none;transition:border-color 0.2s;margin-bottom:18px;}
    input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(26,58,42,0.08);}
    .btn-login{width:100%;background:var(--accent);color:#fff;border:none;border-radius:4px;padding:13px;font-size:0.9rem;font-weight:600;cursor:pointer;transition:background 0.2s;}
    .btn-login:hover{background:var(--accent-light);}
    .error{color:#8b2e2e;background:#fde8e8;border:1px solid rgba(139,46,46,0.2);border-radius:4px;padding:10px 14px;font-size:0.82rem;margin-bottom:16px;text-align:center;}
    .back-link{display:block;text-align:center;margin-top:20px;font-size:0.8rem;color:#8c897f;text-decoration:none;}
    .back-link:hover{color:var(--accent);}
  </style>
</head>
<body>
  <div class="login-wrap">
    <div class="login-logo">
      <div class="mark">EXTENSÃO ORACLE ERP</div>
      <div class="sub">Ferramenta de Gestão</div>
    </div>
    <h2>Acesso Restrito</h2>
    <p class="subtitle">Somente usuários autorizados</p>
    {% if error %}<div class="error">{{ error }}</div>{% endif %}
    <form method="post">
      <label>Login</label>
      <input type="text" name="login" placeholder="admin" autocomplete="username" required>
      <label>Senha</label>
      <input type="password" name="senha" placeholder="••••••" autocomplete="current-password" required>
      <button type="submit" class="btn-login">Entrar</button>
    </form>
    <a href="/" class="back-link">← Voltar ao painel público</a>
  </div>
</body>
</html>'''

@app.route('/login', methods=['GET', 'POST'])
def login_page():
    if session.get('logged_in'):
        return redirect(url_for('gestao_dashboard'))
    error = None
    if request.method == 'POST':
        login_val = request.form.get('login', '').strip()
        senha_val = request.form.get('senha', '').strip()
        if login_val == ADMIN_LOGIN and senha_val == ADMIN_SENHA:
            session['logged_in'] = True
            session['user'] = login_val
            return redirect(url_for('gestao_dashboard'))
        else:
            error = 'Login ou senha incorretos.'
    return render_template_string(LOGIN_HTML, error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))


# ─── HELPERS PARA DADOS DA GESTÃO ────────────────────────────────────────────
def _load_gestao_file(filename_glob, base=BASE_GESTAO):
    """Carrega primeiro Excel encontrado pelo padrão glob."""
    import glob as _glob
    mes_atual, ano_atual = datetime.datetime.now().strftime('%m'), datetime.datetime.now().strftime('%Y')
    patterns = [
        os.path.join(base, 'Estoque periodico', ano_atual, mes_atual, '*.xlsx'),
        os.path.join(base, 'Estoque periodico', ano_atual, mes_atual, '*.xls'),
        os.path.join(base, filename_glob),
    ]
    for p in patterns:
        matches = _glob.glob(p)
        if matches:
            return matches[0]
    return None

def _find_latest_excel(subfolder):
    """Encontra o Excel mais recente em Relatórios/subfolder/ANO/MES/."""
    base = r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios"
    pattern = os.path.join(base, subfolder, '*', '*', '*.xls*')
    files = glob.glob(pattern)
    if not files:
        # Fallback: procurar no diretório da ferramenta
        pattern2 = os.path.join(BASE_GESTAO, '**', '*.xlsx')
        files = glob.glob(pattern2, recursive=True)
        files = [f for f in files if subfolder.split()[0].lower() in f.lower()]
    if files:
        return max(files, key=os.path.getmtime)
    return None

def _fmt_brl(val):
    try:
        return f"R$ {float(val):,.2f}".replace(',','X').replace('.',',').replace('X','.')
    except Exception:
        return "R$ 0,00"

def _fmt_n(val):
    try:
        return f"{float(val):,.0f}".replace(',','.')
    except Exception:
        return "0"

CATEGORIA_MAP = {
    '01.01.01':'COMBUSTÍVEIS','01.01.02':'LUBRIFICANTES','01.01.03':'PEÇAS DE FROTA',
    '01.01.04':'COMPONENTES E IMPLEMENTOS','01.01.05':'COMPONENTES E IMPLEMENTOS',
    '01.01.06':'MATERIAIS ADMINISTRATIVOS','01.01.07':'MATERIAIS ADMINISTRATIVOS',
    '01.01.08':'EPIS/UNIFORMES','01.01.09':'INSUMOS','01.01.10':'COMPONENTES E IMPLEMENTOS',
    '01.01.11':'MATERIAIS ADMINISTRATIVOS','01.01.12':'INSUMOS','01.01.13':'INSUMOS',
    '01.01.14':'INSUMOS','01.01.15':'INSUMOS','01.01.16':'INSUMOS','01.01.17':'INSUMOS',
    '01.01.18':'INSUMOS','01.01.19':'INSUMOS','01.01.20':'INSUMOS',
    '01.01.21':'MATERIAIS ADMINISTRATIVOS','01.01.22':'MATERIAIS ADMINISTRATIVOS',
    '01.01.23':'MATERIAIS ADMINISTRATIVOS',
}

def _classify_item(nome_item):
    s = str(nome_item).strip()
    for prefix in sorted(CATEGORIA_MAP.keys(), key=len, reverse=True):
        if s.startswith(prefix):
            return CATEGORIA_MAP[prefix]
    return 'OUTROS'

def _load_estoque_df():
    """Carrega DataFrame de estoque periódico."""
    path = _find_latest_excel('Estoque periodico')
    if not path:
        return None
    try:
        df = pd.read_excel(path, engine='openpyxl')
        df.columns = df.columns.str.strip()
        org_col = next((c for c in df.columns if 'Organiz' in c and 'Nome' not in c), None)
        name_col = next((c for c in df.columns if 'Organiz' in c and 'Nome' in c), None)
        sub_col  = next((c for c in df.columns if 'Subinvent' in c and 'Nome' in c), None)
        item_col = next((c for c in df.columns if c.strip() == 'Nome do Item'), None)
        if org_col:
            df = df[df[org_col].astype(str).str.strip() == '00301_EMG_BETIM']
        elif name_col:
            df = df[df[name_col].astype(str).str.strip() == 'EMG BETIM']
        if sub_col:
            df = df[df[sub_col].astype(str).str.strip().isin(['MAT1','MAT3','TNQ1'])]
        if item_col:
            df['Categoria'] = df[item_col].apply(_classify_item)
        else:
            df['Categoria'] = 'OUTROS'
        return df.fillna(0)
    except Exception as e:
        print(f"[GESTAO] Erro estoque: {e}")
        return None

def _load_pagamentos_df():
    """Carrega DataFrame de pagamentos."""
    path = _find_latest_excel('Pagamentos')
    if not path:
        return None
    try:
        df = pd.read_excel(path, engine='openpyxl')
        df.columns = df.columns.str.strip()
        return df.fillna('')
    except Exception as e:
        print(f"[GESTAO] Erro pagamentos: {e}")
        return None

def _load_ordens_df():
    """Carrega DataFrame de ordens de compra (Geral contábil)."""
    path = _find_latest_excel('Geral contabil')
    if not path:
        return None
    try:
        df = pd.read_excel(path, engine='openpyxl')
        df.columns = df.columns.str.strip()
        return df.fillna('')
    except Exception as e:
        print(f"[GESTAO] Erro ordens: {e}")
        return None


# ═══════════════════════════════════════════════════════════════════
# ROTA: /gestao — Dashboard principal
# ═══════════════════════════════════════════════════════════════════
@app.route('/gestao')
@login_required
def gestao_dashboard():
    df_est  = _load_estoque_df()
    df_pag  = _load_pagamentos_df()
    df_ord  = _load_ordens_df()

    # Defaults seguros — nunca levantam KeyError
    kpis = {
        'valor_estoque': 'N/D', 'itens_estoque': '0', 'skus_estoque': '0',
        'total_pago': 'N/D', 'qtd_nfs': '0',
        'qtd_ordens': '0', 'valor_ordens': 'N/D',
    }

    # Estoque
    if df_est is not None and not df_est.empty:
        kpis['skus_estoque'] = _fmt_n(len(df_est))
        custo_col = next((c for c in df_est.columns if 'Custo' in c and 'Total' in c), None)
        if not custo_col:
            custo_col = next((c for c in df_est.columns if 'Custo' in c), None)
        if custo_col:
            df_est[custo_col] = pd.to_numeric(df_est[custo_col], errors='coerce').fillna(0)
            kpis['valor_estoque'] = _fmt_brl(df_est[custo_col].sum())
            kpis['itens_estoque'] = _fmt_n((df_est[custo_col] > 0).sum())

    # Pagamentos
    if df_pag is not None and not df_pag.empty:
        kpis['qtd_nfs'] = _fmt_n(len(df_pag))
        pag_col = next((c for c in df_pag.columns if 'Pago' in c or 'Valor' in c), None)
        if pag_col:
            kpis['total_pago'] = _fmt_brl(pd.to_numeric(df_pag[pag_col], errors='coerce').sum())

    # Ordens
    if df_ord is not None and not df_ord.empty:
        kpis['qtd_ordens'] = _fmt_n(len(df_ord))
        val_col = next((c for c in df_ord.columns if 'Valor' in c or 'Total' in c), None)
        if val_col:
            kpis['valor_ordens'] = _fmt_brl(pd.to_numeric(df_ord[val_col], errors='coerce').sum())

    content = f'''
<div class="g-header">
  <h1><i class="fas fa-chart-line"></i> Painel de Gestão</h1>
  <p>Dashboard executivo — Essencis · EMG Betim</p>
</div>

<div class="g-kpi-grid">
  <div class="g-kpi success">
    <div class="g-kpi-label">Valor em Estoque</div>
    <div class="g-kpi-value">{kpis['valor_estoque']}</div>
    <div class="g-kpi-sub">{kpis['skus_estoque']} SKUs · MAT1/MAT3/TNQ1</div>
  </div>
  <div class="g-kpi warning">
    <div class="g-kpi-label">Total Pago (período)</div>
    <div class="g-kpi-value">{kpis['total_pago']}</div>
    <div class="g-kpi-sub">{kpis['qtd_nfs']} notas fiscais</div>
  </div>
  <div class="g-kpi">
    <div class="g-kpi-label">Ordens de Compra</div>
    <div class="g-kpi-value">{kpis['qtd_ordens']}</div>
    <div class="g-kpi-sub">Valor: {kpis['valor_ordens']}</div>
  </div>
</div>

<div class="g-grid3">
  <div class="g-card" style="cursor:pointer;" onclick="location.href='/gestao/estoque'">
    <h3><i class="fas fa-warehouse" style="color:var(--accent-light);"></i> Estoque Análise</h3>
    <p style="font-size:0.85rem;color:var(--muted);">Curva ABC, grupos, subinventários e KPIs de estoque periódico.</p>
    <div style="margin-top:14px;font-size:0.78rem;font-weight:600;color:var(--accent-light);">Acessar →</div>
  </div>
  <div class="g-card" style="cursor:pointer;" onclick="location.href='/gestao/pagamentos'">
    <h3><i class="fas fa-file-invoice-dollar" style="color:var(--gold);"></i> Pagamentos</h3>
    <p style="font-size:0.85rem;color:var(--muted);">Ranking fornecedores, curva ABC de pagamentos e formas de pagamento.</p>
    <div style="margin-top:14px;font-size:0.78rem;font-weight:600;color:var(--gold);">Acessar →</div>
  </div>
  <div class="g-card" style="cursor:pointer;" onclick="location.href='/gestao/ordens'">
    <h3><i class="fas fa-shopping-cart" style="color:var(--accent);"></i> Ordens de Compra</h3>
    <p style="font-size:0.85rem;color:var(--muted);">Análise de ordens emitidas, status e valor por organização.</p>
    <div style="margin-top:14px;font-size:0.78rem;font-weight:600;color:var(--accent);">Acessar →</div>
  </div>
  <div class="g-card" style="cursor:pointer;" onclick="location.href='/gestao/precos'">
    <h3><i class="fas fa-chart-line" style="color:#4299e1;"></i> Análise de Preços</h3>
    <p style="font-size:0.85rem;color:var(--muted);">Variação de preços mês a mês, performance de negociação.</p>
    <div style="margin-top:14px;font-size:0.78rem;font-weight:600;color:#4299e1;">Acessar →</div>
  </div>
  <div class="g-card" style="cursor:pointer;" onclick="location.href='/gestao/pipefy'">
    <h3><i class="fas fa-tasks" style="color:#9f7aea;"></i> Atividades Pipefy</h3>
    <p style="font-size:0.85rem;color:var(--muted);">Compras CSC, Locais, Serviços, Envio NFE e Reservas.</p>
    <div style="margin-top:14px;font-size:0.78rem;font-weight:600;color:#9f7aea;">Acessar →</div>
  </div>
</div>
'''
    return _gestao_base('Dashboard Gestão', 'gestao', 'Ferramenta de Gestão', content)


# ═══════════════════════════════════════════════════════════════════
# ROTA: /gestao/estoque — Análise de Estoque
# ═══════════════════════════════════════════════════════════════════
@app.route('/gestao/estoque')
@login_required
def gestao_estoque():
    df = _load_estoque_df()
    if df is None or df.empty:
        content = '''<div class="g-header"><h1><i class="fas fa-warehouse"></i> Estoque — Análise</h1></div>
<div class="g-no-data"><div class="icon">📦</div>
<strong>Nenhum dado de estoque encontrado.</strong><br>
<span style="font-size:0.85rem;">Verifique a pasta: G:\\SUPRIMENTOS\\...\\Relatórios\\Estoque periodico\\ANO\\MES\\</span></div>'''
        return _gestao_base('Estoque Análise', 'gestao_estoque', 'Estoque — Análise', content)

    custo_col = next((c for c in df.columns if 'Custo' in c and 'Total' in c), None)
    qtd_col   = next((c for c in df.columns if 'Quantidade' in c), None)
    item_col  = next((c for c in df.columns if c.strip() == 'Nome do Item'), None)
    desc_col  = next((c for c in df.columns if 'Descri' in c), None)

    if custo_col:
        df[custo_col] = pd.to_numeric(df[custo_col], errors='coerce').fillna(0)
    if qtd_col:
        df[qtd_col] = pd.to_numeric(df[qtd_col], errors='coerce').fillna(0)

    total_valor = df[custo_col].sum() if custo_col else 0
    total_skus  = len(df)

    # Grupos
    grupos = df.groupby('Categoria')[custo_col].sum().sort_values(ascending=False) if custo_col else pd.Series()
    grupos_labels = json.dumps(grupos.index.tolist())
    grupos_values = json.dumps([round(v, 2) for v in grupos.values.tolist()])

    # Curva ABC
    abc_rows = ''
    if custo_col and item_col:
        df_abc = df.groupby(item_col)[custo_col].sum().sort_values(ascending=False).reset_index()
        df_abc['pct'] = df_abc[custo_col] / df_abc[custo_col].sum() * 100
        df_abc['cum'] = df_abc['pct'].cumsum()
        df_abc['cls'] = df_abc['cum'].apply(lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))
        a_val = _fmt_brl(df_abc[df_abc['cls']=='A'][custo_col].sum())
        b_val = _fmt_brl(df_abc[df_abc['cls']=='B'][custo_col].sum())
        c_val = _fmt_brl(df_abc[df_abc['cls']=='C'][custo_col].sum())
        a_qtd = len(df_abc[df_abc['cls']=='A'])
        b_qtd = len(df_abc[df_abc['cls']=='B'])
        c_qtd = len(df_abc[df_abc['cls']=='C'])
        for _, row in df_abc.head(30).iterrows():
            abc_rows += f'''<tr>
              <td style="font-family:monospace;font-size:0.82rem;color:var(--accent-light);">{str(row[item_col])[:20]}</td>
              <td>{str(row.get(desc_col,''))[:50] if desc_col else ''}</td>
              <td style="text-align:right;">{_fmt_brl(row[custo_col])}</td>
              <td style="text-align:right;">{row['pct']:.1f}%</td>
              <td style="text-align:right;">{row['cum']:.1f}%</td>
              <td><span class="g-badge {row['cls'].lower()}">{row['cls']}</span></td>
            </tr>'''
    else:
        a_val = b_val = c_val = 'N/D'
        a_qtd = b_qtd = c_qtd = 0

    # Top 10
    top10_rows = ''
    if custo_col and item_col:
        top10 = df.nlargest(10, custo_col)
        for i, (_, row) in enumerate(top10.iterrows(), 1):
            top10_rows += f'''<tr>
              <td><strong>{i}</strong></td>
              <td style="font-family:monospace;font-size:0.82rem;color:var(--accent-light);">{str(row.get(item_col,''))[:20]}</td>
              <td>{str(row.get(desc_col,''))[:50] if desc_col else ''}</td>
              <td style="text-align:right;">{_fmt_brl(row[custo_col])}</td>
              <td style="text-align:center;">{str(row.get('Categoria',''))}</td>
            </tr>'''

    content = f'''
<div class="g-header">
  <h1><i class="fas fa-warehouse"></i> Estoque — Análise Completa</h1>
  <p>Organização 00301_EMG_BETIM · Subinventários MAT1/MAT3/TNQ1</p>
</div>

<div class="g-kpi-grid">
  <div class="g-kpi success">
    <div class="g-kpi-label">Valor Total Estoque</div>
    <div class="g-kpi-value">{_fmt_brl(total_valor)}</div>
    <div class="g-kpi-sub">{_fmt_n(total_skus)} SKUs</div>
  </div>
  <div class="g-kpi danger"><div class="g-kpi-label">Classe A</div><div class="g-kpi-value">{a_val}</div><div class="g-kpi-sub">{a_qtd} itens · 80% do valor</div></div>
  <div class="g-kpi warning"><div class="g-kpi-label">Classe B</div><div class="g-kpi-value">{b_val}</div><div class="g-kpi-sub">{b_qtd} itens · 15% do valor</div></div>
  <div class="g-kpi"><div class="g-kpi-label">Classe C</div><div class="g-kpi-value">{c_val}</div><div class="g-kpi-sub">{c_qtd} itens · 5% do valor</div></div>
</div>

<div class="g-grid2">
  <div class="g-card">
    <h3><i class="fas fa-chart-bar"></i> Valor por Grupo</h3>
    <div class="g-chart-wrap"><canvas id="chartGrupos"></canvas></div>
  </div>
  <div class="g-card">
    <h3><i class="fas fa-chart-pie"></i> Distribuição %</h3>
    <div class="g-chart-wrap"><canvas id="chartPie"></canvas></div>
  </div>
</div>

<div class="g-card">
  <h3><i class="fas fa-trophy"></i> Top 10 Itens por Valor</h3>
  <div style="overflow-x:auto;">
  <table class="g-table">
    <thead><tr><th>#</th><th>Código</th><th>Descrição</th><th style="text-align:right;">Valor</th><th>Grupo</th></tr></thead>
    <tbody>{top10_rows or '<tr><td colspan="5" class="g-no-data">Sem dados</td></tr>'}</tbody>
  </table>
  </div>
</div>

<div class="g-card">
  <h3><i class="fas fa-chart-line"></i> Curva ABC — Top 30 Itens</h3>
  <div style="overflow-x:auto;max-height:420px;overflow-y:auto;">
  <table class="g-table">
    <thead><tr><th>Código</th><th>Descrição</th><th style="text-align:right;">Valor</th><th style="text-align:right;">% Item</th><th style="text-align:right;">% Acum.</th><th>Classe</th></tr></thead>
    <tbody>{abc_rows or '<tr><td colspan="6" class="g-no-data">Sem dados</td></tr>'}</tbody>
  </table>
  </div>
</div>

<script>
const gLabels = {grupos_labels};
const gValues = {grupos_values};
const colors = ['#1a3a2a','#2d6348','#b8903a','#8b2e2e','#4299e1','#9f7aea','#48bb78','#ed8936'];

new Chart(document.getElementById('chartGrupos'),{{
  type:'bar',
  data:{{labels:gLabels,datasets:[{{label:'Valor',data:gValues,backgroundColor:colors,borderRadius:4}}]}},
  options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},
    scales:{{x:{{ticks:{{callback:v=>'R$ '+v.toLocaleString('pt-BR')}}}}}}}}
}});
new Chart(document.getElementById('chartPie'),{{
  type:'doughnut',
  data:{{labels:gLabels,datasets:[{{data:gValues,backgroundColor:colors,borderWidth:0}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'right',labels:{{font:{{size:10}}}}}}}}}}
}});
</script>
'''
    return _gestao_base('Estoque Análise', 'gestao_estoque', 'Estoque — Análise', content)


# ═══════════════════════════════════════════════════════════════════
# ROTA: /gestao/pagamentos — Relatório de Pagamentos
# ═══════════════════════════════════════════════════════════════════
@app.route('/gestao/pagamentos')
@login_required
def gestao_pagamentos():
    df = _load_pagamentos_df()
    if df is None or df.empty:
        content = '''<div class="g-header"><h1><i class="fas fa-file-invoice-dollar"></i> Pagamentos</h1></div>
<div class="g-no-data"><div class="icon">💰</div>
<strong>Nenhum dado de pagamento encontrado.</strong><br>
<span style="font-size:0.85rem;">Verifique a pasta: G:\\...\\Relatórios\\Pagamentos\\ANO\\MES\\</span></div>'''
        return _gestao_base('Pagamentos', 'gestao_pagamentos', 'Pagamentos', content)

    # Identificar colunas
    val_col  = next((c for c in df.columns if 'Pago' in c or ('Valor' in c and 'Pago' in c)), None)
    if not val_col:
        val_col = next((c for c in df.columns if 'Valor' in c), None)
    forn_col = next((c for c in df.columns if 'Fornec' in c or 'Razão' in c or 'Nome' in c), None)
    if val_col:
        df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)

    total_pago = df[val_col].sum() if val_col else 0
    qtd_nfs    = len(df)
    ticket_med = total_pago / qtd_nfs if qtd_nfs else 0

    # Top 10 fornecedores
    ranking = []
    rank_rows = ''
    if val_col and forn_col:
        grp = df.groupby(forn_col)[val_col].sum().sort_values(ascending=False).head(10).reset_index()
        grp.columns = ['Fornecedor', 'Total']
        grp['pct'] = grp['Total'] / total_pago * 100 if total_pago else 0
        grp['cum'] = grp['pct'].cumsum()
        grp['cls'] = grp['cum'].apply(lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))
        ranking = grp.to_dict('records')
        for i, r in enumerate(ranking, 1):
            rank_rows += f'''<tr>
              <td><strong>{i}</strong></td>
              <td>{str(r['Fornecedor'])[:40]}</td>
              <td style="text-align:right;">{_fmt_brl(r['Total'])}</td>
              <td style="text-align:right;">{r['pct']:.1f}%</td>
              <td><span class="g-badge {r['cls'].lower()}">{r['cls']}</span></td>
            </tr>'''

    rank_labels = json.dumps([str(r['Fornecedor'])[:20] for r in ranking])
    rank_values = json.dumps([round(r['Total'], 2) for r in ranking])

    content = f'''
<div class="g-header">
  <h1><i class="fas fa-file-invoice-dollar"></i> Relatório de Pagamentos</h1>
  <p>Análise de fornecedores e valores pagos no período</p>
</div>

<div class="g-kpi-grid">
  <div class="g-kpi success">
    <div class="g-kpi-label">Total Pago</div>
    <div class="g-kpi-value">{_fmt_brl(total_pago)}</div>
    <div class="g-kpi-sub">{_fmt_n(qtd_nfs)} notas fiscais</div>
  </div>
  <div class="g-kpi warning">
    <div class="g-kpi-label">Ticket Médio</div>
    <div class="g-kpi-value">{_fmt_brl(ticket_med)}</div>
    <div class="g-kpi-sub">por nota fiscal</div>
  </div>
  <div class="g-kpi">
    <div class="g-kpi-label">Fornecedores Únicos</div>
    <div class="g-kpi-value">{_fmt_n(df[forn_col].nunique() if forn_col else 0)}</div>
    <div class="g-kpi-sub">no período</div>
  </div>
</div>

<div class="g-grid2">
  <div class="g-card">
    <h3><i class="fas fa-chart-bar"></i> Top 10 Fornecedores</h3>
    <div class="g-chart-wrap"><canvas id="chartForn"></canvas></div>
  </div>
  <div class="g-card">
    <h3><i class="fas fa-chart-pie"></i> Participação %</h3>
    <div class="g-chart-wrap"><canvas id="chartPieForn"></canvas></div>
  </div>
</div>

<div class="g-card">
  <h3><i class="fas fa-list"></i> Ranking de Fornecedores — Curva ABC</h3>
  <div style="overflow-x:auto;">
  <table class="g-table">
    <thead><tr><th>#</th><th>Fornecedor</th><th style="text-align:right;">Total Pago</th><th style="text-align:right;">Participação</th><th>Classe</th></tr></thead>
    <tbody>{rank_rows or '<tr><td colspan="5" class="g-no-data">Sem dados</td></tr>'}</tbody>
  </table>
  </div>
</div>

<script>
const fLabels = {rank_labels};
const fValues = {rank_values};
const fColors = ['#1a3a2a','#2d6348','#b8903a','#8b2e2e','#4299e1','#9f7aea','#48bb78','#ed8936','#38b2ac','#667eea'];

new Chart(document.getElementById('chartForn'),{{
  type:'bar',
  data:{{labels:fLabels,datasets:[{{label:'Total',data:fValues,backgroundColor:'rgba(45,99,72,0.7)',borderRadius:4}}]}},
  options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},
    scales:{{x:{{ticks:{{callback:v=>'R$ '+v.toLocaleString('pt-BR')}}}}}}}}
}});
new Chart(document.getElementById('chartPieForn'),{{
  type:'doughnut',
  data:{{labels:fLabels,datasets:[{{data:fValues,backgroundColor:fColors,borderWidth:0}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'right',labels:{{font:{{size:10}}}}}}}}}}
}});
</script>
'''
    return _gestao_base('Pagamentos', 'gestao_pagamentos', 'Pagamentos', content)


# ═══════════════════════════════════════════════════════════════════
# ROTA: /gestao/ordens — Ordens de Compra
# ═══════════════════════════════════════════════════════════════════
@app.route('/gestao/ordens')
@login_required
def gestao_ordens():
    df = _load_ordens_df()
    if df is None or df.empty:
        content = '''<div class="g-header"><h1><i class="fas fa-shopping-cart"></i> Ordens de Compra</h1></div>
<div class="g-no-data"><div class="icon">🛒</div>
<strong>Nenhum dado de ordens de compra encontrado.</strong><br>
<span style="font-size:0.85rem;">Verifique a pasta: G:\\...\\Relatórios\\Geral contabil\\ANO\\MES\\</span></div>'''
        return _gestao_base('Ordens de Compra', 'gestao_ordens', 'Ordens de Compra', content)

    val_col   = next((c for c in df.columns if 'Valor' in c or 'Total' in c), None)
    forn_col  = next((c for c in df.columns if 'Fornec' in c or 'Razão' in c or 'Nome do Fornec' in c), None)
    org_col   = next((c for c in df.columns if 'Organiz' in c), None)
    status_col= next((c for c in df.columns if 'Status' in c or 'Situação' in c), None)
    tipo_col  = next((c for c in df.columns if 'Tipo' in c), None)

    if val_col:
        df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)

    total_val = df[val_col].sum() if val_col else 0
    qtd_ocs   = len(df)

    # Por organização
    org_chart_labels = '[]'
    org_chart_values = '[]'
    if org_col and val_col:
        grp_org = df.groupby(org_col)[val_col].sum().sort_values(ascending=False).head(8)
        org_chart_labels = json.dumps(grp_org.index.tolist())
        org_chart_values = json.dumps([round(v,2) for v in grp_org.values.tolist()])

    # Por status
    status_rows = ''
    if status_col:
        grp_st = df.groupby(status_col).size().sort_values(ascending=False)
        for st, cnt in grp_st.items():
            status_rows += f'<tr><td>{st}</td><td style="text-align:right;">{cnt}</td></tr>'

    # Preview tabela
    cols_show = [c for c in [forn_col, org_col, status_col, tipo_col, val_col] if c][:5]
    table_head = ''.join(f'<th>{c}</th>' for c in cols_show)
    table_body = ''
    for _, row in df.head(50).iterrows():
        cells = ''.join(f'<td>{str(row.get(c,""))[:40]}</td>' for c in cols_show)
        table_body += f'<tr>{cells}</tr>'

    content = f'''
<div class="g-header">
  <h1><i class="fas fa-shopping-cart"></i> Ordens de Compra</h1>
  <p>Análise de ordens emitidas no período</p>
</div>

<div class="g-kpi-grid">
  <div class="g-kpi success">
    <div class="g-kpi-label">Valor Total OCs</div>
    <div class="g-kpi-value">{_fmt_brl(total_val)}</div>
    <div class="g-kpi-sub">{_fmt_n(qtd_ocs)} ordens emitidas</div>
  </div>
  <div class="g-kpi warning">
    <div class="g-kpi-label">Organizações</div>
    <div class="g-kpi-value">{_fmt_n(df[org_col].nunique() if org_col else 0)}</div>
    <div class="g-kpi-sub">unidades com ordens</div>
  </div>
  <div class="g-kpi">
    <div class="g-kpi-label">Fornecedores</div>
    <div class="g-kpi-value">{_fmt_n(df[forn_col].nunique() if forn_col else 0)}</div>
    <div class="g-kpi-sub">únicos no período</div>
  </div>
</div>

<div class="g-grid2">
  <div class="g-card">
    <h3><i class="fas fa-building"></i> Valor por Organização</h3>
    <div class="g-chart-wrap"><canvas id="chartOrg"></canvas></div>
  </div>
  <div class="g-card">
    <h3><i class="fas fa-tag"></i> Ordens por Status</h3>
    <div style="overflow-x:auto;max-height:260px;overflow-y:auto;">
    <table class="g-table">
      <thead><tr><th>Status</th><th style="text-align:right;">Qtd</th></tr></thead>
      <tbody>{status_rows or '<tr><td colspan="2" class="g-no-data">N/D</td></tr>'}</tbody>
    </table>
    </div>
  </div>
</div>

<div class="g-card">
  <h3><i class="fas fa-table"></i> Ordens — Primeiros 50 registros</h3>
  <div style="overflow-x:auto;max-height:400px;overflow-y:auto;">
  <table class="g-table">
    <thead><tr>{table_head}</tr></thead>
    <tbody>{table_body or '<tr><td colspan="5" class="g-no-data">Sem dados</td></tr>'}</tbody>
  </table>
  </div>
</div>

<script>
new Chart(document.getElementById('chartOrg'),{{
  type:'bar',
  data:{{labels:{org_chart_labels},datasets:[{{label:'Valor',data:{org_chart_values},backgroundColor:'rgba(45,99,72,0.75)',borderRadius:4}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},
    scales:{{y:{{ticks:{{callback:v=>'R$ '+v.toLocaleString('pt-BR')}}}}}}}}
}});
</script>
'''
    return _gestao_base('Ordens de Compra', 'gestao_ordens', 'Ordens de Compra', content)


# ═══════════════════════════════════════════════════════════════════
# ROTA: /gestao/precos — Análise de Preços (variação mensal)
# ═══════════════════════════════════════════════════════════════════
@app.route('/gestao/precos')
@login_required
def gestao_precos():
    base = r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios\Geral contabil"
    # Coleta todos os arquivos dos últimos 12 meses
    all_files = glob.glob(os.path.join(base, '*', '*', '*.xls*'))
    if not all_files:
        all_files = glob.glob(os.path.join(BASE_GESTAO, '**', 'Geral*contabil*.xls*'), recursive=True)

    if not all_files:
        content = '''<div class="g-header"><h1><i class="fas fa-chart-line"></i> Análise de Preços</h1></div>
<div class="g-no-data"><div class="icon">📊</div>
<strong>Nenhum arquivo de ordens de compra encontrado.</strong><br>
<span style="font-size:0.85rem;">Verifique a pasta: G:\\...\\Relatórios\\Geral contabil\\ANO\\MES\\</span></div>'''
        return _gestao_base('Análise de Preços', 'gestao_precos', 'Análise de Preços', content)

    dfs = []
    for f in sorted(all_files, key=os.path.getmtime)[-12:]:
        try:
            _df = pd.read_excel(f, engine='openpyxl')
            _df.columns = _df.columns.str.strip()
            # Tentar extrair ano/mes do path
            parts = f.replace('\\', '/').split('/')
            try:
                mes_str = parts[-2]; ano_str = parts[-3]
                _df['_mes_ref'] = f"{ano_str}-{mes_str.zfill(2)}"
            except Exception:
                _df['_mes_ref'] = 'N/D'
            dfs.append(_df)
        except Exception:
            pass

    if not dfs:
        content = '<div class="g-no-data"><div class="icon">📊</div><strong>Erro ao ler arquivos de ordens.</strong></div>'
        return _gestao_base('Análise de Preços', 'gestao_precos', 'Análise de Preços', content)

    df_all = pd.concat(dfs, ignore_index=True)
    val_col  = next((c for c in df_all.columns if 'Preço' in c or 'Price' in c or ('Valor' in c and 'Unit' in c.title())), None)
    if not val_col:
        val_col = next((c for c in df_all.columns if 'Valor' in c), None)
    item_col = next((c for c in df_all.columns if 'Item' in c or 'Material' in c or 'Descri' in c), None)

    if val_col:
        df_all[val_col] = pd.to_numeric(df_all[val_col], errors='coerce').fillna(0)

    # Evolução mensal de valores
    evol_labels = '[]'; evol_values = '[]'
    if val_col and '_mes_ref' in df_all.columns:
        grp = df_all.groupby('_mes_ref')[val_col].sum().sort_index()
        evol_labels = json.dumps(grp.index.tolist())
        evol_values = json.dumps([round(v,2) for v in grp.values.tolist()])

    # Top itens com maior variação
    variacao_rows = ''
    if val_col and item_col and '_mes_ref' in df_all.columns:
        pivot = df_all.pivot_table(index=item_col, columns='_mes_ref', values=val_col, aggfunc='mean').fillna(0)
        if pivot.shape[1] >= 2:
            pivot['variacao'] = pivot.iloc[:, -1] - pivot.iloc[:, -2]
            pivot['var_pct']  = pivot['variacao'] / pivot.iloc[:, -2].replace(0, float('nan')) * 100
            top_var = pivot.nlargest(20, 'variacao').reset_index()
            for _, row in top_var.iterrows():
                cor_var = '#8b2e2e' if row['variacao'] > 0 else '#2d6348'
                sinal   = '+' if row['variacao'] > 0 else ''
                variacao_rows += f'''<tr>
                  <td style="font-size:0.82rem;">{str(row[item_col])[:50]}</td>
                  <td style="text-align:right;">{_fmt_brl(row.iloc[-3] if len(row) > 3 else 0)}</td>
                  <td style="text-align:right;">{_fmt_brl(row.iloc[-2] if len(row) > 2 else 0)}</td>
                  <td style="text-align:right;color:{cor_var};font-weight:600;">{sinal}{_fmt_brl(row['variacao'])}</td>
                  <td style="text-align:right;color:{cor_var};">{sinal}{row['var_pct']:.1f}%</td>
                </tr>'''

    meses_label = df_all['_mes_ref'].unique().tolist() if '_mes_ref' in df_all.columns else []
    col_headers = ''.join(f'<th>{m}</th>' for m in sorted(meses_label)[-2:]) if len(meses_label) >= 2 else '<th>Anterior</th><th>Atual</th>'

    content = f'''
<div class="g-header">
  <h1><i class="fas fa-chart-line"></i> Análise de Variação de Preços</h1>
  <p>Comparativo de preços mês a mês — últimos 12 meses</p>
</div>

<div class="g-kpi-grid">
  <div class="g-kpi">
    <div class="g-kpi-label">Meses analisados</div>
    <div class="g-kpi-value">{len(meses_label)}</div>
  </div>
  <div class="g-kpi warning">
    <div class="g-kpi-label">Registros totais</div>
    <div class="g-kpi-value">{_fmt_n(len(df_all))}</div>
  </div>
  <div class="g-kpi success">
    <div class="g-kpi-label">Valor total</div>
    <div class="g-kpi-value">{_fmt_brl(df_all[val_col].sum() if val_col else 0)}</div>
  </div>
</div>

<div class="g-card">
  <h3><i class="fas fa-chart-area"></i> Evolução do Valor Total de Compras por Mês</h3>
  <div class="g-chart-wrap" style="height:320px;">
    <canvas id="chartEvol"></canvas>
  </div>
</div>

<div class="g-card">
  <h3><i class="fas fa-arrow-up"></i> Top 20 Itens com Maior Variação de Preço</h3>
  <div style="overflow-x:auto;max-height:440px;overflow-y:auto;">
  <table class="g-table">
    <thead><tr><th>Item</th>{col_headers}<th style="text-align:right;">Variação R$</th><th style="text-align:right;">Variação %</th></tr></thead>
    <tbody>{variacao_rows or '<tr><td colspan="5" class="g-no-data">Dados insuficientes para calcular variação</td></tr>'}</tbody>
  </table>
  </div>
</div>

<script>
new Chart(document.getElementById('chartEvol'),{{
  type:'line',
  data:{{
    labels:{evol_labels},
    datasets:[{{
      label:'Valor Total',
      data:{evol_values},
      borderColor:'#1a3a2a',backgroundColor:'rgba(26,58,42,0.1)',
      borderWidth:2.5,fill:true,tension:0.4,pointRadius:5
    }}]
  }},
  options:{{
    responsive:true,maintainAspectRatio:false,
    plugins:{{legend:{{display:false}}}},
    scales:{{y:{{ticks:{{callback:v=>'R$ '+v.toLocaleString('pt-BR')}}}}}}
  }}
}});
</script>
'''
    return _gestao_base('Análise de Preços', 'gestao_precos', 'Análise de Preços', content)


# ═══════════════════════════════════════════════════════════════════
# ROTA: /gestao/pipefy — Dashboard Pipefy (Atividades)
# ═══════════════════════════════════════════════════════════════════
@app.route('/gestao/pipefy')
@app.route('/gestao/pipefy/<categoria>')
@login_required
def gestao_pipefy(categoria='compras_csc'):
    if categoria not in PIPEFY_PIPES:
        categoria = 'compras_csc'

    TITULOS = {
        'compras_csc':       'Compras CSC',
        'compras_locais':    'Compras Locais',
        'compras_servicos':  'Serviços',
        'envio_nfe':         'Envio de NFE',
        'reserva_materiais': 'Reserva de Materiais',
    }

    # Sub-nav tabs
    tabs_html = ''
    for key, titulo in TITULOS.items():
        cls = 'g-tab active' if key == categoria else 'g-tab'
        tabs_html += f'<button class="{cls}" onclick="location.href=\'/gestao/pipefy/{key}\'">{titulo}</button>'

    pipe_id = PIPEFY_PIPES[categoria]
    try:
        cards = _pipefy_get_all_cards(pipe_id)
    except Exception as e:
        content = f'''
<div class="g-header"><h1><i class="fas fa-tasks"></i> Atividades Pipefy — {TITULOS[categoria]}</h1></div>
<div class="g-tab-bar">{tabs_html}</div>
<div class="g-no-data"><div class="icon">⚠️</div>
<strong>Erro ao conectar ao Pipefy:</strong> {e}</div>'''
        return _gestao_base('Atividades Pipefy', 'gestao_pipefy', f'Pipefy — {TITULOS[categoria]}', content)

    if not cards:
        content = f'''
<div class="g-header"><h1><i class="fas fa-tasks"></i> Atividades Pipefy — {TITULOS[categoria]}</h1></div>
<div class="g-tab-bar">{tabs_html}</div>
<div class="g-no-data"><div class="icon">📋</div><strong>Nenhum card encontrado.</strong><br>
<span style="font-size:0.85rem;">Verifique as credenciais Pipefy em config.json ou o pipe ID.</span></div>'''
        return _gestao_base('Atividades Pipefy', 'gestao_pipefy', f'Pipefy — {TITULOS[categoria]}', content)

    from datetime import timezone
    hoje = datetime.datetime.now(timezone.utc)
    total  = len(cards)
    concl  = sum(1 for c in cards if c.get('finished_at'))
    pend   = total - concl
    venc   = 0
    for c in cards:
        if not c.get('finished_at') and c.get('due_date'):
            try:
                due_dt = datetime.datetime.fromisoformat(c['due_date'].replace('Z','+00:00'))
                if due_dt < hoje:
                    venc += 1
            except Exception:
                pass

    # Distribuição por fase
    fases = {}
    for c in cards:
        ph = (c.get('current_phase') or {}).get('name', 'Sem Fase')
        fases[ph] = fases.get(ph,0) + 1
    fases_sorted = sorted(fases.items(), key=lambda x: x[1], reverse=True)
    fases_labels = json.dumps([f[0] for f in fases_sorted[:8]])
    fases_values = json.dumps([f[1] for f in fases_sorted[:8]])

    # Responsáveis
    resp = {}
    for c in cards:
        for a in (c.get('assignees') or []):
            n = a.get('name','?')
            resp[n] = resp.get(n,0) + 1
    resp_sorted = sorted(resp.items(), key=lambda x: x[1], reverse=True)

    resp_rows = ''
    for nome, qtd in resp_sorted[:10]:
        resp_rows += f'<tr><td>{nome}</td><td style="text-align:right;">{qtd}</td></tr>'

    # Gargalos (mais antigos abertos)
    gargalos = []
    for c in cards:
        if not c.get('finished_at') and c.get('created_at'):
            try:
                dt = datetime.datetime.fromisoformat(c['created_at'].replace('Z','+00:00'))
                dias = (hoje - dt).days
                gargalos.append({
                    'titulo': (c.get('title','')[:35] or 'Sem título'),
                    'fase':   (c.get('current_phase') or {}).get('name','N/A'),
                    'dias':   dias,
                    'url':    c.get('url','#'),
                })
            except Exception:
                pass
    gargalos.sort(key=lambda x: x['dias'], reverse=True)
    gargalo_rows = ''
    for g in gargalos[:8]:
        cor = '#8b2e2e' if g['dias'] > 30 else ('#b8903a' if g['dias'] > 14 else '#2d6348')
        gargalo_rows += f'''<tr>
          <td><a href="{g['url']}" target="_blank" style="color:var(--accent-light);text-decoration:none;">{g['titulo']}</a></td>
          <td>{g['fase'][:20]}</td>
          <td style="text-align:center;"><span style="background:{cor};color:#fff;padding:2px 8px;border-radius:10px;font-size:0.72rem;">{g['dias']}d</span></td>
        </tr>'''

    content = f'''
<div class="g-header">
  <h1><i class="fas fa-tasks"></i> Atividades Pipefy — {TITULOS[categoria]}</h1>
  <p>Dados em tempo real via API Pipefy</p>
</div>

<div class="g-tab-bar">{tabs_html}</div>

<div class="g-kpi-grid">
  <div class="g-kpi"><div class="g-kpi-label">Total</div><div class="g-kpi-value">{total}</div><div class="g-kpi-sub">cards no pipe</div></div>
  <div class="g-kpi success"><div class="g-kpi-label">Concluídos</div><div class="g-kpi-value">{concl}</div><div class="g-kpi-sub">{(concl/total*100):.0f}% de conclusão</div></div>
  <div class="g-kpi warning"><div class="g-kpi-label">Em Andamento</div><div class="g-kpi-value">{pend}</div></div>
  <div class="g-kpi danger"><div class="g-kpi-label">Vencidos</div><div class="g-kpi-value">{venc}</div><div class="g-kpi-sub">por due_date</div></div>
</div>

<div class="g-grid2">
  <div class="g-card">
    <h3><i class="fas fa-chart-pie"></i> Distribuição por Fase</h3>
    <div class="g-chart-wrap"><canvas id="chartFases"></canvas></div>
  </div>
  <div class="g-card">
    <h3><i class="fas fa-user-check"></i> Responsáveis — Top 10</h3>
    <div style="overflow-x:auto;max-height:260px;overflow-y:auto;">
    <table class="g-table">
      <thead><tr><th>Responsável</th><th style="text-align:right;">Cards</th></tr></thead>
      <tbody>{resp_rows or '<tr><td colspan="2" class="g-no-data">Sem atribuições</td></tr>'}</tbody>
    </table>
    </div>
  </div>
</div>

<div class="g-card">
  <h3><i class="fas fa-exclamation-triangle" style="color:var(--gold);"></i> Principais Gargalos (cards mais antigos)</h3>
  <div style="overflow-x:auto;">
  <table class="g-table">
    <thead><tr><th>Título</th><th>Fase</th><th style="text-align:center;">Tempo</th></tr></thead>
    <tbody>{gargalo_rows or '<tr><td colspan="3" class="g-no-data">Nenhum gargalo identificado</td></tr>'}</tbody>
  </table>
  </div>
</div>

<script>
new Chart(document.getElementById('chartFases'),{{
  type:'doughnut',
  data:{{
    labels:{fases_labels},
    datasets:[{{data:{fases_values},
      backgroundColor:['#1a3a2a','#2d6348','#b8903a','#8b2e2e','#4299e1','#9f7aea','#48bb78','#ed8936'],
      borderWidth:0
    }}]
  }},
  options:{{responsive:true,maintainAspectRatio:false,cutout:'60%',
    plugins:{{legend:{{position:'right',labels:{{font:{{size:10}}}}}}}}}}
}});
</script>
'''
    return _gestao_base('Atividades Pipefy', 'gestao_pipefy', f'Pipefy — {TITULOS[categoria]}', content)


# ─── API endpoint para atualizar token Pipefy via AJAX ───────────────────────
@app.route('/gestao/api/pipefy-status')
@login_required
def pipefy_status():
    token = _get_pipefy_token()
    return jsonify({'ok': token is not None, 'token': bool(token)})


if __name__ == '__main__' and not os.environ.get('PYTHONANYWHERE_DOMAIN'):
    app.run(debug=True, host='0.0.0.0', port=5000)
