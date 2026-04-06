import sys
import requests
import csv
import json
from datetime import datetime
from typing import Dict, List, Optional
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                             QLabel, QComboBox, QTabWidget, QTextEdit, QMessageBox,
                             QProgressBar, QHeaderView, QDialog,
                             QDialogButtonBox, QLineEdit, QFormLayout, QRadioButton)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor

# ============================================================================
# CLASSE CLIENTE API
# ============================================================================
class PipefyClient:
    """Cliente para integração com Pipefy via API GraphQL"""
    
    def __init__(self, client_id: str = None, client_secret: str = None, token_url: str = None, personal_token: str = None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.token_url = token_url
        self.api_url = "https://api.pipefy.com/graphql"
        self.access_token = personal_token
        self.use_personal_token = bool(personal_token)
        
    def authenticate(self) -> bool:
        if self.use_personal_token:
            print("Usando Personal Access Token")
            return True
        try:
            headers = {'Content-Type': 'application/x-www-form-urlencoded'}
            payload = {
                'grant_type': 'client_credentials',
                'client_id': self.client_id,
                'client_secret': self.client_secret
            }
            response = requests.post(self.token_url, data=payload, headers=headers, timeout=30)
            if response.status_code != 200:
                raise Exception(f"Autenticação falhou: {response.status_code}")
            self.access_token = response.json().get('access_token')
            return True
        except Exception as e:
            raise Exception(f"Erro na autenticação: {e}")
    
    def _make_request(self, query: str, variables: Optional[Dict] = None) -> Dict:
        if not self.access_token:
            raise Exception("Cliente não autenticado.")
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        response = requests.post(self.api_url, json={'query': query, 'variables': variables}, headers=headers, timeout=120)
        if response.status_code != 200:
            raise Exception(f"Erro HTTP {response.status_code}: {response.text}")
        return response.json()

    def get_multiple_pipes(self, pipe_ids: List[str]) -> List[Dict]:
        query = """
        query ($pipeId: ID!) {
          pipe(id: $pipeId) { id name }
        }
        """
        pipes = []
        for pipe_id in pipe_ids:
            try:
                res = self._make_request(query, {'pipeId': str(pipe_id)})
                if 'errors' not in res:
                    pipes.append(res['data']['pipe'])
            except:
                pass
        return pipes

    # ========================================================================
    # MÉTODOS PARA DATABASE (TABELA)
    # ========================================================================
    def get_database(self, table_id: str) -> Dict:
        """
        Busca informações da tabela (database) do Pipefy
        """
        query = """
        query ($tableId: ID!) {
          table(id: $tableId) {
            id
            name
            description
            public
            authorization
            create_record_button_label
            title_field {
              id
              label
            }
            table_fields {
              id
              label
              type
              required
              description
            }
          }
        }
        """
        result = self._make_request(query, {'tableId': table_id})
        if 'errors' in result:
            raise Exception(f"Erro ao buscar database: {result['errors']}")
        return result.get('data', {}).get('table', {})

    def get_all_database_records(self, table_id: str, progress_callback=None) -> Dict:
        """
        Busca TODOS os registros de uma tabela (database) do Pipefy
        """
        query = """
        query ($tableId: ID!, $first: Int!, $after: String) {
          table_records(table_id: $tableId, first: $first, after: $after) {
            pageInfo {
              hasNextPage
              endCursor
            }
            edges {
              node {
                id
                title
                created_at
                updated_at
                url
                status {
                  id
                  name
                }
                record_fields {
                  name
                  value
                  filled_at
                  updated_at
                  field {
                    id
                    label
                    type
                  }
                }
              }
            }
          }
        }
        """
        
        all_records = []
        has_next = True
        cursor = None
        page_count = 0
        
        print(f"Iniciando download da Database {table_id}...")
        
        while has_next:
            variables = {'tableId': table_id, 'first': 50, 'after': cursor}
            result = self._make_request(query, variables)
            
            if 'errors' in result:
                msg = result['errors'][0].get('message', str(result['errors']))
                raise Exception(f"Erro API Database: {msg}")
            
            records_conn = result.get('data', {}).get('table_records', {})
            
            new_records = [edge['node'] for edge in records_conn.get('edges', [])]
            all_records.extend(new_records)
            
            page_info = records_conn.get('pageInfo', {})
            has_next = page_info.get('hasNextPage', False)
            cursor = page_info.get('endCursor')
            
            page_count += 1
            if progress_callback:
                progress_callback(f"Baixando página {page_count}... ({len(all_records)} registros)")
            print(f"  -> Página {page_count}: +{len(new_records)} registros (Total: {len(all_records)})")
        
        # Buscar info da tabela
        table_info = self.get_database(table_id)
        
        return {
            'name': table_info.get('name', f'Database {table_id}'),
            'records': all_records,
            'fields': table_info.get('table_fields', [])
        }

    def create_database_record(self, table_id: str, fields_values: List[Dict]) -> Dict:
        """
        Cria um novo registro na database
        fields_values: lista de dicts com 'field_id' e 'field_value'
        """
        mutation = """
        mutation ($tableId: ID!, $fieldsValues: [TableRecordFieldInput]!) {
          createTableRecord(input: {table_id: $tableId, fields_attributes: $fieldsValues}) {
            table_record {
              id
              title
              url
            }
          }
        }
        """
        result = self._make_request(mutation, {
            'tableId': table_id,
            'fieldsValues': fields_values
        })
        if 'errors' in result:
            raise Exception(f"Erro ao criar registro: {result['errors']}")
        return result.get('data', {}).get('createTableRecord', {}).get('table_record', {})
    
    def get_all_pipe_cards(self, pipe_id: str, progress_callback=None) -> Dict:
        """
        Busca TODOS os cards com TODOS os campos possíveis do Pipefy
        """
        query = """
        query ($pipeId: ID!, $limit: Int!, $after: String) {
          cards(pipe_id: $pipeId, first: $limit, after: $after) {
            pageInfo {
              hasNextPage
              endCursor
            }
            edges {
              node {
                id
                title
                created_at
                updated_at
                due_date
                finished_at
                url
                current_phase { 
                  id
                  name 
                }
                phases_history {
                  phase {
                    id
                    name
                  }
                  firstTimeIn
                  lastTimeOut
                }
                labels { 
                  id
                  name 
                  color
                }
                assignees { 
                  id
                  name 
                  email
                }
                comments_count
                attachments_count
                fields { 
                  name 
                  value
                  filled_at
                  updated_at
                }
                parent_relations {
                  id
                  name
                }
                child_relations {
                  id
                  name
                }
              }
            }
          }
          pipe(id: $pipeId) { 
            name 
            id
          }
        }
        """
        
        all_cards = []
        has_next = True
        cursor = None
        pipe_name = ""
        page_count = 0
        
        print(f"Iniciando download total do Pipe {pipe_id}...")
        
        while has_next:
            variables = {'pipeId': pipe_id, 'limit': 50, 'after': cursor}
            result = self._make_request(query, variables)
            
            if 'errors' in result:
                msg = result['errors'][0].get('message')
                raise Exception(f"Erro API: {msg}")

            data = result.get('data', {})
            pipe_name = data.get('pipe', {}).get('name', f'Pipe {pipe_id}')
            cards_conn = data.get('cards', {})
            
            new_cards = [edge['node'] for edge in cards_conn.get('edges', [])]
            all_cards.extend(new_cards)
            
            page_info = cards_conn.get('pageInfo', {})
            has_next = page_info.get('hasNextPage', False)
            cursor = page_info.get('endCursor')
            
            page_count += 1
            if progress_callback:
                progress_callback(f"Baixando página {page_count}... ({len(all_cards)} cards)")
            print(f"  -> Página {page_count}: +{len(new_cards)} cards (Total: {len(all_cards)})")

        return {
            'name': pipe_name,
            'cards': all_cards
        }

# ============================================================================
# THREAD DE CARREGAMENTO
# ============================================================================
class LoadDataThread(QThread):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    
    def __init__(self, client, pipe_id=None, pipe_ids=None, database_id=None, database_ids=None):
        super().__init__()
        self.client = client
        self.pipe_id = pipe_id
        self.pipe_ids = pipe_ids
        self.database_id = database_id
        self.database_ids = database_ids
        
    def run(self):
        try:
            if self.pipe_id:
                data = self.client.get_all_pipe_cards(
                    self.pipe_id, 
                    progress_callback=lambda msg: self.progress.emit(msg)
                )
                self.finished.emit(data)
            elif self.pipe_ids:
                self.progress.emit("Carregando lista de pipes...")
                pipes = self.client.get_multiple_pipes(self.pipe_ids)
                self.finished.emit(pipes)
            elif self.database_id:
                data = self.client.get_all_database_records(
                    self.database_id,
                    progress_callback=lambda msg: self.progress.emit(msg)
                )
                self.finished.emit({'type': 'database', 'data': data})
            elif self.database_ids:
                self.progress.emit("Carregando lista de databases...")
                databases = []
                for db_id in self.database_ids:
                    try:
                        db_info = self.client.get_database(db_id)
                        databases.append(db_info)
                    except:
                        pass
                self.finished.emit({'type': 'database_list', 'data': databases})
        except Exception as e:
            self.error.emit(str(e))

# ============================================================================
# INTERFACE
# ============================================================================
class PipefyReportApp(QMainWindow):
    
    def __init__(self):
        super().__init__()
        
        # CREDENCIAIS
        self.CLIENT_ID = "ofgUSnXFhXadEzrDd_ZtUzXsV8-Crv-0NFboRn0CbrU"
        self.CLIENT_SECRET = "DyLPVER8t6SIeVpO7lQiDqTzoquM3UqLDOUgOFtHFpw"
        self.TOKEN_URL = "https://app.pipefy.com/oauth/token"
        
        self.PIPE_IDS = [
            "306527874", "306859940", "306858226", "306864423", "306859726"
        ]
        
        # DATABASE IDs (Tabelas)
        self.DATABASE_IDS = [
            "306859259"  # Database principal (Ordens de Compra)
        ]
        
        # IDs específicos para análise de tempo de entrega
        self.DATABASE_ORDENS_COMPRA = "306859259"  # Database com ordens de compra
        self.PIPE_LANCAMENTO_NFE = None  # Será identificado pelo nome
        
        self.client = None
        self.current_cards = []
        self.current_records = []  # Para registros de database
        self.dynamic_headers = [] 
        
        self.init_ui()
        self.authenticate()
        
    def init_ui(self):
        self.setWindowTitle("Pipefy - Relatório Completo (Todas as Colunas)")
        self.resize(1400, 900)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        # Header
        lbl = QLabel("📊 Relatório Completo Pipefy - Todas as Colunas")
        lbl.setFont(QFont("Arial", 14, QFont.Bold))
        lbl.setStyleSheet("background: #4A90E2; color: white; padding: 10px;")
        lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(lbl)
        
        # Toolbar
        tool = QHBoxLayout()
        self.combo = QComboBox()
        self.combo.setMinimumWidth(300)
        self.combo.currentIndexChanged.connect(self.load_pipe_data)
        
        btn_refresh = QPushButton("🔄 Recarregar Pipes")
        btn_refresh.clicked.connect(self.load_pipes)
        
        btn_csv = QPushButton("📥 Exportar CSV")
        btn_csv.clicked.connect(self.export_csv)
        
        tool.addWidget(QLabel("Pipe:"))
        tool.addWidget(self.combo)
        tool.addWidget(btn_refresh)
        tool.addWidget(btn_csv)
        tool.addStretch()
        layout.addLayout(tool)
        
        # Toolbar Database
        tool_db = QHBoxLayout()
        self.combo_db = QComboBox()
        self.combo_db.setMinimumWidth(300)
        self.combo_db.currentIndexChanged.connect(self.load_database_data)
        
        btn_load_db = QPushButton("📂 Carregar Database")
        btn_load_db.clicked.connect(self.load_databases)
        
        btn_csv_db = QPushButton("📥 Exportar DB CSV")
        btn_csv_db.clicked.connect(self.export_database_csv)
        
        tool_db.addWidget(QLabel("Database:"))
        tool_db.addWidget(self.combo_db)
        tool_db.addWidget(btn_load_db)
        tool_db.addWidget(btn_csv_db)
        tool_db.addStretch()
        layout.addLayout(tool_db)
        
        # Toolbar Análise de Tempo de Entrega
        tool_analise = QHBoxLayout()
        btn_tempo_entrega = QPushButton("📊 Análise Tempo de Entrega")
        btn_tempo_entrega.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 8px 16px;")
        btn_tempo_entrega.clicked.connect(self.analisar_tempo_entrega)
        
        tool_analise.addWidget(btn_tempo_entrega)
        tool_analise.addStretch()
        layout.addLayout(tool_analise)
        
        # Status
        self.prog_bar = QProgressBar()
        self.prog_bar.setVisible(False)
        self.prog_bar.setRange(0, 0)
        layout.addWidget(self.prog_bar)
        
        self.lbl_status = QLabel("Iniciando...")
        self.lbl_status.setStyleSheet("padding: 5px; background: #eee;")
        layout.addWidget(self.lbl_status)
        
        # Tabela
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("alternate-background-color: #f5f5f5;")
        layout.addWidget(self.table)
        
    def authenticate(self):
        self.lbl_status.setText("Autenticando...")
        QApplication.processEvents()
        try:
            self.client = PipefyClient(self.CLIENT_ID, self.CLIENT_SECRET, self.TOKEN_URL)
            if self.client.authenticate():
                self.lbl_status.setText("✅ Autenticado.")
                self.load_pipes()
        except Exception as e:
            QMessageBox.critical(self, "Erro Auth", str(e))

    def load_pipes(self):
        self.prog_bar.setVisible(True)
        self.lbl_status.setText("Listando pipes...")
        self.combo.setEnabled(False)
        
        self.thread = LoadDataThread(self.client, pipe_ids=self.PIPE_IDS)
        self.thread.finished.connect(self.on_pipes_loaded)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_pipes_loaded(self, pipes):
        self.prog_bar.setVisible(False)
        self.combo.blockSignals(True)
        self.combo.clear()
        for p in pipes:
            self.combo.addItem(p['name'], p['id'])
        self.combo.blockSignals(False)
        self.combo.setEnabled(True)
        self.lbl_status.setText(f"✅ {len(pipes)} pipes listados.")
        if self.combo.count() > 0:
            self.load_pipe_data()

    def load_pipe_data(self):
        if self.combo.currentIndex() < 0: return
        pid = self.combo.currentData()
        name = self.combo.currentText()
        
        self.prog_bar.setVisible(True)
        self.lbl_status.setText(f"Baixando TODOS os cards de: {name}...")
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        
        self.thread = LoadDataThread(self.client, pipe_id=pid)
        self.thread.finished.connect(self.on_cards_loaded)
        self.thread.error.connect(self.on_error)
        self.thread.progress.connect(self.lbl_status.setText)
        self.thread.start()

    def on_cards_loaded(self, data):
        self.prog_bar.setVisible(False)
        self.current_cards = data['cards']
        self.lbl_status.setText(f"✅ Download concluído: {len(self.current_cards)} cards.")
        self.populate_dynamic_table(self.current_cards)

    def populate_dynamic_table(self, cards):
        if not cards:
            return

        # --- 1. Colunas de Sistema (Fixas) - EXPANDIDAS ---
        standard_cols = [
            'ID', 
            'URL',
            'Título', 
            'Fase Atual', 
            'Fase ID',
            'Tempo na Fase Atual (dias)',
            'Tempo na Fase Atual (horas)',
            'Responsáveis',
            'Emails Responsáveis',
            'Etiquetas',
            'Cores Etiquetas',
            'Vencimento',
            'Data Criação', 
            'Última Atualização',
            'Data Finalização',
            'Status',
            'Comentários',
            'Anexos',
            'Cards Relacionados (Pais)',
            'Cards Relacionados (Filhos)'
        ]
        
        # --- 2. Descobrir TODAS as Fases do Histórico ---
        all_phases_set = set()
        for c in cards:
            for ph in c.get('phases_history', []):
                phase_name = ph.get('phase', {}).get('name', '')
                if phase_name:
                    all_phases_set.add(phase_name)
        
        # Ordenar fases alfabeticamente
        all_phases_sorted = sorted(list(all_phases_set))
        
        # Criar colunas para cada fase (Entrada e Saída)
        phase_cols = []
        for phase_name in all_phases_sorted:
            phase_cols.append(f"[{phase_name}] Entrada")
            phase_cols.append(f"[{phase_name}] Saída")
            phase_cols.append(f"[{phase_name}] Tempo (horas)")
        
        # Guardar lista de fases para uso posterior
        self.phase_names_order = all_phases_sorted
        
        # --- 3. Descobrir TODOS os Campos Customizados ---
        custom_fields_set = set()
        for c in cards:
            for field in c.get('fields', []):
                field_name = field.get('name')
                if field_name:
                    custom_fields_set.add(field_name)
        
        custom_cols = sorted(list(custom_fields_set))
        self.dynamic_headers = standard_cols + phase_cols + custom_cols
        
        # Configurar Tabela
        self.table.setColumnCount(len(self.dynamic_headers))
        self.table.setHorizontalHeaderLabels(self.dynamic_headers)
        self.table.setRowCount(len(cards))
        
        # --- 3. Preencher Dados ---
        for row, card in enumerate(cards):
            # Mapa para busca rápida de campos customizados
            field_map = {f['name']: f for f in card.get('fields', [])}
            
            col_idx = 0
            
            # ID
            self.table.setItem(row, col_idx, QTableWidgetItem(str(card.get('id', ''))))
            col_idx += 1
            
            # URL
            self.table.setItem(row, col_idx, QTableWidgetItem(card.get('url', '')))
            col_idx += 1
            
            # Título
            self.table.setItem(row, col_idx, QTableWidgetItem(card.get('title', '')))
            col_idx += 1
            
            # Fase Atual
            phase = card.get('current_phase', {})
            phase_name = phase.get('name', '') if phase else ''
            self.table.setItem(row, col_idx, QTableWidgetItem(phase_name))
            col_idx += 1
            
            # Fase ID
            phase_id = phase.get('id', '') if phase else ''
            self.table.setItem(row, col_idx, QTableWidgetItem(phase_id))
            col_idx += 1
            
            # Calcular tempo na fase atual
            phases_history = card.get('phases_history', [])
            time_in_current_phase_days = ""
            time_in_current_phase_hours = ""
            
            if phases_history and phase_id:
                # Encontrar a fase atual no histórico
                for ph in phases_history:
                    if ph.get('phase', {}).get('id') == phase_id:
                        first_time_in = ph.get('firstTimeIn')
                        last_time_out = ph.get('lastTimeOut')
                        
                        if first_time_in:
                            try:
                                from datetime import datetime
                                start = datetime.fromisoformat(first_time_in.replace('Z', '+00:00'))
                                
                                # Se ainda está na fase (lastTimeOut é None)
                                if not last_time_out:
                                    end = datetime.now(start.tzinfo)
                                else:
                                    end = datetime.fromisoformat(last_time_out.replace('Z', '+00:00'))
                                
                                time_diff = end - start
                                days = time_diff.days
                                hours = time_diff.total_seconds() / 3600
                                
                                time_in_current_phase_days = f"{days}"
                                time_in_current_phase_hours = f"{hours:.1f}"
                            except:
                                pass
                        break
            
            # Tempo na Fase Atual (dias)
            self.table.setItem(row, col_idx, QTableWidgetItem(time_in_current_phase_days))
            col_idx += 1
            
            # Tempo na Fase Atual (horas)
            self.table.setItem(row, col_idx, QTableWidgetItem(time_in_current_phase_hours))
            col_idx += 1
            
            # Responsáveis
            assignees = card.get('assignees', [])
            ass_names = ", ".join([a.get('name', '') for a in assignees])
            self.table.setItem(row, col_idx, QTableWidgetItem(ass_names))
            col_idx += 1
            
            # Emails Responsáveis
            ass_emails = ", ".join([a.get('email', '') for a in assignees])
            self.table.setItem(row, col_idx, QTableWidgetItem(ass_emails))
            col_idx += 1
            
            # Etiquetas
            labels = card.get('labels', [])
            label_names = ", ".join([l.get('name', '') for l in labels])
            self.table.setItem(row, col_idx, QTableWidgetItem(label_names))
            col_idx += 1
            
            # Cores Etiquetas
            label_colors = ", ".join([l.get('color', '') for l in labels])
            self.table.setItem(row, col_idx, QTableWidgetItem(label_colors))
            col_idx += 1
            
            # Vencimento
            due = card.get('due_date', '')
            if due: 
                due = due[:10]  # YYYY-MM-DD
            self.table.setItem(row, col_idx, QTableWidgetItem(due))
            col_idx += 1
            
            # Data Criação
            created = card.get('created_at', '')
            if created:
                created = created[:19].replace('T', ' ')  # YYYY-MM-DD HH:MM:SS
            self.table.setItem(row, col_idx, QTableWidgetItem(created))
            col_idx += 1
            
            # Última Atualização
            updated = card.get('updated_at', '')
            if updated:
                updated = updated[:19].replace('T', ' ')
            self.table.setItem(row, col_idx, QTableWidgetItem(updated))
            col_idx += 1
            
            # Data Finalização
            finished = card.get('finished_at', '')
            if finished:
                finished = finished[:19].replace('T', ' ')
            self.table.setItem(row, col_idx, QTableWidgetItem(finished))
            col_idx += 1
            
            # Status
            status = "Finalizado" if card.get('finished_at') else "Em Andamento"
            self.table.setItem(row, col_idx, QTableWidgetItem(status))
            col_idx += 1
            
            # Comentários
            self.table.setItem(row, col_idx, QTableWidgetItem(str(card.get('comments_count', 0))))
            col_idx += 1
            
            # Anexos
            self.table.setItem(row, col_idx, QTableWidgetItem(str(card.get('attachments_count', 0))))
            col_idx += 1
            
            # Cards Relacionados (Pais)
            parent_rels = card.get('parent_relations', [])
            parent_info = ", ".join([p.get('name', '') for p in parent_rels])
            self.table.setItem(row, col_idx, QTableWidgetItem(parent_info))
            col_idx += 1
            
            # Cards Relacionados (Filhos)
            child_rels = card.get('child_relations', [])
            child_info = ", ".join([c.get('name', '') for c in child_rels])
            self.table.setItem(row, col_idx, QTableWidgetItem(child_info))
            col_idx += 1
            
            # --- Histórico de Fases (colunas separadas por fase) ---
            # Criar mapa do histórico de fases para este card
            phase_history_map = {}
            for ph in phases_history:
                phase_name_hist = ph.get('phase', {}).get('name', '')
                if phase_name_hist:
                    first_in = ph.get('firstTimeIn', '')
                    last_out = ph.get('lastTimeOut', '')
                    
                    # Calcular tempo na fase
                    time_in_phase = ""
                    if first_in:
                        try:
                            start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                            if last_out:
                                end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                            else:
                                end = datetime.now(start.tzinfo)
                            hours = (end - start).total_seconds() / 3600
                            time_in_phase = f"{hours:.1f}"
                        except:
                            pass
                    
                    # Formatar datas
                    if first_in:
                        first_in = first_in[:19].replace('T', ' ')
                    if last_out:
                        last_out = last_out[:19].replace('T', ' ')
                    
                    phase_history_map[phase_name_hist] = {
                        'entrada': first_in,
                        'saida': last_out if last_out else '',
                        'tempo': time_in_phase
                    }
            
            # Preencher colunas de cada fase
            for phase_name in self.phase_names_order:
                phase_data = phase_history_map.get(phase_name, {})
                
                # Entrada
                self.table.setItem(row, col_idx, QTableWidgetItem(phase_data.get('entrada', '')))
                col_idx += 1
                
                # Saída
                self.table.setItem(row, col_idx, QTableWidgetItem(phase_data.get('saida', '')))
                col_idx += 1
                
                # Tempo (horas)
                self.table.setItem(row, col_idx, QTableWidgetItem(phase_data.get('tempo', '')))
                col_idx += 1
            
            # --- Campos Customizados ---
            for custom_field_name in custom_cols:
                field_data = field_map.get(custom_field_name, {})
                val = field_data.get('value', '')
                
                # Tratamento especial para diferentes tipos de valores
                if isinstance(val, list):
                    # Para campos multi-seleção
                    val = ", ".join([str(v) for v in val])
                elif isinstance(val, dict):
                    # Para campos mais complexos
                    val = json.dumps(val, ensure_ascii=False)
                elif val is None:
                    val = ""
                else:
                    val = str(val)
                    
                self.table.setItem(row, col_idx, QTableWidgetItem(val))
                col_idx += 1
                
        self.table.resizeColumnsToContents()
        
        # Ajustar largura máxima das colunas para melhor visualização
        for col in range(self.table.columnCount()):
            if self.table.columnWidth(col) > 300:
                self.table.setColumnWidth(col, 300)

    def export_csv(self):
        if not self.current_cards or not self.dynamic_headers:
            QMessageBox.warning(self, "Aviso", "Não há dados para exportar.")
            return
            
        try:
            filename = f"relatorio_pipefy_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(self.dynamic_headers)
                
                for row in range(self.table.rowCount()):
                    row_data = []
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else "")
                    writer.writerow(row_data)
                    
            QMessageBox.information(self, "Sucesso", f"✅ Arquivo exportado com sucesso!\n\n{filename}\n\nTotal: {self.table.rowCount()} linhas x {self.table.columnCount()} colunas")
        except Exception as e:
            QMessageBox.critical(self, "Erro CSV", f"Erro ao exportar:\n{str(e)}")

    # ========================================================================
    # MÉTODOS PARA DATABASE
    # ========================================================================
    def load_databases(self):
        self.prog_bar.setVisible(True)
        self.lbl_status.setText("Carregando databases...")
        self.combo_db.setEnabled(False)
        
        self.thread = LoadDataThread(self.client, database_ids=self.DATABASE_IDS)
        self.thread.finished.connect(self.on_databases_loaded)
        self.thread.error.connect(self.on_error)
        self.thread.start()

    def on_databases_loaded(self, result):
        self.prog_bar.setVisible(False)
        if result.get('type') == 'database_list':
            databases = result.get('data', [])
            self.combo_db.blockSignals(True)
            self.combo_db.clear()
            for db in databases:
                self.combo_db.addItem(db.get('name', f"Database {db.get('id')}"), db.get('id'))
            self.combo_db.blockSignals(False)
            self.combo_db.setEnabled(True)
            self.lbl_status.setText(f"✅ {len(databases)} database(s) listadas.")
            if self.combo_db.count() > 0:
                self.load_database_data()
        elif result.get('type') == 'database':
            self.on_database_records_loaded(result.get('data'))

    def load_database_data(self):
        if self.combo_db.currentIndex() < 0:
            return
        db_id = self.combo_db.currentData()
        name = self.combo_db.currentText()
        
        self.prog_bar.setVisible(True)
        self.lbl_status.setText(f"Baixando registros de: {name}...")
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        
        self.thread = LoadDataThread(self.client, database_id=db_id)
        self.thread.finished.connect(self.on_databases_loaded)
        self.thread.error.connect(self.on_error)
        self.thread.progress.connect(self.lbl_status.setText)
        self.thread.start()

    def on_database_records_loaded(self, data):
        self.prog_bar.setVisible(False)
        self.current_records = data.get('records', [])
        db_name = data.get('name', 'Database')
        self.lbl_status.setText(f"✅ Download concluído: {len(self.current_records)} registros de '{db_name}'.")
        self.populate_database_table(self.current_records)

    def populate_database_table(self, records):
        if not records:
            self.lbl_status.setText("⚠️ Nenhum registro encontrado na database.")
            return
        
        # Colunas de sistema
        system_cols = ['ID', 'Título', 'Status', 'Criado em', 'Atualizado em', 'URL']
        
        # Descobrir todos os campos dinâmicos
        field_names = set()
        for record in records:
            for rf in record.get('record_fields', []):
                field_name = rf.get('name') or (rf.get('field', {}) or {}).get('label', '')
                if field_name:
                    field_names.add(field_name)
        
        field_names = sorted(field_names)
        all_cols = system_cols + field_names
        
        self.table.setColumnCount(len(all_cols))
        self.table.setHorizontalHeaderLabels(all_cols)
        self.table.setRowCount(len(records))
        
        for row_idx, record in enumerate(records):
            # Colunas de sistema
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(record.get('id', ''))))
            self.table.setItem(row_idx, 1, QTableWidgetItem(str(record.get('title', ''))))
            # Status é um objeto com id e name
            status = record.get('status')
            status_name = status.get('name', '') if isinstance(status, dict) else str(status or '')
            self.table.setItem(row_idx, 2, QTableWidgetItem(status_name))
            self.table.setItem(row_idx, 3, QTableWidgetItem(str(record.get('created_at', ''))))
            self.table.setItem(row_idx, 4, QTableWidgetItem(str(record.get('updated_at', ''))))
            self.table.setItem(row_idx, 5, QTableWidgetItem(str(record.get('url', ''))))
            
            # Campos dinâmicos
            field_map = {}
            for rf in record.get('record_fields', []):
                field_name = rf.get('name') or (rf.get('field', {}) or {}).get('label', '')
                if field_name:
                    field_map[field_name] = rf.get('value', '')
            
            for col_idx, field_name in enumerate(field_names, start=len(system_cols)):
                value = field_map.get(field_name, '')
                # Se for JSON (lista/dict), converter para string
                if isinstance(value, (list, dict)):
                    value = json.dumps(value, ensure_ascii=False)
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(value) if value else ''))
        
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.dynamic_headers = all_cols

    def export_database_csv(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "Aviso", "Nenhum dado para exportar.")
            return
        
        db_name = self.combo_db.currentText() if self.combo_db.currentIndex() >= 0 else "database"
        safe_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in db_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"database_{safe_name}_{timestamp}.csv"
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                
                headers = []
                for col in range(self.table.columnCount()):
                    headers.append(self.table.horizontalHeaderItem(col).text())
                writer.writerow(headers)
                
                for row in range(self.table.rowCount()):
                    row_data = []
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else '')
                    writer.writerow(row_data)
            
            QMessageBox.information(self, "Sucesso", f"✅ Database exportada!\n\n{filename}\n\nTotal: {self.table.rowCount()} registros")
        except Exception as e:
            QMessageBox.critical(self, "Erro CSV", f"Erro ao exportar:\n{str(e)}")

    # ========================================================================
    # ANÁLISE DE TEMPO DE ENTREGA
    # ========================================================================
    def analisar_tempo_entrega(self):
        """Inicia a análise de tempo de entrega"""
        self.prog_bar.setVisible(True)
        self.lbl_status.setText("Iniciando análise de tempo de entrega...")
        
        # Criar thread para buscar dados
        self.thread_analise = AnaliseTempoEntregaThread(self.client, self.DATABASE_ORDENS_COMPRA, self.PIPE_IDS)
        self.thread_analise.finished.connect(self.on_analise_finished)
        self.thread_analise.error.connect(self.on_error)
        self.thread_analise.progress.connect(self.lbl_status.setText)
        self.thread_analise.start()
    
    def on_analise_finished(self, resultado):
        """Exibe os resultados da análise de tempo de entrega"""
        self.prog_bar.setVisible(False)
        self.lbl_status.setText("✅ Análise concluída!")
        
        # Mostrar diálogo com resultados
        dialog = TempoEntregaDialog(resultado, self)
        dialog.exec_()

    def on_error(self, msg):
        self.prog_bar.setVisible(False)
        self.lbl_status.setText("❌ Erro")
        print("ERRO:", msg)
        QMessageBox.critical(self, "Erro", msg)


# ============================================================================
# THREAD PARA ANÁLISE DE TEMPO DE ENTREGA
# ============================================================================
class AnaliseTempoEntregaThread(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    
    def __init__(self, client, database_id, pipe_ids):
        super().__init__()
        self.client = client
        self.database_id = database_id
        self.pipe_ids = pipe_ids
    
    def run(self):
        try:
            # 1. Buscar registros da database (Ordens de Compra)
            self.progress.emit("Buscando ordens de compra na database...")
            db_data = self.client.get_all_database_records(
                self.database_id,
                progress_callback=lambda msg: self.progress.emit(f"Database: {msg}")
            )
            
            # Mapear ordens de compra: número -> {data_emissao, solicitante}
            ordens_compra = {}
            for record in db_data.get('records', []):
                # Extrair campos do registro
                field_map = {}
                for rf in record.get('record_fields', []):
                    field_name = rf.get('name') or (rf.get('field', {}) or {}).get('label', '')
                    if field_name:
                        field_map[field_name] = rf.get('value', '')
                
                # Buscar número da ordem de compra e solicitante
                # Tentar diferentes nomes de campo possíveis
                numero_oc = None
                for campo in ['Ordem de compra', 'Ordem de Compra', 'OC', 'Numero OC', 'Número OC', 'N° OC']:
                    if campo in field_map and field_map[campo]:
                        numero_oc = str(field_map[campo]).strip()
                        break
                
                # Se não encontrou em campo específico, usar o título
                if not numero_oc:
                    numero_oc = str(record.get('title', '')).strip()
                
                if numero_oc:
                    # Buscar solicitante
                    solicitante = None
                    for campo in ['Solicitante', 'Requisitante', 'Solicitado por', 'Usuário']:
                        if campo in field_map and field_map[campo]:
                            solicitante = str(field_map[campo]).strip()
                            break
                    
                    # Buscar requisição de compra
                    requisicao = None
                    for campo in ['Requisição de compra', 'Requisicao de compra', 'Requisição', 'Requisicao', 'RC', 'N° RC', 'N° Requisição']:
                        if campo in field_map and field_map[campo]:
                            requisicao = str(field_map[campo]).strip()
                            break
                    
                    # Data de emissão = updated_at do registro
                    data_emissao = record.get('updated_at', '') or record.get('created_at', '')
                    
                    ordens_compra[numero_oc] = {
                        'data_emissao': data_emissao,
                        'solicitante': solicitante or 'Não identificado',
                        'requisicao': requisicao or '',
                        'id': record.get('id')
                    }
            
            self.progress.emit(f"Encontradas {len(ordens_compra)} ordens de compra. Buscando NF-Es...")
            
            # 2. Buscar cards de todos os pipes para encontrar Lançamento NF-E
            entregas_concluidas = []
            
            for pipe_id in self.pipe_ids:
                self.progress.emit(f"Analisando pipe {pipe_id}...")
                
                try:
                    pipe_data = self.client.get_all_pipe_cards(
                        pipe_id,
                        progress_callback=lambda msg: self.progress.emit(f"Pipe {pipe_id}: {msg}")
                    )
                    
                    pipe_name = pipe_data.get('name', '')
                    
                    # Verificar se é um pipe relacionado a NF-E ou se tem cards com NF concluída
                    for card in pipe_data.get('cards', []):
                        # Verificar se o card está finalizado
                        finished_at = card.get('finished_at')
                        if not finished_at:
                            continue
                        
                        # Buscar número da ordem de compra nos campos do card
                        field_map = {f['name']: f['value'] for f in card.get('fields', []) if f.get('name')}
                        
                        numero_oc_card = None
                        for campo in ['Ordem de compra', 'Ordem de Compra', 'OC', 'N° OC', 'Número OC', 
                                       'Numero OC', 'N° Ordem de compra', 'Ordem_compra']:
                            if campo in field_map and field_map[campo]:
                                numero_oc_card = str(field_map[campo]).strip()
                                break
                        
                        if numero_oc_card and numero_oc_card in ordens_compra:
                            ordem = ordens_compra[numero_oc_card]
                            entregas_concluidas.append({
                                'numero_oc': numero_oc_card,
                                'requisicao': ordem['requisicao'],
                                'data_emissao': ordem['data_emissao'],
                                'data_entrega': finished_at,
                                'solicitante': ordem['solicitante'],
                                'pipe_name': pipe_name,
                                'card_title': card.get('title', '')
                            })
                except Exception as e:
                    print(f"Erro ao processar pipe {pipe_id}: {e}")
                    continue
            
            self.progress.emit(f"Calculando tempos de entrega...")
            
            # 3. Calcular tempos de entrega
            resultados = {
                'entregas': [],
                'media_geral_dias': 0,
                'media_geral_horas': 0,
                'por_solicitante': {},
                'total_entregas': 0
            }
            
            tempos_totais = []
            tempos_por_solicitante = {}
            
            for entrega in entregas_concluidas:
                try:
                    data_emissao = entrega['data_emissao']
                    data_entrega = entrega['data_entrega']
                    
                    # Converter datas
                    if isinstance(data_emissao, str):
                        data_emissao = datetime.fromisoformat(data_emissao.replace('Z', '+00:00'))
                    if isinstance(data_entrega, str):
                        data_entrega = datetime.fromisoformat(data_entrega.replace('Z', '+00:00'))
                    
                    # Calcular diferença
                    diferenca = data_entrega - data_emissao
                    dias = diferenca.days
                    horas = diferenca.total_seconds() / 3600
                    
                    # Ignorar valores negativos (dados inconsistentes)
                    if dias < 0:
                        continue
                    
                    entrega['tempo_dias'] = dias
                    entrega['tempo_horas'] = round(horas, 1)
                    resultados['entregas'].append(entrega)
                    
                    tempos_totais.append(horas)
                    
                    # Agrupar por solicitante
                    solicitante = entrega['solicitante']
                    if solicitante not in tempos_por_solicitante:
                        tempos_por_solicitante[solicitante] = []
                    tempos_por_solicitante[solicitante].append(horas)
                    
                except Exception as e:
                    print(f"Erro ao calcular tempo para OC {entrega.get('numero_oc')}: {e}")
                    continue
            
            # Calcular médias
            if tempos_totais:
                media_horas = sum(tempos_totais) / len(tempos_totais)
                resultados['media_geral_horas'] = round(media_horas, 1)
                resultados['media_geral_dias'] = round(media_horas / 24, 1)
            
            resultados['total_entregas'] = len(resultados['entregas'])
            
            # Médias por solicitante
            for solicitante, tempos in tempos_por_solicitante.items():
                media_horas = sum(tempos) / len(tempos)
                resultados['por_solicitante'][solicitante] = {
                    'media_horas': round(media_horas, 1),
                    'media_dias': round(media_horas / 24, 1),
                    'quantidade': len(tempos)
                }
            
            self.finished.emit(resultados)
            
        except Exception as e:
            self.error.emit(str(e))


# ============================================================================
# DIÁLOGO PARA EXIBIR RESULTADOS DE TEMPO DE ENTREGA
# ============================================================================
class TempoEntregaDialog(QDialog):
    def __init__(self, resultado, parent=None):
        super().__init__(parent)
        self.resultado = resultado
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("📊 Análise de Tempo de Entrega")
        self.resize(900, 700)
        
        layout = QVBoxLayout(self)
        
        # Header
        lbl_header = QLabel("📦 Análise de Tempo de Entrega de Materiais")
        lbl_header.setFont(QFont("Arial", 14, QFont.Bold))
        lbl_header.setStyleSheet("background: #28a745; color: white; padding: 15px; border-radius: 5px;")
        lbl_header.setAlignment(Qt.AlignCenter)
        layout.addWidget(lbl_header)
        
        # Resumo Geral
        frame_resumo = QWidget()
        frame_resumo.setStyleSheet("background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 5px; padding: 10px;")
        resumo_layout = QVBoxLayout(frame_resumo)
        
        lbl_titulo_resumo = QLabel("📈 RESUMO GERAL")
        lbl_titulo_resumo.setFont(QFont("Arial", 12, QFont.Bold))
        resumo_layout.addWidget(lbl_titulo_resumo)
        
        total = self.resultado.get('total_entregas', 0)
        media_dias = self.resultado.get('media_geral_dias', 0)
        media_horas = self.resultado.get('media_geral_horas', 0)
        
        resumo_text = f"""
        <table style='font-size: 14px;'>
        <tr><td><b>Total de Entregas Analisadas:</b></td><td style='padding-left: 20px;'>{total}</td></tr>
        <tr><td><b>Tempo Médio de Entrega:</b></td><td style='padding-left: 20px;'><span style='color: #007bff; font-weight: bold;'>{media_dias} dias</span> ({media_horas} horas)</td></tr>
        </table>
        """
        lbl_resumo = QLabel(resumo_text)
        lbl_resumo.setTextFormat(Qt.RichText)
        resumo_layout.addWidget(lbl_resumo)
        
        layout.addWidget(frame_resumo)
        
        # Abas para detalhes
        tabs = QTabWidget()
        
        # Tab 1: Por Solicitante
        tab_solicitante = QWidget()
        tab_solicitante_layout = QVBoxLayout(tab_solicitante)
        
        table_solicitante = QTableWidget()
        por_solicitante = self.resultado.get('por_solicitante', {})
        
        table_solicitante.setColumnCount(4)
        table_solicitante.setHorizontalHeaderLabels(['Solicitante', 'Qtd Entregas', 'Média (Dias)', 'Média (Horas)'])
        table_solicitante.setRowCount(len(por_solicitante))
        
        # Ordenar por média de dias (do mais rápido ao mais lento)
        solicitantes_ordenados = sorted(por_solicitante.items(), key=lambda x: x[1]['media_dias'])
        
        for row, (solicitante, dados) in enumerate(solicitantes_ordenados):
            table_solicitante.setItem(row, 0, QTableWidgetItem(solicitante))
            table_solicitante.setItem(row, 1, QTableWidgetItem(str(dados['quantidade'])))
            
            item_dias = QTableWidgetItem(f"{dados['media_dias']}")
            item_dias.setTextAlignment(Qt.AlignCenter)
            # Colorir baseado no tempo
            if dados['media_dias'] <= 7:
                item_dias.setBackground(QColor(200, 255, 200))  # Verde claro
            elif dados['media_dias'] <= 15:
                item_dias.setBackground(QColor(255, 255, 200))  # Amarelo claro
            else:
                item_dias.setBackground(QColor(255, 200, 200))  # Vermelho claro
            table_solicitante.setItem(row, 2, item_dias)
            
            item_horas = QTableWidgetItem(f"{dados['media_horas']}")
            item_horas.setTextAlignment(Qt.AlignCenter)
            table_solicitante.setItem(row, 3, item_horas)
        
        table_solicitante.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table_solicitante.setAlternatingRowColors(True)
        tab_solicitante_layout.addWidget(table_solicitante)
        
        tabs.addTab(tab_solicitante, "👤 Por Solicitante")
        
        # Tab 2: Detalhes de Entregas
        tab_detalhes = QWidget()
        tab_detalhes_layout = QVBoxLayout(tab_detalhes)
        
        table_detalhes = QTableWidget()
        entregas = self.resultado.get('entregas', [])
        por_solicitante = self.resultado.get('por_solicitante', {})
        
        table_detalhes.setColumnCount(9)
        table_detalhes.setHorizontalHeaderLabels([
            'Requisição de Compra', 'Ordem de Compra', 'Solicitante', 
            'Data Emissão', 'Data Entrega', 'Tempo (Dias)', 'Tempo (Horas)',
            'Média Solicitante (Dias)', 'Diferença da Média'
        ])
        table_detalhes.setRowCount(len(entregas))
        
        for row, entrega in enumerate(entregas):
            # Requisição de Compra
            requisicao = str(entrega.get('requisicao', '')) or 'N/A'
            item_req = QTableWidgetItem(requisicao)
            item_req.setFont(QFont("Arial", 9, QFont.Bold))
            table_detalhes.setItem(row, 0, item_req)
            
            # Ordem de Compra
            table_detalhes.setItem(row, 1, QTableWidgetItem(str(entrega.get('numero_oc', ''))))
            
            # Solicitante
            solicitante = str(entrega.get('solicitante', ''))
            table_detalhes.setItem(row, 2, QTableWidgetItem(solicitante))
            
            # Data Emissão
            data_emissao = str(entrega.get('data_emissao', ''))[:19].replace('T', ' ')
            table_detalhes.setItem(row, 3, QTableWidgetItem(data_emissao))
            
            # Data Entrega
            data_entrega = str(entrega.get('data_entrega', ''))[:19].replace('T', ' ')
            table_detalhes.setItem(row, 4, QTableWidgetItem(data_entrega))
            
            # Tempo (Dias)
            tempo_dias = entrega.get('tempo_dias', 0)
            item_dias = QTableWidgetItem(str(tempo_dias))
            item_dias.setTextAlignment(Qt.AlignCenter)
            table_detalhes.setItem(row, 5, item_dias)
            
            # Tempo (Horas)
            item_horas = QTableWidgetItem(str(entrega.get('tempo_horas', '')))
            item_horas.setTextAlignment(Qt.AlignCenter)
            table_detalhes.setItem(row, 6, item_horas)
            
            # Média do Solicitante (Dias)
            media_solicitante = por_solicitante.get(solicitante, {}).get('media_dias', 0)
            item_media = QTableWidgetItem(f"{media_solicitante}")
            item_media.setTextAlignment(Qt.AlignCenter)
            table_detalhes.setItem(row, 7, item_media)
            
            # Diferença da Média (positivo = mais lento, negativo = mais rápido)
            diferenca = round(tempo_dias - media_solicitante, 1)
            item_diff = QTableWidgetItem(f"{diferenca:+.1f}")
            item_diff.setTextAlignment(Qt.AlignCenter)
            if diferenca < 0:
                item_diff.setBackground(QColor(200, 255, 200))  # Verde - mais rápido que a média
                item_diff.setToolTip("Mais rápido que a média do solicitante")
            elif diferenca > 0:
                item_diff.setBackground(QColor(255, 200, 200))  # Vermelho - mais lento que a média
                item_diff.setToolTip("Mais lento que a média do solicitante")
            else:
                item_diff.setBackground(QColor(255, 255, 200))  # Amarelo - na média
                item_diff.setToolTip("Na média do solicitante")
            table_detalhes.setItem(row, 8, item_diff)
        
        table_detalhes.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table_detalhes.setAlternatingRowColors(True)
        table_detalhes.setSortingEnabled(True)  # Permitir ordenação por coluna
        tab_detalhes_layout.addWidget(table_detalhes)
        
        tabs.addTab(tab_detalhes, "📋 Detalhes das Entregas")
        
        layout.addWidget(tabs)
        
        # Botões
        btn_layout = QHBoxLayout()
        
        btn_exportar = QPushButton("📥 Exportar CSV")
        btn_exportar.clicked.connect(self.exportar_csv)
        
        btn_fechar = QPushButton("Fechar")
        btn_fechar.clicked.connect(self.close)
        
        btn_layout.addWidget(btn_exportar)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_fechar)
        
        layout.addLayout(btn_layout)
    
    def exportar_csv(self):
        """Exporta os resultados para CSV"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"analise_tempo_entrega_{timestamp}.csv"
            
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                
                # Resumo
                writer.writerow(['=== RESUMO GERAL ==='])
                writer.writerow(['Total de Entregas', self.resultado.get('total_entregas', 0)])
                writer.writerow(['Média Geral (Dias)', self.resultado.get('media_geral_dias', 0)])
                writer.writerow(['Média Geral (Horas)', self.resultado.get('media_geral_horas', 0)])
                writer.writerow([])
                
                # Por Solicitante
                writer.writerow(['=== MÉDIA POR SOLICITANTE ==='])
                writer.writerow(['Solicitante', 'Quantidade', 'Média (Dias)', 'Média (Horas)'])
                for solicitante, dados in self.resultado.get('por_solicitante', {}).items():
                    writer.writerow([solicitante, dados['quantidade'], dados['media_dias'], dados['media_horas']])
                writer.writerow([])
                
                # Detalhes
                writer.writerow(['=== DETALHES DAS ENTREGAS ==='])
                writer.writerow(['Requisição de Compra', 'Ordem de Compra', 'Solicitante', 'Data Emissão', 'Data Entrega', 'Tempo (Dias)', 'Tempo (Horas)', 'Média Solicitante (Dias)', 'Diferença da Média'])
                por_solicitante = self.resultado.get('por_solicitante', {})
                for entrega in self.resultado.get('entregas', []):
                    data_emissao = str(entrega.get('data_emissao', ''))[:19].replace('T', ' ')
                    data_entrega = str(entrega.get('data_entrega', ''))[:19].replace('T', ' ')
                    solicitante = entrega.get('solicitante', '')
                    tempo_dias = entrega.get('tempo_dias', 0)
                    media_solicitante = por_solicitante.get(solicitante, {}).get('media_dias', 0)
                    diferenca = round(tempo_dias - media_solicitante, 1)
                    writer.writerow([
                        entrega.get('requisicao', '') or 'N/A',
                        entrega.get('numero_oc', ''),
                        solicitante,
                        data_emissao,
                        data_entrega,
                        tempo_dias,
                        entrega.get('tempo_horas', ''),
                        media_solicitante,
                        f"{diferenca:+.1f}"
                    ])
            
            QMessageBox.information(self, "Sucesso", f"✅ Relatório exportado!\n\n{filename}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = PipefyReportApp()
    win.show()
    sys.exit(app.exec_())