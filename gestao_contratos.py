"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                     GESTÃO DE CONTRATOS E MEDIÇÃO                          ║
║                     Sistema de Gerenciamento v1.0                          ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import json
import shutil
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
import locale

try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        pass

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QPushButton, QLabel, QFileDialog, QMessageBox,
    QFrame, QScrollArea, QSplitter, QProgressBar, QComboBox,
    QDateEdit, QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QDialogButtonBox, QFormLayout, QSpinBox, QDoubleSpinBox,
    QGroupBox, QCheckBox, QTabWidget, QTextEdit, QProgressDialog,
    QAction, QInputDialog, QSizePolicy
)
from PyQt5.QtCore import Qt, QUrl, QDate, QTimer, pyqtSignal, QThread, QSize, QObject, pyqtSlot
from PyQt5.QtGui import QIcon, QFont, QColor, QPalette, QPixmap, QDesktopServices
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWebChannel import QWebChannel

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTES E CONFIGURAÇÕES
# ═══════════════════════════════════════════════════════════════════════════════

APP_NAME = "Gestão de Contratos e Medição"
APP_VERSION = "1.0.0"

COLORS = {
    'primary': '#2d3748',
    'primary_dark': '#5a67d8',
    'secondary': '#1a202c',
    'success': '#48bb78',
    'warning': '#ed8936',
    'danger': '#f56565',
    'info': '#4299e1',
    'dark': '#1a202c',
    'light': '#f7fafc',
}

# Departamentos padrão
DEFAULT_DEPARTMENTS = [
    "Administrativo",
    "Comercial",
    "Aterro",
    "ETE",
    "UVE",
    "Manutenção",
    "Triagem",
    "Termogas",
    "Departamento Pessoal",
    "Financeiro",
]

# Ícones FontAwesome para cada departamento
DEPARTMENT_ICONS = {
    "Administrativo": "fa-building",
    "Comercial": "fa-handshake",
    "Aterro": "fa-mountain",
    "ETE": "fa-water",
    "UVE": "fa-fire",
    "Manutenção": "fa-wrench",
    "Triagem": "fa-recycle",
    "Termogas": "fa-gas-pump",
    "Departamento Pessoal": "fa-users",
    "Financeiro": "fa-dollar-sign",
}

# Cores únicas por departamento
DEPARTMENT_COLORS = {
    "Administrativo": "#5a67d8",
    "Comercial": "#48bb78",
    "Aterro": "#ed8936",
    "ETE": "#4299e1",
    "UVE": "#f56565",
    "Manutenção": "#9f7aea",
    "Triagem": "#38b2ac",
    "Termogas": "#e53e3e",
    "Departamento Pessoal": "#667eea",
    "Financeiro": "#d69e2e",
}

# Diretório de dados
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contratos_data")
CONTRACTS_FILE = os.path.join(DATA_DIR, "contratos.json")
ATTACHMENTS_DIR = os.path.join(DATA_DIR, "anexos")
CONFIG_FILE = os.path.join(DATA_DIR, "config_contratos.json")


# ═══════════════════════════════════════════════════════════════════════════════
# GERENCIADOR DE DADOS
# ═══════════════════════════════════════════════════════════════════════════════

class ContractDataManager:
    """Gerencia persistência de dados de contratos."""
    
    def __init__(self):
        self._ensure_dirs()
        self.contracts = self._load_contracts()
        self.config = self._load_config()
    
    def _ensure_dirs(self):
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(ATTACHMENTS_DIR, exist_ok=True)
    
    def _load_contracts(self) -> dict:
        if os.path.exists(CONTRACTS_FILE):
            try:
                with open(CONTRACTS_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {"departments": {}, "contracts": []}
        return {"departments": {}, "contracts": []}
    
    def _save_contracts(self):
        with open(CONTRACTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.contracts, f, ensure_ascii=False, indent=2)
    
    def _load_config(self) -> dict:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {"custom_departments": []}
        return {"custom_departments": []}
    
    def _save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)
    
    def get_all_departments(self) -> list:
        custom = self.config.get("custom_departments", [])
        return DEFAULT_DEPARTMENTS + custom
    
    def add_department(self, name: str) -> bool:
        if name in self.get_all_departments():
            return False
        custom = self.config.get("custom_departments", [])
        custom.append(name)
        self.config["custom_departments"] = custom
        self._save_config()
        return True
    
    def remove_department(self, name: str) -> bool:
        custom = self.config.get("custom_departments", [])
        if name in custom:
            custom.remove(name)
            self.config["custom_departments"] = custom
            self._save_config()
            return True
        return False
    
    def get_contracts_by_department(self, department: str) -> list:
        return [c for c in self.contracts.get("contracts", []) if c.get("departamento") == department]
    
    def get_all_contracts(self) -> list:
        return self.contracts.get("contracts", [])
    
    def add_contract(self, contract_data: dict) -> str:
        contract_data["id"] = datetime.now().strftime("%Y%m%d%H%M%S%f")
        contract_data["criado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        contract_data["atualizado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        contract_data["anexos"] = contract_data.get("anexos", [])
        contract_data["medicoes"] = contract_data.get("medicoes", [])
        
        if "contracts" not in self.contracts:
            self.contracts["contracts"] = []
        self.contracts["contracts"].append(contract_data)
        self._save_contracts()
        return contract_data["id"]
    
    def update_contract(self, contract_id: str, updates: dict):
        for i, c in enumerate(self.contracts.get("contracts", [])):
            if c.get("id") == contract_id:
                self.contracts["contracts"][i].update(updates)
                self.contracts["contracts"][i]["atualizado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self._save_contracts()
                return True
        return False
    
    def delete_contract(self, contract_id: str) -> bool:
        contracts = self.contracts.get("contracts", [])
        for i, c in enumerate(contracts):
            if c.get("id") == contract_id:
                # Remover anexos físicos
                for anexo in c.get("anexos", []):
                    path = anexo.get("caminho", "")
                    if os.path.exists(path):
                        try:
                            os.remove(path)
                        except:
                            pass
                contracts.pop(i)
                self._save_contracts()
                return True
        return False
    
    def get_contract_by_id(self, contract_id: str) -> Optional[dict]:
        for c in self.contracts.get("contracts", []):
            if c.get("id") == contract_id:
                return c
        return None
    
    def add_attachment(self, contract_id: str, file_path: str) -> Optional[str]:
        """Copia arquivo para pasta de anexos e registra no contrato."""
        contract = self.get_contract_by_id(contract_id)
        if not contract:
            return None
        
        # Criar subpasta do contrato
        contract_dir = os.path.join(ATTACHMENTS_DIR, contract_id)
        os.makedirs(contract_dir, exist_ok=True)
        
        # Copiar arquivo
        filename = os.path.basename(file_path)
        dest_path = os.path.join(contract_dir, filename)
        
        # Se já existe, adicionar timestamp
        if os.path.exists(dest_path):
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%H%M%S")
            filename = f"{name}_{timestamp}{ext}"
            dest_path = os.path.join(contract_dir, filename)
        
        shutil.copy2(file_path, dest_path)
        
        anexo_info = {
            "nome": filename,
            "caminho": dest_path,
            "tipo": os.path.splitext(filename)[1].lower(),
            "tamanho": os.path.getsize(dest_path),
            "data_upload": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        for i, c in enumerate(self.contracts.get("contracts", [])):
            if c.get("id") == contract_id:
                if "anexos" not in c:
                    c["anexos"] = []
                c["anexos"].append(anexo_info)
                self._save_contracts()
                return dest_path
        return None
    
    def add_medicao(self, contract_id: str, medicao_data: dict) -> bool:
        for i, c in enumerate(self.contracts.get("contracts", [])):
            if c.get("id") == contract_id:
                if "medicoes" not in c:
                    c["medicoes"] = []
                medicao_data["id"] = datetime.now().strftime("%Y%m%d%H%M%S%f")
                medicao_data["data_registro"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c["medicoes"].append(medicao_data)
                self._save_contracts()
                return True
        return False
    
    def get_dashboard_stats(self) -> dict:
        contracts = self.get_all_contracts()
        total = len(contracts)
        ativos = len([c for c in contracts if c.get("status") == "Ativo"])
        vencidos = len([c for c in contracts if c.get("status") == "Vencido"])
        encerrados = len([c for c in contracts if c.get("status") == "Encerrado"])
        em_renovacao = len([c for c in contracts if c.get("status") == "Em Renovação"])
        
        valor_total = sum(float(c.get("valor_total", 0) or 0) for c in contracts)
        valor_mensal = sum(float(c.get("valor_mensal", 0) or 0) for c in contracts if c.get("status") == "Ativo")
        
        # Contratos por departamento
        por_departamento = {}
        for c in contracts:
            dept = c.get("departamento", "Sem Departamento")
            por_departamento[dept] = por_departamento.get(dept, 0) + 1
        
        # Contratos próximos do vencimento (30 dias)
        proximos_vencimento = []
        hoje = datetime.now()
        for c in contracts:
            if c.get("status") == "Ativo" and c.get("data_fim"):
                try:
                    data_fim = datetime.strptime(c["data_fim"], "%Y-%m-%d")
                    dias_restantes = (data_fim - hoje).days
                    if 0 <= dias_restantes <= 30:
                        c_copy = c.copy()
                        c_copy["dias_restantes"] = dias_restantes
                        proximos_vencimento.append(c_copy)
                except:
                    pass
        
        # Total de medições
        total_medicoes = sum(len(c.get("medicoes", [])) for c in contracts)
        
        return {
            "total": total,
            "ativos": ativos,
            "vencidos": vencidos,
            "encerrados": encerrados,
            "em_renovacao": em_renovacao,
            "valor_total": valor_total,
            "valor_mensal": valor_mensal,
            "por_departamento": por_departamento,
            "proximos_vencimento": proximos_vencimento,
            "total_medicoes": total_medicoes,
        }


# ═══════════════════════════════════════════════════════════════════════════════
# WEB BRIDGE - Comunicação JS <-> Python
# ═══════════════════════════════════════════════════════════════════════════════

class ContractWebBridge(QObject):
    """Ponte de comunicação entre JavaScript e Python."""
    
    # Signals
    navigateToSignal = pyqtSignal(str)
    openDepartmentSignal = pyqtSignal(str)
    addContractSignal = pyqtSignal(str)
    viewContractSignal = pyqtSignal(str)
    editContractSignal = pyqtSignal(str)
    deleteContractSignal = pyqtSignal(str)
    attachFileSignal = pyqtSignal(str)
    downloadContractSignal = pyqtSignal(str)
    openPropostasSignal = pyqtSignal(str)
    addDepartmentSignal = pyqtSignal()
    removeDepartmentSignal = pyqtSignal(str)
    saveContractFormSignal = pyqtSignal(str)
    addMedicaoSignal = pyqtSignal(str)
    viewMedicaoSignal = pyqtSignal(str)
    goBackSignal = pyqtSignal()
    exportDeptReportSignal = pyqtSignal(str)
    openFileSignal = pyqtSignal(str)
    
    @pyqtSlot(str)
    def navigateTo(self, page):
        self.navigateToSignal.emit(page)
    
    @pyqtSlot(str)
    def openDepartment(self, dept):
        self.openDepartmentSignal.emit(dept)
    
    @pyqtSlot(str)
    def addContract(self, dept):
        self.addContractSignal.emit(dept)
    
    @pyqtSlot(str)
    def viewContract(self, contract_id):
        self.viewContractSignal.emit(contract_id)
    
    @pyqtSlot(str)
    def editContract(self, contract_id):
        self.editContractSignal.emit(contract_id)
    
    @pyqtSlot(str)
    def deleteContract(self, contract_id):
        self.deleteContractSignal.emit(contract_id)
    
    @pyqtSlot(str)
    def attachFile(self, contract_id):
        self.attachFileSignal.emit(contract_id)
    
    @pyqtSlot(str)
    def downloadContract(self, contract_id):
        self.downloadContractSignal.emit(contract_id)
    
    @pyqtSlot(str)
    def openPropostas(self, contract_id):
        self.openPropostasSignal.emit(contract_id)
    
    @pyqtSlot()
    def addDepartment(self):
        self.addDepartmentSignal.emit()
    
    @pyqtSlot(str)
    def removeDepartment(self, dept):
        self.removeDepartmentSignal.emit(dept)
    
    @pyqtSlot(str)
    def saveContractForm(self, json_data):
        self.saveContractFormSignal.emit(json_data)
    
    @pyqtSlot(str)
    def addMedicao(self, contract_id):
        self.addMedicaoSignal.emit(contract_id)
    
    @pyqtSlot(str)
    def viewMedicao(self, contract_id):
        self.viewMedicaoSignal.emit(contract_id)
    
    @pyqtSlot()
    def goBack(self):
        self.goBackSignal.emit()
    
    @pyqtSlot(str)
    def exportDeptReport(self, dept):
        self.exportDeptReportSignal.emit(dept)
    
    @pyqtSlot(str)
    def openFile(self, file_path):
        self.openFileSignal.emit(file_path)


# ═══════════════════════════════════════════════════════════════════════════════
# HTML GENERATOR - Geração de páginas HTML
# ═══════════════════════════════════════════════════════════════════════════════

class HTMLGenerator:
    """Gera o HTML para todas as páginas do sistema."""
    
    @staticmethod
    def get_base_html(content: str, title: str = "") -> str:
        return f"""
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>{title}</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
            <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
            <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
            <style>
                :root {{
                    --primary: {COLORS['primary']};
                    --primary-dark: {COLORS['primary_dark']};
                    --secondary: {COLORS['secondary']};
                    --success: {COLORS['success']};
                    --warning: {COLORS['warning']};
                    --danger: {COLORS['danger']};
                    --info: {COLORS['info']};
                    --dark: {COLORS['dark']};
                    --light: {COLORS['light']};
                    --radius: 16px;
                }}
                
                * {{ margin: 0; padding: 0; box-sizing: border-box; }}
                
                body {{
                    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
                    background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
                    color: var(--dark);
                    min-height: 100vh;
                    -webkit-font-smoothing: antialiased;
                }}
                
                ::-webkit-scrollbar {{ width: 8px; height: 8px; }}
                ::-webkit-scrollbar-track {{ background: transparent; }}
                ::-webkit-scrollbar-thumb {{
                    background: rgba(0,0,0,0.15);
                    border-radius: 4px;
                }}
                ::-webkit-scrollbar-thumb:hover {{ background: rgba(0,0,0,0.25); }}
                
                .dashboard-container {{
                    padding: 30px;
                    max-width: 1600px;
                    margin: 0 auto;
                }}
                
                /* Header */
                .page-header {{
                    background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
                    border-radius: var(--radius);
                    padding: 35px 40px;
                    margin-bottom: 30px;
                    color: white;
                    position: relative;
                    overflow: hidden;
                }}
                .page-header::before {{
                    content: '';
                    position: absolute;
                    top: -50%;
                    right: -20%;
                    width: 400px;
                    height: 400px;
                    background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%);
                    border-radius: 50%;
                }}
                .page-header h1 {{
                    font-size: 28px;
                    font-weight: 800;
                    margin-bottom: 8px;
                    position: relative;
                }}
                .page-header .subtitle {{
                    font-size: 14px;
                    opacity: 0.8;
                    font-weight: 400;
                    position: relative;
                }}
                
                /* KPI Cards */
                .kpi-grid {{
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
                    gap: 20px;
                    margin-bottom: 30px;
                }}
                .kpi-card {{
                    background: white;
                    border-radius: var(--radius);
                    padding: 24px;
                    box-shadow: 0 2px 15px rgba(0,0,0,0.06);
                    transition: all 0.3s ease;
                    border: 1px solid rgba(0,0,0,0.04);
                    position: relative;
                    overflow: hidden;
                }}
                .kpi-card:hover {{
                    transform: translateY(-3px);
                    box-shadow: 0 8px 25px rgba(0,0,0,0.1);
                }}
                .kpi-card .kpi-icon {{
                    width: 48px;
                    height: 48px;
                    border-radius: 12px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 20px;
                    margin-bottom: 16px;
                    color: white;
                }}
                .kpi-card .kpi-label {{
                    font-size: 13px;
                    color: #718096;
                    font-weight: 500;
                    margin-bottom: 6px;
                }}
                .kpi-card .kpi-value {{
                    font-size: 28px;
                    font-weight: 800;
                    color: var(--dark);
                }}
                .kpi-card .kpi-detail {{
                    font-size: 12px;
                    color: #a0aec0;
                    margin-top: 8px;
                }}
                
                /* Department Cards */
                .dept-grid {{
                    display: grid;
                    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
                    gap: 24px;
                    margin-bottom: 30px;
                }}
                .dept-card {{
                    background: white;
                    border-radius: var(--radius);
                    padding: 28px;
                    box-shadow: 0 2px 15px rgba(0,0,0,0.06);
                    transition: all 0.3s ease;
                    border: 1px solid rgba(0,0,0,0.04);
                    cursor: pointer;
                    position: relative;
                    overflow: hidden;
                }}
                .dept-card:hover {{
                    transform: translateY(-5px);
                    box-shadow: 0 12px 30px rgba(0,0,0,0.12);
                }}
                .dept-card .dept-icon {{
                    width: 56px;
                    height: 56px;
                    border-radius: 14px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 24px;
                    color: white;
                    margin-bottom: 18px;
                }}
                .dept-card .dept-name {{
                    font-size: 18px;
                    font-weight: 700;
                    color: var(--dark);
                    margin-bottom: 8px;
                }}
                .dept-card .dept-count {{
                    font-size: 13px;
                    color: #718096;
                }}
                .dept-card .dept-bar {{
                    height: 4px;
                    border-radius: 2px;
                    background: #edf2f7;
                    margin-top: 16px;
                    overflow: hidden;
                }}
                .dept-card .dept-bar-fill {{
                    height: 100%;
                    border-radius: 2px;
                    transition: width 0.8s ease;
                }}
                
                /* Add Department Card */
                .add-dept-card {{
                    background: white;
                    border-radius: var(--radius);
                    padding: 28px;
                    box-shadow: 0 2px 15px rgba(0,0,0,0.06);
                    transition: all 0.3s ease;
                    border: 2px dashed #cbd5e0;
                    cursor: pointer;
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    justify-content: center;
                    min-height: 180px;
                }}
                .add-dept-card:hover {{
                    border-color: var(--primary-dark);
                    background: #f7fafc;
                    transform: translateY(-3px);
                }}
                .add-dept-card i {{
                    font-size: 36px;
                    color: #a0aec0;
                    margin-bottom: 12px;
                }}
                .add-dept-card span {{
                    font-size: 14px;
                    font-weight: 600;
                    color: #718096;
                }}
                
                /* Charts */
                .chart-grid {{
                    display: grid;
                    grid-template-columns: repeat(2, 1fr);
                    gap: 24px;
                    margin-bottom: 30px;
                }}
                .chart-card {{
                    background: white;
                    border-radius: var(--radius);
                    padding: 24px;
                    box-shadow: 0 2px 15px rgba(0,0,0,0.06);
                    border: 1px solid rgba(0,0,0,0.04);
                }}
                .chart-title {{
                    font-size: 16px;
                    font-weight: 700;
                    color: var(--dark);
                    margin-bottom: 20px;
                    display: flex;
                    align-items: center;
                    gap: 10px;
                }}
                .chart-title i {{
                    color: var(--primary-dark);
                }}
                
                /* Tables */
                .table-card {{
                    background: white;
                    border-radius: var(--radius);
                    padding: 24px;
                    box-shadow: 0 2px 15px rgba(0,0,0,0.06);
                    border: 1px solid rgba(0,0,0,0.04);
                    margin-bottom: 24px;
                    overflow-x: auto;
                }}
                .table-card table {{
                    width: 100%;
                    border-collapse: separate;
                    border-spacing: 0;
                }}
                .table-card th {{
                    background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
                    color: white;
                    padding: 14px 16px;
                    text-align: left;
                    font-weight: 600;
                    font-size: 13px;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }}
                .table-card th:first-child {{ border-radius: 10px 0 0 0; }}
                .table-card th:last-child {{ border-radius: 0 10px 0 0; }}
                .table-card td {{
                    padding: 14px 16px;
                    border-bottom: 1px solid #edf2f7;
                    font-size: 13px;
                    color: #4a5568;
                }}
                .table-card tr:hover td {{
                    background: #f7fafc;
                }}
                .table-card tr:last-child td {{ border-bottom: none; }}
                
                /* Badges */
                .badge {{
                    display: inline-flex;
                    align-items: center;
                    padding: 4px 12px;
                    border-radius: 20px;
                    font-size: 11px;
                    font-weight: 600;
                    letter-spacing: 0.3px;
                }}
                .badge-success {{ background: #c6f6d5; color: #22543d; }}
                .badge-warning {{ background: #fefcbf; color: #744210; }}
                .badge-danger {{ background: #fed7d7; color: #742a2a; }}
                .badge-info {{ background: #bee3f8; color: #2a4365; }}
                .badge-secondary {{ background: #e2e8f0; color: #4a5568; }}
                
                /* Buttons */
                .btn {{
                    display: inline-flex;
                    align-items: center;
                    gap: 8px;
                    padding: 10px 20px;
                    border-radius: 10px;
                    font-size: 13px;
                    font-weight: 600;
                    border: none;
                    cursor: pointer;
                    transition: all 0.3s ease;
                    text-decoration: none;
                    font-family: 'Inter', sans-serif;
                }}
                .btn:hover {{ transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,0,0,0.15); }}
                .btn-primary {{ background: var(--primary); color: white; }}
                .btn-primary:hover {{ background: var(--primary-dark); }}
                .btn-success {{ background: var(--success); color: white; }}
                .btn-warning {{ background: var(--warning); color: white; }}
                .btn-danger {{ background: var(--danger); color: white; }}
                .btn-info {{ background: var(--info); color: white; }}
                .btn-outline {{
                    background: transparent;
                    color: var(--primary);
                    border: 2px solid var(--primary);
                }}
                .btn-outline:hover {{ background: var(--primary); color: white; }}
                .btn-sm {{ padding: 6px 14px; font-size: 12px; border-radius: 8px; }}
                .btn-icon {{
                    width: 36px;
                    height: 36px;
                    padding: 0;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    border-radius: 10px;
                    font-size: 14px;
                }}
                
                /* Back button */
                .back-btn {{
                    display: inline-flex;
                    align-items: center;
                    gap: 8px;
                    padding: 8px 16px;
                    background: rgba(255,255,255,0.15);
                    color: white;
                    border-radius: 10px;
                    border: 1px solid rgba(255,255,255,0.2);
                    cursor: pointer;
                    font-size: 13px;
                    font-weight: 500;
                    transition: all 0.2s ease;
                    margin-bottom: 16px;
                    font-family: 'Inter', sans-serif;
                }}
                .back-btn:hover {{
                    background: rgba(255,255,255,0.25);
                }}
                
                /* Form */
                .form-card {{
                    background: white;
                    border-radius: var(--radius);
                    padding: 32px;
                    box-shadow: 0 2px 15px rgba(0,0,0,0.06);
                    border: 1px solid rgba(0,0,0,0.04);
                    margin-bottom: 24px;
                }}
                .form-group {{
                    margin-bottom: 20px;
                }}
                .form-group label {{
                    display: block;
                    font-size: 13px;
                    font-weight: 600;
                    color: #4a5568;
                    margin-bottom: 8px;
                }}
                .form-group input,
                .form-group select,
                .form-group textarea {{
                    width: 100%;
                    padding: 12px 16px;
                    border: 2px solid #e2e8f0;
                    border-radius: 10px;
                    font-size: 14px;
                    font-family: 'Inter', sans-serif;
                    color: var(--dark);
                    transition: all 0.2s ease;
                    background: white;
                }}
                .form-group input:focus,
                .form-group select:focus,
                .form-group textarea:focus {{
                    outline: none;
                    border-color: var(--primary-dark);
                    box-shadow: 0 0 0 3px rgba(90, 103, 216, 0.15);
                }}
                .form-group textarea {{
                    min-height: 100px;
                    resize: vertical;
                }}
                .form-row {{
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                    gap: 20px;
                }}
                
                /* Section Title */
                .section-title {{
                    font-size: 20px;
                    font-weight: 700;
                    color: var(--dark);
                    margin-bottom: 20px;
                    display: flex;
                    align-items: center;
                    gap: 12px;
                }}
                .section-title i {{ color: var(--primary-dark); }}
                
                /* Alert boxes */
                .alert {{
                    padding: 16px 20px;
                    border-radius: 12px;
                    margin-bottom: 20px;
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    font-size: 14px;
                }}
                .alert-warning {{
                    background: #fffbeb;
                    border: 1px solid #f59e0b;
                    color: #92400e;
                }}
                .alert-info {{
                    background: #eff6ff;
                    border: 1px solid #3b82f6;
                    color: #1e40af;
                }}
                .alert-success {{
                    background: #ecfdf5;
                    border: 1px solid #10b981;
                    color: #065f46;
                }}
                
                /* Empty state */
                .empty-state {{
                    text-align: center;
                    padding: 60px 20px;
                    color: #a0aec0;
                }}
                .empty-state i {{
                    font-size: 48px;
                    margin-bottom: 16px;
                    opacity: 0.5;
                }}
                .empty-state h3 {{
                    font-size: 18px;
                    font-weight: 600;
                    margin-bottom: 8px;
                    color: #718096;
                }}
                .empty-state p {{
                    font-size: 14px;
                }}
                
                /* Attachment list */
                .attachment-list {{
                    list-style: none;
                }}
                .attachment-item {{
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    padding: 12px 16px;
                    background: #f7fafc;
                    border-radius: 10px;
                    margin-bottom: 8px;
                    transition: all 0.2s ease;
                }}
                .attachment-item:hover {{
                    background: #edf2f7;
                }}
                .attachment-item .file-icon {{
                    width: 40px;
                    height: 40px;
                    border-radius: 10px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 18px;
                }}
                .attachment-item .file-info {{
                    flex: 1;
                }}
                .attachment-item .file-name {{
                    font-size: 13px;
                    font-weight: 600;
                    color: var(--dark);
                }}
                .attachment-item .file-size {{
                    font-size: 11px;
                    color: #a0aec0;
                }}

                /* Responsive */
                @media (max-width: 900px) {{
                    .chart-grid {{ grid-template-columns: 1fr; }}
                    .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }}
                    .dept-grid {{ grid-template-columns: repeat(2, 1fr); }}
                }}
            </style>
        </head>
        <body>
            <div class="dashboard-container" id="content">
                {content}
            </div>
            <script>
                var bridge = null;
                new QWebChannel(qt.webChannelTransport, function(channel) {{
                    bridge = channel.objects.contractBridge;
                }});
                
                function navigateTo(page) {{
                    if (bridge) bridge.navigateTo(page);
                }}
                function openDepartment(dept) {{
                    if (bridge) bridge.openDepartment(dept);
                }}
                function addContract(dept) {{
                    if (bridge) bridge.addContract(dept);
                }}
                function viewContract(id) {{
                    if (bridge) bridge.viewContract(id);
                }}
                function editContract(id) {{
                    if (bridge) bridge.editContract(id);
                }}
                function deleteContract(id) {{
                    if (confirm('Tem certeza que deseja excluir este contrato?')) {{
                        if (bridge) bridge.deleteContract(id);
                    }}
                }}
                function attachFile(id) {{
                    if (bridge) bridge.attachFile(id);
                }}
                function downloadContract(id) {{
                    if (bridge) bridge.downloadContract(id);
                }}
                function openPropostas(id) {{
                    if (bridge) bridge.openPropostas(id);
                }}
                function addDepartment() {{
                    if (bridge) bridge.addDepartment();
                }}
                function removeDepartment(dept) {{
                    if (confirm('Remover a categoria "' + dept + '"? Os contratos associados não serão excluídos.')) {{
                        if (bridge) bridge.removeDepartment(dept);
                    }}
                }}
                function saveContractForm() {{
                    var form = document.getElementById('contractForm');
                    if (!form) return;
                    
                    var formData = {{}};
                    var inputs = form.querySelectorAll('input, select, textarea');
                    for (var i = 0; i < inputs.length; i++) {{
                        var input = inputs[i];
                        if (input.name) {{
                            formData[input.name] = input.value;
                        }}
                    }}
                    
                    if (bridge) bridge.saveContractForm(JSON.stringify(formData));
                }}
                function addMedicao(id) {{
                    if (bridge) bridge.addMedicao(id);
                }}
                function viewMedicao(id) {{
                    if (bridge) bridge.viewMedicao(id);
                }}
                function goBack() {{
                    if (bridge) bridge.goBack();
                }}
                function exportDeptReport(dept) {{
                    if (bridge) bridge.exportDeptReport(dept);
                }}
                function openFile(path) {{
                    if (bridge) bridge.openFile(path);
                }}
                
                function formatCurrency(value) {{
                    return 'R$ ' + parseFloat(value || 0).toLocaleString('pt-BR', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
                }}
            </script>
        </body>
        </html>
        """
    
    @staticmethod
    def fmt_currency(value) -> str:
        try:
            v = float(value or 0)
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return "R$ 0,00"
    
    @staticmethod
    def fmt_date(date_str: str) -> str:
        if not date_str:
            return "-"
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except:
            return date_str
    
    @staticmethod
    def get_status_badge(status: str) -> str:
        badges = {
            "Ativo": '<span class="badge badge-success"><i class="fas fa-check-circle"></i> Ativo</span>',
            "Vencido": '<span class="badge badge-danger"><i class="fas fa-exclamation-circle"></i> Vencido</span>',
            "Encerrado": '<span class="badge badge-secondary"><i class="fas fa-times-circle"></i> Encerrado</span>',
            "Em Renovação": '<span class="badge badge-warning"><i class="fas fa-sync-alt"></i> Em Renovação</span>',
            "Suspenso": '<span class="badge badge-info"><i class="fas fa-pause-circle"></i> Suspenso</span>',
        }
        return badges.get(status, f'<span class="badge badge-secondary">{status}</span>')
    
    @staticmethod
    def generate_dashboard(stats: dict, departments: list, contracts: list) -> str:
        """Gera o HTML do dashboard principal."""
        fmt = HTMLGenerator.fmt_currency
        
        # KPIs
        kpis_html = f"""
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #5a67d8, #4c51bf);">
                    <i class="fas fa-file-contract"></i>
                </div>
                <div class="kpi-label">Total de Contratos</div>
                <div class="kpi-value">{stats['total']}</div>
                <div class="kpi-detail">{stats['ativos']} ativos</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #48bb78, #38a169);">
                    <i class="fas fa-check-circle"></i>
                </div>
                <div class="kpi-label">Contratos Ativos</div>
                <div class="kpi-value">{stats['ativos']}</div>
                <div class="kpi-detail">{fmt(stats['valor_mensal'])}/mês</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #ed8936, #dd6b20);">
                    <i class="fas fa-exclamation-triangle"></i>
                </div>
                <div class="kpi-label">Próx. Vencimento</div>
                <div class="kpi-value">{len(stats['proximos_vencimento'])}</div>
                <div class="kpi-detail">nos próximos 30 dias</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #f56565, #e53e3e);">
                    <i class="fas fa-times-circle"></i>
                </div>
                <div class="kpi-label">Vencidos</div>
                <div class="kpi-value">{stats['vencidos']}</div>
                <div class="kpi-detail">{stats['encerrados']} encerrados</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #4299e1, #3182ce);">
                    <i class="fas fa-dollar-sign"></i>
                </div>
                <div class="kpi-label">Valor Total Contratos</div>
                <div class="kpi-value">{fmt(stats['valor_total'])}</div>
                <div class="kpi-detail">{stats['total_medicoes']} medições registradas</div>
            </div>
        </div>
        """
        
        # Departments grid
        dept_cards_html = ""
        for dept in departments:
            count = stats['por_departamento'].get(dept, 0)
            total_dept = stats['total'] if stats['total'] > 0 else 1
            pct = (count / total_dept) * 100
            icon = DEPARTMENT_ICONS.get(dept, "fa-folder")
            color = DEPARTMENT_COLORS.get(dept, "#5a67d8")
            
            dept_cards_html += f"""
            <div class="dept-card" onclick="openDepartment('{dept}')">
                <div class="dept-icon" style="background: linear-gradient(135deg, {color}, {color}dd);">
                    <i class="fas {icon}"></i>
                </div>
                <div class="dept-name">{dept}</div>
                <div class="dept-count">
                    <strong>{count}</strong> contrato{"s" if count != 1 else ""}
                </div>
                <div class="dept-bar">
                    <div class="dept-bar-fill" style="width: {pct}%; background: {color};"></div>
                </div>
            </div>
            """
        
        # Add department button
        dept_cards_html += """
        <div class="add-dept-card" onclick="addDepartment()">
            <i class="fas fa-plus-circle"></i>
            <span>Adicionar Categoria</span>
        </div>
        """
        
        # Alertas de vencimento
        alertas_html = ""
        if stats['proximos_vencimento']:
            alertas_html = '<div class="alert alert-warning"><i class="fas fa-bell"></i> <strong>Atenção!</strong> Existem contratos próximos do vencimento:</div>'
            alertas_html += '<div class="table-card"><table>'
            alertas_html += '<tr><th>Contrato</th><th>Fornecedor</th><th>Departamento</th><th>Vencimento</th><th>Dias Restantes</th><th>Ações</th></tr>'
            for c in sorted(stats['proximos_vencimento'], key=lambda x: x.get('dias_restantes', 999)):
                badge_dias = f'<span class="badge badge-danger">{c["dias_restantes"]} dias</span>' if c["dias_restantes"] <= 7 else f'<span class="badge badge-warning">{c["dias_restantes"]} dias</span>'
                alertas_html += f"""
                <tr>
                    <td><strong>{c.get('numero_contrato', '-')}</strong></td>
                    <td>{c.get('fornecedor', '-')}</td>
                    <td>{c.get('departamento', '-')}</td>
                    <td>{HTMLGenerator.fmt_date(c.get('data_fim', ''))}</td>
                    <td>{badge_dias}</td>
                    <td>
                        <button class="btn btn-sm btn-primary" onclick="viewContract('{c.get('id', '')}')">
                            <i class="fas fa-eye"></i> Ver
                        </button>
                    </td>
                </tr>
                """
            alertas_html += '</table></div>'
        
        # Charts
        dept_labels = json.dumps([d for d in departments if stats['por_departamento'].get(d, 0) > 0])
        dept_values = json.dumps([stats['por_departamento'].get(d, 0) for d in departments if stats['por_departamento'].get(d, 0) > 0])
        dept_colors_list = json.dumps([DEPARTMENT_COLORS.get(d, "#5a67d8") for d in departments if stats['por_departamento'].get(d, 0) > 0])
        
        # Status chart data
        status_data = {
            "Ativo": stats['ativos'],
            "Vencido": stats['vencidos'],
            "Encerrado": stats['encerrados'],
            "Em Renovação": stats['em_renovacao']
        }
        status_labels = json.dumps([k for k, v in status_data.items() if v > 0])
        status_values = json.dumps([v for v in status_data.values() if v > 0])
        status_colors = json.dumps(['#48bb78', '#f56565', '#a0aec0', '#ed8936'][:len([v for v in status_data.values() if v > 0])])
        
        charts_html = f"""
        <div class="chart-grid">
            <div class="chart-card">
                <div class="chart-title"><i class="fas fa-chart-bar"></i> Contratos por Departamento</div>
                <canvas id="chartDept" height="300"></canvas>
            </div>
            <div class="chart-card">
                <div class="chart-title"><i class="fas fa-chart-pie"></i> Status dos Contratos</div>
                <canvas id="chartStatus" height="300"></canvas>
            </div>
        </div>
        <script>
            document.addEventListener('DOMContentLoaded', function() {{
                // Dept chart
                var deptLabels = {dept_labels};
                var deptValues = {dept_values};
                var deptColors = {dept_colors_list};
                
                if (deptLabels.length > 0) {{
                    new Chart(document.getElementById('chartDept'), {{
                        type: 'bar',
                        data: {{
                            labels: deptLabels,
                            datasets: [{{
                                label: 'Contratos',
                                data: deptValues,
                                backgroundColor: deptColors.map(c => c + '99'),
                                borderColor: deptColors,
                                borderWidth: 2,
                                borderRadius: 8
                            }}]
                        }},
                        options: {{
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {{ legend: {{ display: false }} }},
                            scales: {{
                                y: {{ beginAtZero: true, ticks: {{ stepSize: 1 }} }},
                                x: {{ ticks: {{ maxRotation: 45, font: {{ size: 11 }} }} }}
                            }}
                        }}
                    }});
                }}
                
                // Status chart
                var statusLabels = {status_labels};
                var statusValues = {status_values};
                var statusColors = {status_colors};
                
                if (statusLabels.length > 0) {{
                    new Chart(document.getElementById('chartStatus'), {{
                        type: 'doughnut',
                        data: {{
                            labels: statusLabels,
                            datasets: [{{
                                data: statusValues,
                                backgroundColor: statusColors,
                                borderWidth: 3,
                                borderColor: '#ffffff'
                            }}]
                        }},
                        options: {{
                            responsive: true,
                            maintainAspectRatio: false,
                            cutout: '60%',
                            plugins: {{
                                legend: {{
                                    position: 'bottom',
                                    labels: {{ padding: 20, usePointStyle: true, font: {{ size: 12 }} }}
                                }}
                            }}
                        }}
                    }});
                }}
            }});
        </script>
        """
        
        content = f"""
        <div class="page-header">
            <h1><i class="fas fa-file-contract"></i> Gestão de Contratos e Medição</h1>
            <p class="subtitle">Painel de controle e monitoramento de contratos por departamento</p>
        </div>
        
        {kpis_html}
        {alertas_html}
        
        <div class="section-title">
            <i class="fas fa-th-large"></i> Departamentos
        </div>
        <div class="dept-grid">
            {dept_cards_html}
        </div>
        
        {charts_html}
        """
        
        return HTMLGenerator.get_base_html(content, "Dashboard - Gestão de Contratos")
    
    @staticmethod
    def generate_department_page(department: str, contracts: list) -> str:
        """Gera a página de um departamento específico."""
        fmt = HTMLGenerator.fmt_currency
        icon = DEPARTMENT_ICONS.get(department, "fa-folder")
        color = DEPARTMENT_COLORS.get(department, "#5a67d8")
        
        total = len(contracts)
        ativos = len([c for c in contracts if c.get("status") == "Ativo"])
        valor_total = sum(float(c.get("valor_total", 0) or 0) for c in contracts)
        total_medicoes = sum(len(c.get("medicoes", [])) for c in contracts)
        
        # KPIs do departamento
        kpis_html = f"""
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, {color}, {color}dd);">
                    <i class="fas fa-file-contract"></i>
                </div>
                <div class="kpi-label">Total de Contratos</div>
                <div class="kpi-value">{total}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #48bb78, #38a169);">
                    <i class="fas fa-check-circle"></i>
                </div>
                <div class="kpi-label">Ativos</div>
                <div class="kpi-value">{ativos}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #4299e1, #3182ce);">
                    <i class="fas fa-dollar-sign"></i>
                </div>
                <div class="kpi-label">Valor Total</div>
                <div class="kpi-value">{fmt(valor_total)}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon" style="background: linear-gradient(135deg, #9f7aea, #805ad5);">
                    <i class="fas fa-ruler"></i>
                </div>
                <div class="kpi-label">Medições</div>
                <div class="kpi-value">{total_medicoes}</div>
            </div>
        </div>
        """
        
        # Tabela de contratos
        if contracts:
            rows = ""
            for c in sorted(contracts, key=lambda x: x.get("criado_em", ""), reverse=True):
                status_badge = HTMLGenerator.get_status_badge(c.get("status", ""))
                anexos_count = len(c.get("anexos", []))
                medicoes_count = len(c.get("medicoes", []))
                
                rows += f"""
                <tr>
                    <td><strong>{c.get('numero_contrato', '-')}</strong></td>
                    <td>{c.get('fornecedor', '-')}</td>
                    <td>{c.get('objeto', '-')[:50]}{'...' if len(c.get('objeto', '')) > 50 else ''}</td>
                    <td>{fmt(c.get('valor_total', 0))}</td>
                    <td>{HTMLGenerator.fmt_date(c.get('data_inicio', ''))}</td>
                    <td>{HTMLGenerator.fmt_date(c.get('data_fim', ''))}</td>
                    <td>{status_badge}</td>
                    <td>
                        <span class="badge badge-info" title="Anexos"><i class="fas fa-paperclip"></i> {anexos_count}</span>
                        <span class="badge badge-secondary" title="Medições"><i class="fas fa-ruler"></i> {medicoes_count}</span>
                    </td>
                    <td>
                        <div style="display: flex; gap: 6px;">
                            <button class="btn btn-icon btn-primary" onclick="viewContract('{c.get('id', '')}')" title="Ver Detalhes">
                                <i class="fas fa-eye"></i>
                            </button>
                            <button class="btn btn-icon btn-info" onclick="attachFile('{c.get('id', '')}')" title="Anexar Documento">
                                <i class="fas fa-paperclip"></i>
                            </button>
                            <button class="btn btn-icon btn-success" onclick="downloadContract('{c.get('id', '')}')" title="Baixar Contrato">
                                <i class="fas fa-download"></i>
                            </button>
                            <button class="btn btn-icon btn-warning" onclick="editContract('{c.get('id', '')}')" title="Editar">
                                <i class="fas fa-edit"></i>
                            </button>
                            <button class="btn btn-icon btn-danger" onclick="deleteContract('{c.get('id', '')}')" title="Excluir">
                                <i class="fas fa-trash"></i>
                            </button>
                        </div>
                    </td>
                </tr>
                """
            
            table_html = f"""
            <div class="table-card">
                <table>
                    <tr>
                        <th>Nº Contrato</th>
                        <th>Fornecedor</th>
                        <th>Objeto</th>
                        <th>Valor Total</th>
                        <th>Início</th>
                        <th>Término</th>
                        <th>Status</th>
                        <th>Docs</th>
                        <th>Ações</th>
                    </tr>
                    {rows}
                </table>
            </div>
            """
        else:
            table_html = """
            <div class="form-card">
                <div class="empty-state">
                    <i class="fas fa-file-contract"></i>
                    <h3>Nenhum contrato cadastrado</h3>
                    <p>Clique em "Novo Contrato" para adicionar o primeiro contrato deste departamento.</p>
                </div>
            </div>
            """
        
        is_custom = department not in DEFAULT_DEPARTMENTS
        remove_btn = f"""
        <button class="btn btn-sm btn-danger" onclick="removeDepartment('{department}')" style="margin-left: 12px;">
            <i class="fas fa-trash"></i> Remover Categoria
        </button>
        """ if is_custom else ""
        
        content = f"""
        <div class="page-header" style="background: linear-gradient(135deg, {color} 0%, {color}cc 100%);">
            <button class="back-btn" onclick="navigateTo('dashboard')">
                <i class="fas fa-arrow-left"></i> Voltar ao Dashboard
            </button>
            <h1><i class="fas {icon}"></i> {department}</h1>
            <p class="subtitle">Gestão de contratos do departamento</p>
        </div>
        
        {kpis_html}
        
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
            <div class="section-title" style="margin-bottom: 0;">
                <i class="fas fa-list"></i> Contratos
            </div>
            <div style="display: flex; gap: 8px;">
                <button class="btn btn-primary" onclick="addContract('{department}')">
                    <i class="fas fa-plus"></i> Novo Contrato
                </button>
                {remove_btn}
            </div>
        </div>
        
        {table_html}
        """
        
        return HTMLGenerator.get_base_html(content, f"{department} - Gestão de Contratos")
    
    @staticmethod
    def generate_contract_form(department: str, contract: dict = None) -> str:
        """Gera o formulário de cadastro/edição de contrato."""
        is_edit = contract is not None
        title = "Editar Contrato" if is_edit else "Novo Contrato"
        icon = DEPARTMENT_ICONS.get(department, "fa-folder")
        color = DEPARTMENT_COLORS.get(department, "#5a67d8")
        
        c = contract or {}
        
        # Status options
        status_options = ""
        for status in ["Ativo", "Vencido", "Encerrado", "Em Renovação", "Suspenso"]:
            selected = "selected" if c.get("status") == status else ""
            status_options += f'<option value="{status}" {selected}>{status}</option>'
        
        # Tipo options
        tipo_options = ""
        for tipo in ["Prestação de Serviço", "Fornecimento de Material", "Locação", "Empreitada", "Consultoria", "Manutenção", "Outro"]:
            selected = "selected" if c.get("tipo_contrato") == tipo else ""
            tipo_options += f'<option value="{tipo}" {selected}>{tipo}</option>'
        
        # Forma pagamento options
        pgto_options = ""
        for pgto in ["Mensal", "Quinzenal", "Semanal", "Por Medição", "À Vista", "Parcelado", "Outro"]:
            selected = "selected" if c.get("forma_pagamento") == pgto else ""
            pgto_options += f'<option value="{pgto}" {selected}>{pgto}</option>'
        
        # Índice reajuste options
        reajuste_options = ""
        for idx in ["IPCA", "IGPM", "INPC", "INCC", "Sem Reajuste", "Outro"]:
            selected = "selected" if c.get("indice_reajuste") == idx else ""
            reajuste_options += f'<option value="{idx}" {selected}>{idx}</option>'
        
        contract_id = c.get("id", "")
        
        content = f"""
        <div class="page-header" style="background: linear-gradient(135deg, {color} 0%, {color}cc 100%);">
            <button class="back-btn" onclick="openDepartment('{department}')">
                <i class="fas fa-arrow-left"></i> Voltar para {department}
            </button>
            <h1><i class="fas {icon}"></i> {title}</h1>
            <p class="subtitle">Departamento: {department}</p>
        </div>
        
        <form id="contractForm">
            <input type="hidden" name="id" value="{contract_id}">
            <input type="hidden" name="departamento" value="{department}">
            
            <!-- Dados Gerais -->
            <div class="form-card">
                <div class="section-title"><i class="fas fa-info-circle"></i> Dados Gerais do Contrato</div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Número do Contrato *</label>
                        <input type="text" name="numero_contrato" value="{c.get('numero_contrato', '')}" placeholder="Ex: CT-2026/001" required>
                    </div>
                    <div class="form-group">
                        <label>Tipo de Contrato *</label>
                        <select name="tipo_contrato">{tipo_options}</select>
                    </div>
                    <div class="form-group">
                        <label>Status *</label>
                        <select name="status">{status_options}</select>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Objeto do Contrato *</label>
                    <textarea name="objeto" placeholder="Descreva o objeto do contrato...">{c.get('objeto', '')}</textarea>
                </div>
            </div>
            
            <!-- Fornecedor -->
            <div class="form-card">
                <div class="section-title"><i class="fas fa-building"></i> Dados do Fornecedor</div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Razão Social / Fornecedor *</label>
                        <input type="text" name="fornecedor" value="{c.get('fornecedor', '')}" placeholder="Nome da empresa">
                    </div>
                    <div class="form-group">
                        <label>CNPJ</label>
                        <input type="text" name="cnpj" value="{c.get('cnpj', '')}" placeholder="00.000.000/0000-00">
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Contato do Fornecedor</label>
                        <input type="text" name="contato_fornecedor" value="{c.get('contato_fornecedor', '')}" placeholder="Nome do contato">
                    </div>
                    <div class="form-group">
                        <label>Telefone</label>
                        <input type="text" name="telefone_fornecedor" value="{c.get('telefone_fornecedor', '')}" placeholder="(00) 00000-0000">
                    </div>
                    <div class="form-group">
                        <label>E-mail</label>
                        <input type="email" name="email_fornecedor" value="{c.get('email_fornecedor', '')}" placeholder="email@empresa.com">
                    </div>
                </div>
            </div>
            
            <!-- Valores e Vigência -->
            <div class="form-card">
                <div class="section-title"><i class="fas fa-calendar-alt"></i> Vigência e Valores</div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Data de Início *</label>
                        <input type="date" name="data_inicio" value="{c.get('data_inicio', '')}">
                    </div>
                    <div class="form-group">
                        <label>Data de Término *</label>
                        <input type="date" name="data_fim" value="{c.get('data_fim', '')}">
                    </div>
                    <div class="form-group">
                        <label>Prazo (meses)</label>
                        <input type="number" name="prazo_meses" value="{c.get('prazo_meses', '')}" placeholder="12">
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Valor Total do Contrato (R$)</label>
                        <input type="number" name="valor_total" value="{c.get('valor_total', '')}" step="0.01" placeholder="0,00">
                    </div>
                    <div class="form-group">
                        <label>Valor Mensal (R$)</label>
                        <input type="number" name="valor_mensal" value="{c.get('valor_mensal', '')}" step="0.01" placeholder="0,00">
                    </div>
                    <div class="form-group">
                        <label>Forma de Pagamento</label>
                        <select name="forma_pagamento">{pgto_options}</select>
                    </div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Índice de Reajuste</label>
                        <select name="indice_reajuste">{reajuste_options}</select>
                    </div>
                    <div class="form-group">
                        <label>Data do Próximo Reajuste</label>
                        <input type="date" name="data_reajuste" value="{c.get('data_reajuste', '')}">
                    </div>
                    <div class="form-group">
                        <label>Garantia Contratual</label>
                        <input type="text" name="garantia" value="{c.get('garantia', '')}" placeholder="Tipo e valor da garantia">
                    </div>
                </div>
            </div>
            
            <!-- Fiscal e Responsável -->
            <div class="form-card">
                <div class="section-title"><i class="fas fa-user-tie"></i> Gestão e Fiscalização</div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Gestor do Contrato</label>
                        <input type="text" name="gestor" value="{c.get('gestor', '')}" placeholder="Nome do gestor responsável">
                    </div>
                    <div class="form-group">
                        <label>Fiscal do Contrato</label>
                        <input type="text" name="fiscal" value="{c.get('fiscal', '')}" placeholder="Nome do fiscal">
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Observações</label>
                    <textarea name="observacoes" placeholder="Observações adicionais sobre o contrato...">{c.get('observacoes', '')}</textarea>
                </div>
            </div>
            
            <!-- Link de Propostas -->
            <div class="form-card">
                <div class="section-title"><i class="fas fa-link"></i> Links e Referências</div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label>Link das Propostas</label>
                        <input type="url" name="link_propostas" value="{c.get('link_propostas', '')}" placeholder="https://...">
                    </div>
                    <div class="form-group">
                        <label>Link do Contrato Digital</label>
                        <input type="url" name="link_contrato" value="{c.get('link_contrato', '')}" placeholder="https://...">
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Caminho do Contrato Físico (Rede)</label>
                    <input type="text" name="caminho_contrato_fisico" value="{c.get('caminho_contrato_fisico', '')}" placeholder="G:\\Contratos\\...">
                </div>
            </div>
            
            <!-- Botões -->
            <div style="display: flex; gap: 12px; justify-content: flex-end; margin-top: 10px;">
                <button type="button" class="btn btn-outline" onclick="openDepartment('{department}')">
                    <i class="fas fa-times"></i> Cancelar
                </button>
                <button type="button" class="btn btn-success" onclick="saveContractForm()">
                    <i class="fas fa-save"></i> {'Salvar Alterações' if is_edit else 'Cadastrar Contrato'}
                </button>
            </div>
        </form>
        """
        
        return HTMLGenerator.get_base_html(content, f"{title} - {department}")
    
    @staticmethod
    def generate_contract_detail(contract: dict, department: str) -> str:
        """Gera a página de detalhes de um contrato."""
        fmt = HTMLGenerator.fmt_currency
        c = contract
        icon = DEPARTMENT_ICONS.get(department, "fa-folder")
        color = DEPARTMENT_COLORS.get(department, "#5a67d8")
        
        # Info cards
        info_html = f"""
        <div class="form-card">
            <div class="section-title"><i class="fas fa-info-circle"></i> Informações Gerais</div>
            <div class="form-row">
                <div>
                    <div class="form-group">
                        <label>Número do Contrato</label>
                        <div style="font-size: 16px; font-weight: 700; color: var(--dark); padding: 8px 0;">{c.get('numero_contrato', '-')}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Tipo</label>
                        <div style="padding: 8px 0;">{c.get('tipo_contrato', '-')}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Status</label>
                        <div style="padding: 8px 0;">{HTMLGenerator.get_status_badge(c.get('status', ''))}</div>
                    </div>
                </div>
            </div>
            <div class="form-group">
                <label>Objeto</label>
                <div style="padding: 8px 0; line-height: 1.6;">{c.get('objeto', '-')}</div>
            </div>
        </div>
        
        <div class="form-card">
            <div class="section-title"><i class="fas fa-building"></i> Fornecedor</div>
            <div class="form-row">
                <div>
                    <div class="form-group">
                        <label>Razão Social</label>
                        <div style="padding: 8px 0; font-weight: 600;">{c.get('fornecedor', '-')}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>CNPJ</label>
                        <div style="padding: 8px 0;">{c.get('cnpj', '-')}</div>
                    </div>
                </div>
            </div>
            <div class="form-row">
                <div>
                    <div class="form-group">
                        <label>Contato</label>
                        <div style="padding: 8px 0;">{c.get('contato_fornecedor', '-')}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Telefone</label>
                        <div style="padding: 8px 0;">{c.get('telefone_fornecedor', '-')}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>E-mail</label>
                        <div style="padding: 8px 0;">{c.get('email_fornecedor', '-')}</div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="form-card">
            <div class="section-title"><i class="fas fa-calendar-alt"></i> Vigência e Valores</div>
            <div class="form-row">
                <div>
                    <div class="form-group">
                        <label>Data de Início</label>
                        <div style="padding: 8px 0;">{HTMLGenerator.fmt_date(c.get('data_inicio', ''))}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Data de Término</label>
                        <div style="padding: 8px 0;">{HTMLGenerator.fmt_date(c.get('data_fim', ''))}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Prazo</label>
                        <div style="padding: 8px 0;">{c.get('prazo_meses', '-')} meses</div>
                    </div>
                </div>
            </div>
            <div class="form-row">
                <div>
                    <div class="form-group">
                        <label>Valor Total</label>
                        <div style="padding: 8px 0; font-size: 18px; font-weight: 700; color: var(--primary-dark);">{fmt(c.get('valor_total', 0))}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Valor Mensal</label>
                        <div style="padding: 8px 0; font-size: 16px; font-weight: 600;">{fmt(c.get('valor_mensal', 0))}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Forma de Pagamento</label>
                        <div style="padding: 8px 0;">{c.get('forma_pagamento', '-')}</div>
                    </div>
                </div>
            </div>
            <div class="form-row">
                <div>
                    <div class="form-group">
                        <label>Índice de Reajuste</label>
                        <div style="padding: 8px 0;">{c.get('indice_reajuste', '-')}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Próximo Reajuste</label>
                        <div style="padding: 8px 0;">{HTMLGenerator.fmt_date(c.get('data_reajuste', ''))}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Garantia</label>
                        <div style="padding: 8px 0;">{c.get('garantia', '-')}</div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="form-card">
            <div class="section-title"><i class="fas fa-user-tie"></i> Gestão e Fiscalização</div>
            <div class="form-row">
                <div>
                    <div class="form-group">
                        <label>Gestor</label>
                        <div style="padding: 8px 0;">{c.get('gestor', '-')}</div>
                    </div>
                </div>
                <div>
                    <div class="form-group">
                        <label>Fiscal</label>
                        <div style="padding: 8px 0;">{c.get('fiscal', '-')}</div>
                    </div>
                </div>
            </div>
            <div class="form-group">
                <label>Observações</label>
                <div style="padding: 8px 0; line-height: 1.6;">{c.get('observacoes', '-')}</div>
            </div>
        </div>
        """
        
        # Botões de ação
        propostas_btn = ""
        if c.get('link_propostas'):
            propostas_btn = f"""
            <button class="btn btn-info" onclick="openPropostas('{c.get('id', '')}')">
                <i class="fas fa-external-link-alt"></i> Acessar Propostas
            </button>"""
        
        contrato_fisico_btn = ""
        if c.get('caminho_contrato_fisico'):
            contrato_fisico_btn = f"""
            <button class="btn btn-success" onclick="downloadContract('{c.get('id', '')}')">
                <i class="fas fa-download"></i> Baixar Contrato Físico
            </button>"""
        
        actions_html = f"""
        <div class="form-card">
            <div class="section-title"><i class="fas fa-cogs"></i> Ações</div>
            <div style="display: flex; gap: 12px; flex-wrap: wrap;">
                <button class="btn btn-primary" onclick="editContract('{c.get('id', '')}')">
                    <i class="fas fa-edit"></i> Editar Contrato
                </button>
                <button class="btn btn-warning" onclick="attachFile('{c.get('id', '')}')">
                    <i class="fas fa-paperclip"></i> Anexar Documento
                </button>
                {contrato_fisico_btn}
                {propostas_btn}
                <button class="btn btn-info" onclick="addMedicao('{c.get('id', '')}')">
                    <i class="fas fa-ruler"></i> Registrar Medição
                </button>
            </div>
        </div>
        """
        
        # Anexos
        anexos = c.get("anexos", [])
        if anexos:
            anexos_items = ""
            for a in anexos:
                ext = a.get("tipo", "").lower()
                file_icon_class = "fa-file-pdf" if ext == ".pdf" else "fa-file-excel" if ext in [".xls", ".xlsx"] else "fa-file-word" if ext in [".doc", ".docx"] else "fa-file-image" if ext in [".jpg", ".jpeg", ".png", ".gif"] else "fa-file"
                file_color = "#e53e3e" if ext == ".pdf" else "#38a169" if ext in [".xls", ".xlsx"] else "#4299e1" if ext in [".doc", ".docx"] else "#ed8936" if ext in [".jpg", ".jpeg", ".png", ".gif"] else "#718096"
                tamanho = a.get("tamanho", 0)
                tamanho_fmt = f"{tamanho/1024:.1f} KB" if tamanho < 1048576 else f"{tamanho/1048576:.1f} MB"
                
                caminho_escaped = a.get('caminho', '').replace('\\', '/')
                anexos_items += f"""
                <div class="attachment-item">
                    <div class="file-icon" style="background: {file_color}22; color: {file_color};">
                        <i class="fas {file_icon_class}"></i>
                    </div>
                    <div class="file-info">
                        <div class="file-name">{a.get('nome', '-')}</div>
                        <div class="file-size">{tamanho_fmt} - {a.get('data_upload', '-')}</div>
                    </div>
                    <button class="btn btn-sm btn-outline" onclick="openFile('{caminho_escaped}')">
                        <i class="fas fa-external-link-alt"></i> Abrir
                    </button>
                </div>
                """
            
            anexos_html = f"""
            <div class="form-card">
                <div class="section-title"><i class="fas fa-paperclip"></i> Documentos Anexados ({len(anexos)})</div>
                <div class="attachment-list">
                    {anexos_items}
                </div>
            </div>
            """
        else:
            anexos_html = """
            <div class="form-card">
                <div class="section-title"><i class="fas fa-paperclip"></i> Documentos Anexados</div>
                <div class="empty-state" style="padding: 30px;">
                    <i class="fas fa-file-upload"></i>
                    <h3>Nenhum documento anexado</h3>
                    <p>Use o botão "Anexar Documento" para adicionar arquivos ao contrato.</p>
                </div>
            </div>
            """
        
        # Medições
        medicoes = c.get("medicoes", [])
        if medicoes:
            med_rows = ""
            for m in sorted(medicoes, key=lambda x: x.get("data_registro", ""), reverse=True):
                med_rows += f"""
                <tr>
                    <td>{m.get('numero', '-')}</td>
                    <td>{m.get('periodo_referencia', '-')}</td>
                    <td>{HTMLGenerator.fmt_date(m.get('data_medicao', ''))}</td>
                    <td>{fmt(m.get('valor', 0))}</td>
                    <td>{HTMLGenerator.get_status_badge(m.get('status', 'Pendente'))}</td>
                    <td>{m.get('observacao', '-')}</td>
                </tr>
                """
            
            medicoes_html = f"""
            <div class="form-card">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div class="section-title" style="margin-bottom: 0;"><i class="fas fa-ruler"></i> Medições ({len(medicoes)})</div>
                    <button class="btn btn-sm btn-info" onclick="addMedicao('{c.get('id', '')}')">
                        <i class="fas fa-plus"></i> Nova Medição
                    </button>
                </div>
                <div style="margin-top: 16px;">
                    <table style="width: 100%; border-collapse: separate; border-spacing: 0;">
                        <tr>
                            <th style="background: linear-gradient(135deg, var(--primary), var(--secondary)); color: white; padding: 12px 16px; text-align: left; font-size: 13px; font-weight: 600;">Nº</th>
                            <th style="background: linear-gradient(135deg, var(--primary), var(--secondary)); color: white; padding: 12px 16px; text-align: left; font-size: 13px; font-weight: 600;">Período</th>
                            <th style="background: linear-gradient(135deg, var(--primary), var(--secondary)); color: white; padding: 12px 16px; text-align: left; font-size: 13px; font-weight: 600;">Data</th>
                            <th style="background: linear-gradient(135deg, var(--primary), var(--secondary)); color: white; padding: 12px 16px; text-align: left; font-size: 13px; font-weight: 600;">Valor</th>
                            <th style="background: linear-gradient(135deg, var(--primary), var(--secondary)); color: white; padding: 12px 16px; text-align: left; font-size: 13px; font-weight: 600;">Status</th>
                            <th style="background: linear-gradient(135deg, var(--primary), var(--secondary)); color: white; padding: 12px 16px; text-align: left; font-size: 13px; font-weight: 600;">Observação</th>
                        </tr>
                        {med_rows}
                    </table>
                </div>
            </div>
            """
        else:
            medicoes_html = """
            <div class="form-card">
                <div class="section-title"><i class="fas fa-ruler"></i> Medições</div>
                <div class="empty-state" style="padding: 30px;">
                    <i class="fas fa-ruler-combined"></i>
                    <h3>Nenhuma medição registrada</h3>
                    <p>Use o botão "Registrar Medição" para adicionar medições ao contrato.</p>
                </div>
            </div>
            """
        
        content = f"""
        <div class="page-header" style="background: linear-gradient(135deg, {color} 0%, {color}cc 100%);">
            <button class="back-btn" onclick="openDepartment('{department}')">
                <i class="fas fa-arrow-left"></i> Voltar para {department}
            </button>
            <h1><i class="fas fa-file-contract"></i> {c.get('numero_contrato', 'Contrato')} - {c.get('fornecedor', '')}</h1>
            <p class="subtitle">{HTMLGenerator.get_status_badge(c.get('status', ''))} | Vigência: {HTMLGenerator.fmt_date(c.get('data_inicio', ''))} a {HTMLGenerator.fmt_date(c.get('data_fim', ''))}</p>
        </div>
        
        {actions_html}
        {info_html}
        {anexos_html}
        {medicoes_html}
        """
        
        return HTMLGenerator.get_base_html(content, f"Contrato {c.get('numero_contrato', '')} - Detalhes")


# ═══════════════════════════════════════════════════════════════════════════════
# JANELA PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

class ContractMainWindow(QMainWindow):
    """Janela principal da aplicação de Gestão de Contratos."""
    
    def __init__(self):
        super().__init__()
        self.data_manager = ContractDataManager()
        self.current_view = "dashboard"
        self.current_department = None
        self.current_contract_id = None
        self.setup_ui()
        self.load_dashboard()
    
    def setup_ui(self):
        """Configura a interface principal."""
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(1400, 900)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # ── Sidebar ──
        sidebar = QWidget()
        sidebar.setStyleSheet("""
            QWidget { background: qlineargradient(y1:0, y2:1, stop:0 #1a1a2e, stop:1 #16213e); min-width: 260px; max-width: 260px; }
            QPushButton { background: transparent; color: rgba(255,255,255,0.7); border: none; border-radius: 8px; padding: 14px 20px; text-align: left; font-size: 13px; font-weight: 500; }
            QPushButton:hover { background: rgba(255,255,255,0.1); color: white; }
            QPushButton:checked { background: qlineargradient(x1:0, x2:1, stop:0 #2d3748, stop:1 #1a202c); color: white; font-weight: bold; }
            QLabel { color: white; padding: 20px; }
        """)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        
        # Título
        title = QLabel("📋 Gestão de Contratos")
        title.setStyleSheet("QLabel { font-size: 17px; font-weight: bold; padding: 25px 20px 5px; }")
        sidebar_layout.addWidget(title)
        
        subtitle = QLabel("v1.0 - Contratos e Medição")
        subtitle.setStyleSheet("QLabel { font-size: 10px; color: rgba(255,255,255,0.5); padding: 0 20px 15px; }")
        sidebar_layout.addWidget(subtitle)
        
        # Separador
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("QFrame { background: rgba(255,255,255,0.1); max-height: 1px; margin: 0 15px; }")
        sidebar_layout.addWidget(sep)
        
        # Botões de navegação
        self.nav_buttons = {}
        
        # Dashboard
        btn_dashboard = QPushButton("📊  Dashboard")
        btn_dashboard.setCheckable(True)
        btn_dashboard.clicked.connect(self.load_dashboard)
        self.nav_buttons["dashboard"] = btn_dashboard
        sidebar_layout.addWidget(btn_dashboard)
        
        # Separador de departamentos
        dept_label = QLabel("DEPARTAMENTOS")
        dept_label.setStyleSheet("QLabel { font-size: 10px; color: rgba(255,255,255,0.4); padding: 15px 20px 5px; font-weight: bold; letter-spacing: 1px; }")
        sidebar_layout.addWidget(dept_label)
        
        # Scroll area para departamentos
        dept_scroll = QScrollArea()
        dept_scroll.setWidgetResizable(True)
        dept_scroll.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QWidget { background: transparent; }
            QScrollBar:vertical { width: 4px; background: transparent; }
            QScrollBar::handle:vertical { background: rgba(255,255,255,0.2); border-radius: 2px; }
        """)
        
        self.dept_widget = QWidget()
        self.dept_layout = QVBoxLayout(self.dept_widget)
        self.dept_layout.setContentsMargins(0, 0, 0, 0)
        self.dept_layout.setSpacing(2)
        
        dept_scroll.setWidget(self.dept_widget)
        sidebar_layout.addWidget(dept_scroll, 1)
        
        # Botão adicionar categoria
        btn_add_dept = QPushButton("➕  Adicionar Categoria")
        btn_add_dept.setStyleSheet("""
            QPushButton { 
                color: rgba(255,255,255,0.5); 
                border: 1px dashed rgba(255,255,255,0.2); 
                margin: 5px 10px;
                font-size: 12px;
            }
            QPushButton:hover { 
                color: white; 
                border-color: rgba(255,255,255,0.4);
                background: rgba(255,255,255,0.05);
            }
        """)
        btn_add_dept.clicked.connect(self.handle_add_department)
        sidebar_layout.addWidget(btn_add_dept)
        
        # Separador
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        sep2.setStyleSheet("QFrame { background: rgba(255,255,255,0.1); max-height: 1px; margin: 5px 15px; }")
        sidebar_layout.addWidget(sep2)
        
        # Status
        self.status_label = QLabel("🟢 Sistema Ativo")
        self.status_label.setStyleSheet("QLabel { color: #48bb78; padding: 10px 20px; font-size: 11px; background: rgba(72, 187, 120, 0.1); border-radius: 6px; margin: 5px 10px; }")
        sidebar_layout.addWidget(self.status_label)
        
        version = QLabel(f"v{APP_VERSION}")
        version.setStyleSheet("QLabel { color: rgba(255,255,255,0.3); padding: 10px 20px; font-size: 10px; }")
        sidebar_layout.addWidget(version)
        
        main_layout.addWidget(sidebar)
        
        # ── WebView ──
        self.web_view = QWebEngineView()
        self.web_channel = QWebChannel()
        self.bridge = ContractWebBridge()
        self.web_channel.registerObject('contractBridge', self.bridge)
        self.web_view.page().setWebChannel(self.web_channel)
        
        # Conectar signals
        self.bridge.navigateToSignal.connect(self.handle_navigate)
        self.bridge.openDepartmentSignal.connect(self.handle_open_department)
        self.bridge.addContractSignal.connect(self.handle_add_contract)
        self.bridge.viewContractSignal.connect(self.handle_view_contract)
        self.bridge.editContractSignal.connect(self.handle_edit_contract)
        self.bridge.deleteContractSignal.connect(self.handle_delete_contract)
        self.bridge.attachFileSignal.connect(self.handle_attach_file)
        self.bridge.downloadContractSignal.connect(self.handle_download_contract)
        self.bridge.openPropostasSignal.connect(self.handle_open_propostas)
        self.bridge.addDepartmentSignal.connect(self.handle_add_department)
        self.bridge.removeDepartmentSignal.connect(self.handle_remove_department)
        self.bridge.saveContractFormSignal.connect(self.handle_save_contract)
        self.bridge.addMedicaoSignal.connect(self.handle_add_medicao)
        self.bridge.viewMedicaoSignal.connect(self.handle_view_medicao)
        self.bridge.goBackSignal.connect(self.load_dashboard)
        self.bridge.openFileSignal.connect(self.handle_open_file)
        
        main_layout.addWidget(self.web_view, 1)
        
        # Carregar departamentos na sidebar
        self._refresh_dept_buttons()
    
    def _refresh_dept_buttons(self):
        """Atualiza os botões de departamento na sidebar."""
        # Limpar botões existentes
        while self.dept_layout.count():
            item = self.dept_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Remover botões antigos do nav_buttons
        keys_to_remove = [k for k in self.nav_buttons if k.startswith("dept_")]
        for k in keys_to_remove:
            del self.nav_buttons[k]
        
        departments = self.data_manager.get_all_departments()
        contracts = self.data_manager.get_all_contracts()
        
        for dept in departments:
            count = len([c for c in contracts if c.get("departamento") == dept])
            icon = DEPARTMENT_ICONS.get(dept, "📁")
            
            # Use emoji or text icon
            icon_map = {
                "fa-building": "🏢",
                "fa-handshake": "🤝",
                "fa-mountain": "⛰️",
                "fa-water": "💧",
                "fa-fire": "🔥",
                "fa-wrench": "🔧",
                "fa-recycle": "♻️",
                "fa-gas-pump": "⛽",
                "fa-users": "👥",
                "fa-dollar-sign": "💰",
                "fa-folder": "📁",
            }
            emoji = icon_map.get(icon, "📁")
            
            btn = QPushButton(f"{emoji}  {dept} ({count})")
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, d=dept: self.handle_open_department(d))
            
            key = f"dept_{dept}"
            self.nav_buttons[key] = btn
            self.dept_layout.addWidget(btn)
    
    def _update_nav(self, active_key: str):
        """Atualiza estado ativo da navegação."""
        for key, btn in self.nav_buttons.items():
            btn.setChecked(key == active_key)
    
    # ── Navigation Handlers ──
    
    def handle_navigate(self, page: str):
        if page == "dashboard":
            self.load_dashboard()
    
    def load_dashboard(self):
        """Carrega o dashboard principal."""
        self.current_view = "dashboard"
        self.current_department = None
        self._update_nav("dashboard")
        
        stats = self.data_manager.get_dashboard_stats()
        departments = self.data_manager.get_all_departments()
        contracts = self.data_manager.get_all_contracts()
        
        html = HTMLGenerator.generate_dashboard(stats, departments, contracts)
        self.web_view.setHtml(html)
        self._refresh_dept_buttons()
    
    def handle_open_department(self, department: str):
        """Abre a página de um departamento."""
        self.current_view = "department"
        self.current_department = department
        self._update_nav(f"dept_{department}")
        
        contracts = self.data_manager.get_contracts_by_department(department)
        html = HTMLGenerator.generate_department_page(department, contracts)
        self.web_view.setHtml(html)
    
    def handle_add_contract(self, department: str):
        """Abre o formulário de novo contrato."""
        self.current_view = "contract_form"
        self.current_department = department
        self.current_contract_id = None
        
        html = HTMLGenerator.generate_contract_form(department)
        self.web_view.setHtml(html)
    
    def handle_view_contract(self, contract_id: str):
        """Abre os detalhes de um contrato."""
        contract = self.data_manager.get_contract_by_id(contract_id)
        if not contract:
            QMessageBox.warning(self, "Erro", "Contrato não encontrado.")
            return
        
        self.current_view = "contract_detail"
        self.current_contract_id = contract_id
        department = contract.get("departamento", "")
        
        html = HTMLGenerator.generate_contract_detail(contract, department)
        self.web_view.setHtml(html)
    
    def handle_edit_contract(self, contract_id: str):
        """Abre o formulário de edição de contrato."""
        contract = self.data_manager.get_contract_by_id(contract_id)
        if not contract:
            QMessageBox.warning(self, "Erro", "Contrato não encontrado.")
            return
        
        self.current_view = "contract_form"
        self.current_contract_id = contract_id
        department = contract.get("departamento", "")
        
        html = HTMLGenerator.generate_contract_form(department, contract)
        self.web_view.setHtml(html)
    
    def handle_delete_contract(self, contract_id: str):
        """Exclui um contrato."""
        contract = self.data_manager.get_contract_by_id(contract_id)
        if not contract:
            return
        
        department = contract.get("departamento", "")
        if self.data_manager.delete_contract(contract_id):
            QMessageBox.information(self, "Sucesso", "Contrato excluído com sucesso!")
            self.handle_open_department(department)
        else:
            QMessageBox.warning(self, "Erro", "Erro ao excluir contrato.")
    
    def handle_save_contract(self, json_data: str):
        """Salva ou atualiza um contrato a partir do formulário."""
        try:
            data = json.loads(json_data)
        except:
            QMessageBox.warning(self, "Erro", "Dados inválidos do formulário.")
            return
        
        contract_id = data.get("id", "")
        department = data.get("departamento", "")
        
        # Validações
        if not data.get("numero_contrato"):
            QMessageBox.warning(self, "Campo Obrigatório", "Informe o número do contrato.")
            return
        if not data.get("fornecedor"):
            QMessageBox.warning(self, "Campo Obrigatório", "Informe o fornecedor.")
            return
        
        if contract_id:
            # Atualização
            self.data_manager.update_contract(contract_id, data)
            QMessageBox.information(self, "Sucesso", "Contrato atualizado com sucesso!")
            self.handle_view_contract(contract_id)
        else:
            # Novo contrato
            new_id = self.data_manager.add_contract(data)
            QMessageBox.information(self, "Sucesso", "Contrato cadastrado com sucesso!")
            self.handle_view_contract(new_id)
        
        self._refresh_dept_buttons()
    
    def handle_attach_file(self, contract_id: str):
        """Abre diálogo para anexar arquivo ao contrato."""
        contract = self.data_manager.get_contract_by_id(contract_id)
        if not contract:
            QMessageBox.warning(self, "Erro", "Contrato não encontrado.")
            return
        
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecionar Documentos para Anexar",
            "",
            "Todos os Arquivos (*.*);;PDF (*.pdf);;Excel (*.xlsx *.xls);;Word (*.docx *.doc);;Imagens (*.jpg *.jpeg *.png *.gif)"
        )
        
        if file_paths:
            count = 0
            for fp in file_paths:
                result = self.data_manager.add_attachment(contract_id, fp)
                if result:
                    count += 1
            
            if count > 0:
                QMessageBox.information(self, "Sucesso", f"{count} documento(s) anexado(s) com sucesso!")
                self.handle_view_contract(contract_id)
            else:
                QMessageBox.warning(self, "Erro", "Não foi possível anexar os documentos.")
    
    def handle_download_contract(self, contract_id: str):
        """Abre o contrato físico (pasta de rede ou arquivo local)."""
        contract = self.data_manager.get_contract_by_id(contract_id)
        if not contract:
            return
        
        caminho = contract.get("caminho_contrato_fisico", "")
        if not caminho:
            # Verificar se tem anexos
            anexos = contract.get("anexos", [])
            if anexos:
                # Abrir pasta de anexos
                contract_dir = os.path.join(ATTACHMENTS_DIR, contract_id)
                if os.path.exists(contract_dir):
                    QDesktopServices.openUrl(QUrl.fromLocalFile(contract_dir))
                    return
            
            QMessageBox.information(self, "Informação", "Nenhum caminho de contrato físico configurado.\nEdite o contrato para adicionar o caminho do contrato físico.")
            return
        
        # Tentar abrir o caminho
        if os.path.exists(caminho):
            QDesktopServices.openUrl(QUrl.fromLocalFile(caminho))
        else:
            QMessageBox.warning(self, "Erro", f"Caminho não encontrado:\n{caminho}\n\nVerifique se o caminho de rede está acessível.")
    
    def handle_open_propostas(self, contract_id: str):
        """Abre o link das propostas."""
        contract = self.data_manager.get_contract_by_id(contract_id)
        if not contract:
            return
        
        link = contract.get("link_propostas", "")
        if link:
            QDesktopServices.openUrl(QUrl(link))
        else:
            QMessageBox.information(self, "Informação", "Nenhum link de propostas configurado.\nEdite o contrato para adicionar o link.")
    
    def handle_add_department(self):
        """Adiciona uma nova categoria/departamento."""
        text, ok = QInputDialog.getText(
            self, 
            "Nova Categoria",
            "Nome da nova categoria/departamento:",
            QLineEdit.Normal,
            ""
        )
        
        if ok and text.strip():
            name = text.strip()
            if self.data_manager.add_department(name):
                QMessageBox.information(self, "Sucesso", f"Categoria '{name}' criada com sucesso!")
                self._refresh_dept_buttons()
                if self.current_view == "dashboard":
                    self.load_dashboard()
            else:
                QMessageBox.warning(self, "Erro", f"A categoria '{name}' já existe.")
    
    def handle_remove_department(self, department: str):
        """Remove uma categoria customizada."""
        if self.data_manager.remove_department(department):
            QMessageBox.information(self, "Sucesso", f"Categoria '{department}' removida.")
            self._refresh_dept_buttons()
            self.load_dashboard()
        else:
            QMessageBox.warning(self, "Erro", "Não é possível remover categorias padrão do sistema.")
    
    def handle_add_medicao(self, contract_id: str):
        """Abre diálogo para registrar nova medição."""
        contract = self.data_manager.get_contract_by_id(contract_id)
        if not contract:
            return
        
        dialog = MedicaoDialog(self, contract)
        if dialog.exec_() == QDialog.Accepted:
            medicao_data = dialog.get_data()
            if self.data_manager.add_medicao(contract_id, medicao_data):
                QMessageBox.information(self, "Sucesso", "Medição registrada com sucesso!")
                self.handle_view_contract(contract_id)
    
    def handle_view_medicao(self, contract_id: str):
        """Visualiza medições de um contrato."""
        self.handle_view_contract(contract_id)
    
    def handle_open_file(self, file_path: str):
        """Abre um arquivo usando o programa padrão do sistema."""
        # Normalizar caminho
        file_path = file_path.replace('/', '\\')
        if os.path.exists(file_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
        else:
            QMessageBox.warning(self, "Erro", f"Arquivo não encontrado:\n{file_path}")


# ═══════════════════════════════════════════════════════════════════════════════
# DIÁLOGO DE MEDIÇÃO
# ═══════════════════════════════════════════════════════════════════════════════

class MedicaoDialog(QDialog):
    """Diálogo para registrar medição de contrato."""
    
    def __init__(self, parent, contract: dict):
        super().__init__(parent)
        self.contract = contract
        self.setup_ui()
    
    def setup_ui(self):
        self.setWindowTitle("Registrar Medição")
        self.setFixedSize(500, 550)
        self.setStyleSheet("""
            QDialog {
                background: #f7fafc;
            }
            QLabel {
                font-size: 13px;
                font-weight: 600;
                color: #4a5568;
            }
            QLineEdit, QDoubleSpinBox, QComboBox, QDateEdit, QTextEdit {
                padding: 10px;
                border: 2px solid #e2e8f0;
                border-radius: 8px;
                font-size: 13px;
                background: white;
            }
            QLineEdit:focus, QDoubleSpinBox:focus, QComboBox:focus, QDateEdit:focus, QTextEdit:focus {
                border-color: #5a67d8;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)
        
        # Título
        title = QLabel(f"📏 Medição - {self.contract.get('numero_contrato', '')}")
        title.setStyleSheet("QLabel { font-size: 18px; font-weight: bold; color: #2d3748; }")
        layout.addWidget(title)
        
        # Número da medição
        num_medicoes = len(self.contract.get("medicoes", []))
        
        form = QFormLayout()
        form.setSpacing(12)
        
        self.numero_input = QLineEdit(str(num_medicoes + 1))
        form.addRow("Nº da Medição:", self.numero_input)
        
        self.periodo_input = QLineEdit()
        self.periodo_input.setPlaceholderText("Ex: Janeiro/2026")
        form.addRow("Período de Referência:", self.periodo_input)
        
        self.data_input = QDateEdit()
        self.data_input.setDate(QDate.currentDate())
        self.data_input.setCalendarPopup(True)
        form.addRow("Data da Medição:", self.data_input)
        
        self.valor_input = QDoubleSpinBox()
        self.valor_input.setMaximum(999999999.99)
        self.valor_input.setDecimals(2)
        self.valor_input.setPrefix("R$ ")
        form.addRow("Valor (R$):", self.valor_input)
        
        self.status_input = QComboBox()
        self.status_input.addItems(["Ativo", "Pendente", "Aprovada", "Em Análise", "Rejeitada", "Paga"])
        form.addRow("Status:", self.status_input)
        
        self.obs_input = QTextEdit()
        self.obs_input.setMaximumHeight(80)
        self.obs_input.setPlaceholderText("Observações sobre a medição...")
        form.addRow("Observação:", self.obs_input)
        
        layout.addLayout(form)
        layout.addStretch()
        
        # Botões
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setStyleSheet("""
            QPushButton {
                padding: 10px 24px; border: 2px solid #e2e8f0; border-radius: 8px;
                background: white; color: #4a5568; font-weight: 600;
            }
            QPushButton:hover { background: #f7fafc; }
        """)
        btn_cancel.clicked.connect(self.reject)
        
        btn_save = QPushButton("💾 Registrar Medição")
        btn_save.setStyleSheet("""
            QPushButton {
                padding: 10px 24px; border: none; border-radius: 8px;
                background: #5a67d8; color: white; font-weight: 600;
            }
            QPushButton:hover { background: #4c51bf; }
        """)
        btn_save.clicked.connect(self.accept)
        
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_save)
        layout.addLayout(btn_layout)
    
    def get_data(self) -> dict:
        return {
            "numero": self.numero_input.text(),
            "periodo_referencia": self.periodo_input.text(),
            "data_medicao": self.data_input.date().toString("yyyy-MM-dd"),
            "valor": self.valor_input.value(),
            "status": self.status_input.currentText(),
            "observacao": self.obs_input.toPlainText(),
        }


# ═══════════════════════════════════════════════════════════════════════════════
# SPLASH SCREEN
# ═══════════════════════════════════════════════════════════════════════════════

class ContractSplashScreen(QWidget):
    """Tela de splash/carregamento."""
    
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(480, 340)
        
        # Centralizar
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #1a1a2e, stop:0.5 #16213e, stop:1 #0f3460);
                border-radius: 20px;
            }
        """)
        
        c_layout = QVBoxLayout(container)
        c_layout.setContentsMargins(40, 40, 40, 40)
        c_layout.setSpacing(16)
        
        # Ícone
        icon_label = QLabel("📋")
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet("QLabel { font-size: 48px; background: transparent; }")
        c_layout.addWidget(icon_label)
        
        # Título
        title = QLabel("Gestão de Contratos e Medição")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("QLabel { color: white; font-size: 22px; font-weight: bold; background: transparent; }")
        c_layout.addWidget(title)
        
        sub = QLabel("Carregando sistema...")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet("QLabel { color: rgba(255,255,255,0.6); font-size: 13px; background: transparent; }")
        c_layout.addWidget(sub)
        
        c_layout.addSpacing(10)
        
        # Progress bar
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                background: rgba(255,255,255,0.1);
                border: none;
                border-radius: 4px;
                height: 6px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, x2:1, stop:0 #5a67d8, stop:1 #48bb78);
                border-radius: 4px;
            }
        """)
        c_layout.addWidget(self.progress)
        
        # Status
        self.status_label = QLabel("Inicializando...")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("QLabel { color: rgba(255,255,255,0.5); font-size: 11px; background: transparent; }")
        c_layout.addWidget(self.status_label)
        
        # Versão
        ver = QLabel(f"v{APP_VERSION}")
        ver.setAlignment(Qt.AlignCenter)
        ver.setStyleSheet("QLabel { color: rgba(255,255,255,0.3); font-size: 10px; background: transparent; }")
        c_layout.addWidget(ver)
        
        layout.addWidget(container)
        
        # Timer para animação de progresso
        self.timer = QTimer()
        self.timer.timeout.connect(self._advance_progress)
        self.progress_value = 0
        self.timer.start(30)
    
    def _advance_progress(self):
        self.progress_value += 2
        self.progress.setValue(self.progress_value)
        
        if self.progress_value < 30:
            self.status_label.setText("Carregando módulos...")
        elif self.progress_value < 60:
            self.status_label.setText("Inicializando banco de dados...")
        elif self.progress_value < 80:
            self.status_label.setText("Configurando interface...")
        elif self.progress_value < 100:
            self.status_label.setText("Quase pronto...")
        
        if self.progress_value >= 100:
            self.timer.stop()
            QTimer.singleShot(300, self._open_main)
    
    def _open_main(self):
        self.main_window = ContractMainWindow()
        self.main_window.showMaximized()
        self.close()


# ═══════════════════════════════════════════════════════════════════════════════
# PONTO DE ENTRADA
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    
    # Estilo global
    app.setStyleSheet("""
        QToolTip {
            background: #2d3748;
            color: white;
            border: none;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 12px;
        }
    """)
    
    splash = ContractSplashScreen()
    splash.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
