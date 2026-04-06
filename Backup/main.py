import sys
import os
import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Any
import locale
import traceback
import io
import tempfile

# Imports para geração de PDF profissional
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.barcharts import VerticalBarChart, HorizontalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
import matplotlib
matplotlib.use('Agg')  # Backend não-interativo
import matplotlib.pyplot as plt

# Configurar locale para formato brasileiro
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
    QGroupBox, QCheckBox, QTabWidget, QTextEdit, QProgressDialog, QAction
)
from PyQt5.QtCore import Qt, QUrl, QDate, QTimer, pyqtSignal, QThread, QSize, QRect, QObject, pyqtSlot, QFileSystemWatcher
from PyQt5.QtGui import QIcon, QFont, QColor, QPalette, QPixmap
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWebChannel import QWebChannel


# ═══════════════════════════════════════════════════════════════════════════════
# FUNÇÕES UTILITÁRIAS
# ═══════════════════════════════════════════════════════════════════════════════

def sanitize_for_js(text: str) -> str:
    """
    Sanitiza uma string para uso seguro em JavaScript.
    Remove ou escapa caracteres problemáticos.
    """
    if not isinstance(text, str):
        text = str(text) if text is not None else ''
    
    # Remover caracteres de controle e quebras de linha
    text = text.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    
    # Remover caracteres problemáticos para JSON/JavaScript
    text = text.replace('\\', '/').replace('"', "'").replace('`', "'")
    
    # Remover outros caracteres de controle unicode
    text = ''.join(char for char in text if ord(char) >= 32 or char in '\n\r\t')
    
    # Limpar espaços múltiplos
    while '  ' in text:
        text = text.replace('  ', ' ')
    
    return text.strip()


def calcular_dias_uteis(data_inicio: datetime, data_fim: datetime = None) -> int:
    """
    Calcula o número de dias úteis entre duas datas (excluindo sábado e domingo).
    Se data_fim não for informada, usa a data/hora atual.
    """
    if data_fim is None:
        data_fim = datetime.now(timezone.utc)
    
    # Garantir que ambas as datas tenham timezone
    if data_inicio.tzinfo is None:
        data_inicio = data_inicio.replace(tzinfo=timezone.utc)
    if data_fim.tzinfo is None:
        data_fim = data_fim.replace(tzinfo=timezone.utc)
    
    dias_uteis = 0
    data_atual = data_inicio.date()
    data_final = data_fim.date()
    
    while data_atual <= data_final:
        # 5 = sábado, 6 = domingo
        if data_atual.weekday() < 5:
            dias_uteis += 1
        data_atual += timedelta(days=1)
    
    # Se ainda está no mesmo dia, conta como 0 ou 1 dependendo se é dia útil
    if data_inicio.date() == data_fim.date():
        if data_inicio.date().weekday() < 5:
            return 0  # Mesmo dia útil = 0 dias passados
        else:
            return 0  # Fim de semana = 0 dias úteis
    
    # Subtrai 1 porque começamos a contar do dia seguinte
    return max(0, dias_uteis - 1)


# Configuração de SLA por fase (em dias úteis) - categoria -> fase -> limite em dias
SLA_FASES = {
    'compras_locais': {
        'cotação': 2,  # 2 dias úteis
    },
    'compras_csc': {
        'cotação': 11,  # 11 dias úteis
    },
    'reserva_materiais': {
        'entregue': 2,  # 2 dias úteis
    }
}


def verificar_vencimento_por_fase(card: dict, categoria: str) -> bool:
    """
    Verifica se o card está vencido na fase atual baseado no SLA configurado.
    
    Configurações:
    - Compras Locais: vence após 2 dias úteis na fase Cotação
    - Compras CSC: vence após 11 dias úteis na fase Cotação
    - Reserva de Materiais: vence após 2 dias úteis na fase Pendente
    
    Retorna True se o card está vencido.
    IMPORTANTE: Só considera vencido se o card AINDA está na fase configurada.
    
    Usa phases_history para calcular o tempo real na fase.
    """
    # Verificar se a categoria tem SLA configurado
    if categoria not in SLA_FASES:
        return False
    
    # Verificar se não está concluído
    if card.get('finished_at'):
        return False
    
    # Obter fase atual do card
    phase = card.get('current_phase', {})
    phase_name = (phase.get('name', '') if phase else '').lower().strip()
    
    if not phase_name:
        return False
    
    # Verificar se a fase atual tem SLA configurado para esta categoria
    sla_categoria = SLA_FASES.get(categoria, {})
    limite_dias = None
    fase_alvo = None
    
    for fase_config, dias_limite in sla_categoria.items():
        fase_config_lower = fase_config.lower()
        # Verificar se o nome da fase atual contém o nome da fase configurada
        if fase_config_lower in phase_name or phase_name in fase_config_lower:
            limite_dias = dias_limite
            fase_alvo = fase_config_lower
            break
    
    # Se não encontrou SLA para a fase atual, não está vencido
    if limite_dias is None:
        return False
    
    # Método 1: Calcular tempo na fase usando phases_history (mais confiável)
    for ph in card.get('phases_history', []):
        ph_name = (ph.get('phase', {}).get('name', '') or '').lower().strip()
        
        # Verificar se é a fase alvo
        if fase_alvo in ph_name or ph_name in fase_alvo:
            first_in = ph.get('firstTimeIn')
            last_out = ph.get('lastTimeOut')
            
            if first_in:
                try:
                    start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                    # Se não saiu da fase (lastTimeOut é None), calcula até agora
                    if last_out:
                        end = datetime.fromisoformat(last_out.replace('Z', '+00:00'))
                    else:
                        end = datetime.now(timezone.utc)
                    
                    # Calcular dias úteis
                    dias_na_fase = calcular_dias_uteis(start, end)
                    
                    if dias_na_fase > limite_dias:
                        return True
                except Exception:
                    pass
    
    # Método 2 (fallback): Buscar campo "Tempo Total na Fase X" nos fields do card
    tempo_fase = None
    for field in card.get('fields', []):
        field_name = (field.get('name') or '').lower().strip()
        field_value = field.get('value')
        
        # Verificar variações do nome do campo
        is_tempo_fase = False
        
        # Verificar se contém as palavras-chave da fase
        if ('tempo' in field_name) and (fase_alvo in field_name):
            is_tempo_fase = True
        elif ('fase' in field_name) and (fase_alvo in field_name):
            is_tempo_fase = True
        
        if is_tempo_fase and field_value is not None:
            tempo_fase = field_value
            break
    
    if tempo_fase is not None:
        try:
            tempo_str = str(tempo_fase).strip()
            
            if not tempo_str or tempo_str.lower() == 'none' or tempo_str == '':
                return False
            
            dias = 0
            import re
            
            if tempo_str.lower().endswith('d'):
                match = re.search(r'(\d+(?:[.,]\d+)?)\s*d', tempo_str.lower())
                if match:
                    dias = float(match.group(1).replace(',', '.'))
            elif 'dia' in tempo_str.lower():
                match = re.search(r'(\d+(?:[.,]\d+)?)', tempo_str)
                if match:
                    dias = float(match.group(1).replace(',', '.'))
            elif 'h' in tempo_str.lower() or 'hora' in tempo_str.lower():
                match = re.search(r'(\d+(?:[.,]\d+)?)', tempo_str)
                if match:
                    horas = float(match.group(1).replace(',', '.'))
                    dias = horas / 24
            else:
                match = re.search(r'(\d+(?:[.,]\d+)?)', tempo_str)
                if match:
                    valor = float(match.group(1).replace(',', '.'))
                    if valor > 48:
                        dias = valor / 24
                    else:
                        dias = valor
            
            if dias > limite_dias:
                return True
                
        except Exception:
            pass
    
    return False


# Manter função antiga como alias para compatibilidade
def verificar_vencimento_fase_cotacao(card: dict, categoria: str) -> bool:
    """Alias para verificar_vencimento_por_fase para compatibilidade."""
    return verificar_vencimento_por_fase(card, categoria)


def enviar_email_relatorio(
    remetente_email: str,
    remetente_senha: str,
    destinatarios: List[str],
    assunto: str,
    corpo: str,
    anexo_pdf: Optional[str] = None
) -> bool:
    """
    Envia email com relatório em anexo usando Office 365.
    
    Args:
        remetente_email: Email do remetente (usuário logado)
        remetente_senha: Senha do email do remetente
        destinatarios: Lista de emails destinatários
        assunto: Assunto do email
        corpo: Corpo do email em HTML
        anexo_pdf: Caminho opcional para arquivo PDF em anexo
    
    Returns:
        True se enviado com sucesso, False caso contrário
    """
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    
    try:
        # Criar mensagem
        msg = MIMEMultipart()
        msg['From'] = remetente_email
        msg['To'] = ', '.join(destinatarios)
        msg['Subject'] = assunto
        
        # Corpo do email em HTML
        msg.attach(MIMEText(corpo, 'html'))
        
        # Anexar PDF se fornecido
        if anexo_pdf and os.path.exists(anexo_pdf):
            with open(anexo_pdf, 'rb') as f:
                parte = MIMEBase('application', 'octet-stream')
                parte.set_payload(f.read())
            
            encoders.encode_base64(parte)
            parte.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(anexo_pdf)}'
            )
            msg.attach(parte)
        
        # Conectar ao servidor Office 365
        servidor = smtplib.SMTP('smtp.office365.com', 587)
        servidor.starttls()
        servidor.login(remetente_email, remetente_senha)
        
        # Enviar email
        texto = msg.as_string()
        servidor.sendmail(remetente_email, destinatarios, texto)
        servidor.quit()
        
        return True
        
    except Exception as e:
        print(f"Erro ao enviar email: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTES E CONFIGURAÇÕES GLOBAIS
# ═══════════════════════════════════════════════════════════════════════════════

APP_NAME = "Painel Inteligente de Gestão"
APP_VERSION = "2.0.0"
COMPANY_NAME = "Sistema Empresarial"

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

# Credenciais Pipefy
PIPEFY_CLIENT_ID = "ofgUSnXFhXadEzrDd_ZtUzXsV8-Crv-0NFboRn0CbrU"
PIPEFY_CLIENT_SECRET = "DyLPVER8t6SIeVpO7lQiDqTzoquM3UqLDOUgOFtHFpw"
PIPEFY_TOKEN_URL = "https://app.pipefy.com/oauth/token"
PIPEFY_API_URL = "https://api.pipefy.com/graphql"

# IDs dos pipes Pipefy
PIPE_CONTRATACAO_SERVICOS = "306527874"
PIPE_COMPRAS_CSC = "306864423"
PIPE_COMPRAS_LOCAIS = "306858226"
PIPE_ENVIO_NFE = "306859940"
PIPE_RESERVA_MATERIAIS = "306859726"

# Emails autorizados
AUTHORIZED_EMAILS = [
    "earaujo@essencis.com.br",
    "gmartins@essencis.com.br"
]

# Destinatários por categoria
EMAIL_RECIPIENTS = {
    'compras_servicos': ['jcsouza@essencis.com.br', 'gmartins@essencis.com.br'],
    'compras_csc': ['acsouza@essencis.com.br'],
    'compras_locais': [],  # Não especificado
    'envio_nfe': ['iferreira@essencis.com.br'],
    'reserva_materiais': ['eassis@essencis.com.br', 'agsantos@essencis.com.br'],
}


# ═══════════════════════════════════════════════════════════════════════════════
# LOGIN SCREEN - TELA DE AUTENTICAÇÃO
# ═══════════════════════════════════════════════════════════════════════════════

class LoginScreen(QDialog):
    """Tela de login para autenticação."""
    
    def __init__(self):
        super().__init__()
        self.user_email = None
        self.user_password = None
        self.setup_ui()
        
    def setup_ui(self):
        """Configura a interface de login."""
        self.setWindowTitle("Login - Painel de Gestão")
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(450, 550)
        
        # Layout principal
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Container
        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #e8e8e8, stop:0.5 #f5f5f5, stop:1 #e8e8e8);
                border-radius: 20px;
            }
        """)
        
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(40, 40, 40, 40)
        container_layout.setSpacing(20)
        
        # Botão fechar no topo
        close_btn_layout = QHBoxLayout()
        close_btn_layout.addStretch()
        close_btn = QPushButton("✕")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setFixedSize(30, 30)
        close_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #666666;
                border: none;
                font-size: 20px;
                font-weight: bold;
                border-radius: 15px;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
                color: #000000;
            }
        """)
        close_btn.clicked.connect(self.reject)
        close_btn_layout.addWidget(close_btn)
        container_layout.addLayout(close_btn_layout)
        
        # Subtítulo (agora como título principal)
        subtitle = QLabel("Login")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("""
            QLabel {
                color: #000000;
                font-size: 16px;
                font-weight: bold;
                background: transparent;
                margin-bottom: 20px;
                padding: 30px 10px;
                line-height: 1.5;
            }
        """)
        container_layout.addWidget(subtitle)
        
        # Campo de email
        email_label = QLabel("Email:")
        email_label.setStyleSheet("QLabel { color: #000000; font-weight: bold; background: transparent; }")
        container_layout.addWidget(email_label)
        
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("seu.email@essencis.com.br")
        self.email_input.setStyleSheet("""
            QLineEdit {
                padding: 12px;
                border: 2px solid #cccccc;
                border-radius: 8px;
                background: white;
                color: #000000;
                font-size: 14px;
            }
            QLineEdit:focus {
                border: 2px solid #000000;
            }
        """)
        container_layout.addWidget(self.email_input)
        
        # Campo de senha
        password_label = QLabel("Senha:")
        password_label.setStyleSheet("QLabel { color: #000000; font-weight: bold; background: transparent; }")
        container_layout.addWidget(password_label)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Sua senha de email")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setStyleSheet("""
            QLineEdit {
                padding: 12px;
                border: 2px solid #cccccc;
                border-radius: 8px;
                background: white;
                color: #000000;
                font-size: 14px;
            }
            QLineEdit:focus {
                border: 2px solid #000000;
            }
        """)
        self.password_input.returnPressed.connect(self.do_login)
        container_layout.addWidget(self.password_input)
        
        # Mensagem de erro
        self.error_label = QLabel("")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("""
            QLabel {
                color: #dc2626;
                font-size: 12px;
                background: transparent;
                min-height: 20px;
            }
        """)
        container_layout.addWidget(self.error_label)
        
        # Botão de login
        login_btn = QPushButton("Entrar")
        login_btn.setCursor(Qt.PointingHandCursor)
        login_btn.setStyleSheet("""
            QPushButton {
                background: #000000;
                color: white;
                padding: 14px;
                border: none;
                border-radius: 8px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #333333;
                box-shadow: 0 0 10px #00ff00;
            }
            QPushButton:pressed {
                background: #1a1a1a;
            }
        """)
        login_btn.clicked.connect(self.do_login)
        container_layout.addWidget(login_btn)
        
        # Info
        info_label = QLabel("⚠️ Apenas usuários autorizados")
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 11px;
                background: transparent;
                padding-top: 10px;
            }
        """)
        container_layout.addWidget(info_label)
        
        # Versão
        version_label = QLabel(f"v{APP_VERSION}")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet("""
            QLabel {
                color: #999999;
                font-size: 10px;
                background: transparent;
            }
        """)
        container_layout.addWidget(version_label)
        
        main_layout.addWidget(container)
        
        # Centralizar
        self.center_on_screen()
        
    def center_on_screen(self):
        """Centraliza a janela na tela."""
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)
        
    def do_login(self):
        """Realiza o login."""
        email = self.email_input.text().strip().lower()
        password = self.password_input.text()
        
        if not email or not password:
            self.error_label.setText("❌ Preencha todos os campos")
            return
        
        if email not in AUTHORIZED_EMAILS:
            self.error_label.setText("❌ Usuário não autorizado")
            return
        
        # Armazenar credenciais (senha será usada para envio de email)
        self.user_email = email
        self.user_password = password
        self.accept()


# ═══════════════════════════════════════════════════════════════════════════════
# SPLASH SCREEN - TELA DE CARREGAMENTO INICIAL
# ═══════════════════════════════════════════════════════════════════════════════

class LoadingWorker(QThread):
    """Thread para executar carregamento em background."""
    progress = pyqtSignal(int, str)  # progresso, mensagem
    finished = pyqtSignal(dict)  # dados carregados
    error = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data = {}
        
    def run(self):
        """Executa o carregamento."""
        try:
            steps = [
                (10, "Inicializando sistema...", self.init_system),
                (25, "Autenticando no Pipefy...", self.authenticate_pipefy),
                (35, "Carregando Contratação de Serviços...", lambda: self.load_pipe('compras_servicos', PIPE_CONTRATACAO_SERVICOS)),
                (50, "Carregando Compras CSC...", lambda: self.load_pipe('compras_csc', PIPE_COMPRAS_CSC)),
                (65, "Carregando Compras Locais...", lambda: self.load_pipe('compras_locais', PIPE_COMPRAS_LOCAIS)),
                (80, "Carregando Envio de NFE...", lambda: self.load_pipe('envio_nfe', PIPE_ENVIO_NFE)),
                (90, "Carregando Reserva de Materiais...", lambda: self.load_pipe('reserva_materiais', PIPE_RESERVA_MATERIAIS)),
                (100, "Finalizando...", self.finalize),
            ]
            
            for progress, message, task in steps:
                self.progress.emit(progress, message)
                QThread.msleep(300)  # Pausa para visualização
                task()
                
            self.finished.emit(self.data)
            
        except Exception as e:
            self.error.emit(f"Erro durante carregamento: {str(e)}")
    
    def init_system(self):
        """Inicializa o sistema."""
        self.data['initialized'] = True
        
    def authenticate_pipefy(self):
        """Autentica no Pipefy."""
        try:
            client = PipefyClient(PIPEFY_CLIENT_ID, PIPEFY_CLIENT_SECRET, PIPEFY_TOKEN_URL)
            client.authenticate()
            self.data['pipefy_client'] = client
        except Exception as e:
            # Se falhar, continua sem dados do Pipefy
            print(f"Falha na autenticação Pipefy: {e}")
            self.data['pipefy_client'] = None
    
    def load_pipe(self, category, pipe_id):
        """Carrega dados de um pipe específico."""
        client = self.data.get('pipefy_client')
        if not client:
            return
        
        try:
            result = client.get_all_pipe_cards(pipe_id)
            if 'pipefy_data' not in self.data:
                self.data['pipefy_data'] = {}
            
            cards = result.get('cards', [])
            
            # REGRA ESPECÍFICA: Atribuição automática de responsáveis por categoria e fase
            if category in ['compras_csc', 'compras_locais', 'reserva_materiais']:
                andre_assignee = {
                    'id': 'andre_castro_souza',
                    'name': 'André Castro de Souza',
                    'email': 'acsouza@essencis.com.br'
                }
                
                eviane_assignee = {
                    'id': 'eviane',
                    'name': 'Eviane',
                    'email': 'eviane@essencis.com.br'
                }
                
                for card in cards:
                    # Reserva de Materiais: TODOS os cards são da Eviane
                    if category == 'reserva_materiais':
                        card['assignees'] = [eviane_assignee]
                        continue
                    
                    current_phase = card.get('current_phase', {})
                    if current_phase:
                        phase_name = current_phase.get('name', '').lower()
                        
                        # Compras CSC: Cotação e Pedido de Compra Parcial sempre André
                        # Outros cards sem responsável também recebem André
                        if category == 'compras_csc':
                            if 'cotação' in phase_name or 'cotacao' in phase_name or 'pedido de compra parcial' in phase_name:
                                card['assignees'] = [andre_assignee]
                            elif not card.get('assignees'):
                                # Se não tem responsável, atribui André como padrão
                                card['assignees'] = [andre_assignee]
                        
                        # Compras Locais: Aprovação Pedido e Cotação sempre André
                        # Outros cards sem responsável também recebem André
                        elif category == 'compras_locais':
                            if 'aprovação pedido' in phase_name or 'aprovacao pedido' in phase_name or 'cotação' in phase_name or 'cotacao' in phase_name:
                                card['assignees'] = [andre_assignee]
                            elif not card.get('assignees'):
                                # Se não tem responsável, atribui André como padrão
                                card['assignees'] = [andre_assignee]
            
            self.data['pipefy_data'][category] = {
                'cards': cards,
                'name': result.get('name', ''),
                'loading': False
            }
        except Exception as e:
            print(f"Erro ao carregar {category}: {e}")
    
    def finalize(self):
        """Finaliza o carregamento."""
        self.data['ready'] = True


class SplashScreen(QWidget):
    """Tela de carregamento inicial moderna."""
    
    def __init__(self):
        super().__init__()
        self.loaded_data = None
        self.setup_ui()
        self.start_loading()
        
    def setup_ui(self):
        """Configura a interface da splash screen."""
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Tamanho da janela
        self.setFixedSize(700, 500)
        
        # Layout principal
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Container com fundo
        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #e8e8e8, stop:0.5 #f5f5f5, stop:1 #e8e8e8);
                border-radius: 20px;
            }
        """)
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(40, 40, 40, 40)
        container_layout.setSpacing(20)
        
        # Logo/Título
        title = QLabel("⚙️ Painel de Gestão")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                color: #000000;
                font-size: 32px;
                font-weight: bold;
                background: transparent;
                padding: 20px;
            }
        """)
        container_layout.addWidget(title)
        
        # Área de descrição
        self.description_label = QLabel()
        self.description_label.setAlignment(Qt.AlignCenter)
        self.description_label.setWordWrap(True)
        self.description_label.setStyleSheet("""
            QLabel {
                color: #000000;
                font-size: 14px;
                background: transparent;
                padding: 15px;
                min-height: 80px;
            }
        """)
        container_layout.addWidget(self.description_label)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 10px;
                background: rgba(200, 200, 200, 0.3);
                height: 20px;
            }
            QProgressBar::chunk {
                border-radius: 10px;
                background: #000000;
                box-shadow: 0 0 10px #00ff00, 0 0 20px #00ff00, 0 0 30px #00ff00;
            }
        """)
        container_layout.addWidget(self.progress_bar)
        
        # Status
        self.status_label = QLabel("Iniciando...")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                color: #000000;
                font-size: 12px;
                background: transparent;
                padding-top: 5px;
            }
        """)
        container_layout.addWidget(self.status_label)
        
        # Versão
        version_label = QLabel(f"v{APP_VERSION}")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 10px;
                background: transparent;
                padding-top: 10px;
            }
        """)
        container_layout.addWidget(version_label)
        
        main_layout.addWidget(container)
        
        # Centralizar na tela
        self.center_on_screen()
        
        # Descrições rotativas
        self.descriptions = [
            "📊 Análise inteligente de dados empresariais em tempo real",
            "📦 Gestão completa de estoque e movimentações",
            "💰 Controle financeiro e pagamentos integrados",
            "📋 Acompanhamento de ordens de compra e processos",
            "🔄 Integração automática com Pipefy para gestão de atividades",
            "📈 Relatórios detalhados e exportação em PDF",
            "⚡ Dashboard interativo com métricas e indicadores",
            "🎯 Otimização de processos e tomada de decisão estratégica"
        ]
        self.current_description = 0
        
        # Timer para rotação de descrições
        self.description_timer = QTimer()
        self.description_timer.timeout.connect(self.update_description)
        self.description_timer.start(2000)  # Atualiza a cada 2 segundos
        self.update_description()
        
    def center_on_screen(self):
        """Centraliza a janela na tela."""
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)
        
    def update_description(self):
        """Atualiza a descrição exibida."""
        self.description_label.setText(self.descriptions[self.current_description])
        self.current_description = (self.current_description + 1) % len(self.descriptions)
        
    def start_loading(self):
        """Inicia o processo de carregamento."""
        self.worker = LoadingWorker()
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_loading_finished)
        self.worker.error.connect(self.on_loading_error)
        self.worker.start()
        
    def update_progress(self, value, message):
        """Atualiza a barra de progresso."""
        self.progress_bar.setValue(value)
        self.status_label.setText(message)
        
    def on_loading_finished(self, data):
        """Callback quando carregamento termina."""
        self.loaded_data = data
        self.description_timer.stop()
        self.status_label.setText("✓ Carregamento concluído!")
        QTimer.singleShot(500, self.close)  # Fecha após 500ms
        
    def on_loading_error(self, error_message):
        """Callback quando ocorre erro."""
        self.description_timer.stop()
        self.status_label.setText(f"Erro: {error_message}")
        QTimer.singleShot(3000, self.close)


# ═══════════════════════════════════════════════════════════════════════════════
# CLIENTE PIPEFY API
# ═══════════════════════════════════════════════════════════════════════════════

import requests

class PipefyClient:
    """Cliente para integração com Pipefy via API GraphQL."""
    
    def __init__(self, client_id: str, client_secret: str, token_url: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self.token_url = token_url
        self.api_url = PIPEFY_API_URL
        self.access_token = None
        
    def authenticate(self) -> bool:
        """Autentica com OAuth2 client credentials."""
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
        """Faz requisição GraphQL."""
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

    def get_all_pipe_cards(self, pipe_id: str, progress_callback=None) -> Dict:
        """Busca TODOS os cards com TODOS os campos possíveis do Pipefy."""
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

        return {
            'name': pipe_name,
            'cards': all_cards
        }

    def get_all_database_records(self, table_id: str, progress_callback=None) -> Dict:
        """Busca TODOS os registros de uma tabela (database) do Pipefy."""
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
                progress_callback(f"Database: página {page_count}... ({len(all_records)} registros)")
        
        return {
            'records': all_records
        }


class PipefyLoadThread(QThread):
    """Thread para carregar dados do Pipefy sem bloquear a UI."""
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    
    def __init__(self, client, pipe_id):
        super().__init__()
        self.client = client
        self.pipe_id = pipe_id
        
    def run(self):
        try:
            data = self.client.get_all_pipe_cards(
                self.pipe_id, 
                progress_callback=lambda msg: self.progress.emit(msg)
            )
            self.finished.emit(data)
        except Exception as e:
            self.error.emit(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
# CLASSE DE MANIPULAÇÃO DE DADOS
# ═══════════════════════════════════════════════════════════════════════════════

class DataHandler:
    """Classe responsável por manipular e processar dados de pagamentos e estoque."""
    
    def __init__(self):
        self.df_pagamentos: Optional[pd.DataFrame] = None
        self.df_estoque: Optional[pd.DataFrame] = None
        self.df_ordens: Optional[pd.DataFrame] = None
        self.df_movimentacoes: Optional[pd.DataFrame] = None
        self.df_ordens_compra: Optional[pd.DataFrame] = None  # Dados reais de ordens de compra
        self._generate_sample_data()
    
    def _generate_sample_data(self):
        """Gera dados de amostra para demonstração."""
        np.random.seed(42)
        
        # Lista de fornecedores
        fornecedores = [
            ("ABC Indústria LTDA", "12.345.678/0001-90"),
            ("XYZ Comércio S/A", "23.456.789/0001-01"),
            ("Fornecedor Nacional LTDA", "34.567.890/0001-12"),
            ("Tech Solutions ME", "45.678.901/0001-23"),
            ("Material Express", "56.789.012/0001-34"),
            ("Distribuidora Central", "67.890.123/0001-45"),
            ("Mega Suprimentos", "78.901.234/0001-56"),
            ("Industrial Parts", "89.012.345/0001-67"),
            ("Quality Components", "90.123.456/0001-78"),
            ("Prime Materials", "01.234.567/0001-89"),
            ("Fast Delivery LTDA", "11.222.333/0001-44"),
            ("Super Insumos", "22.333.444/0001-55"),
            ("Master Supply", "33.444.555/0001-66"),
            ("Global Trade", "44.555.666/0001-77"),
            ("Direct Parts", "55.666.777/0001-88"),
        ]
        
        categorias = ["Matéria-Prima", "Componentes", "Equipamentos", "Serviços", "Consumíveis", "Manutenção"]
        centros_custo = ["Produção", "Manutenção", "Administrativo", "Logística", "Qualidade"]
        formas_pagamento = ["Boleto", "Transferência", "Cartão", "PIX", "Cheque"]
        
        # Gerar pagamentos
        n_pagamentos = 500
        datas = pd.date_range(start='2024-01-01', end='2026-01-16', periods=n_pagamentos)
        
        pagamentos_data = []
        for i in range(n_pagamentos):
            forn = fornecedores[np.random.choice(len(fornecedores), p=[0.2, 0.15, 0.12, 0.1, 0.08, 0.07, 0.06, 0.05, 0.04, 0.04, 0.03, 0.02, 0.02, 0.01, 0.01])]
            valor = np.random.exponential(15000) + 500
            pagamentos_data.append({
                'Data do Pagamento': datas[i],
                'CNPJ do Fornecedor': forn[1],
                'Nome do Fornecedor': forn[0],
                'Valor Pago': round(valor, 2),
                'Número da Nota Fiscal': f"NF-{10000 + i}",
                'Centro de Custo': np.random.choice(centros_custo),
                'Categoria de Material': np.random.choice(categorias),
                'Descrição do Pagamento': f"Pagamento referente ao pedido {5000 + i}",
                'Forma de Pagamento': np.random.choice(formas_pagamento)
            })
        
        self.df_pagamentos = pd.DataFrame(pagamentos_data)
        
        # Gerar ordens de compra (dados de amostra)
        ordens_data = []
        for i in range(100):
            forn = fornecedores[np.random.randint(0, len(fornecedores))]
            status_opts = ["Emitida", "Aprovada", "Em Trânsito", "Recebida", "Cancelada"]
            ordens_data.append({
                'Número OC': f"OC-{2024000 + i}",
                'Data Emissão': datetime.now() - timedelta(days=np.random.randint(1, 180)),
                'Fornecedor': forn[0],
                'CNPJ': forn[1],
                'Valor Total': round(np.random.uniform(1000, 50000), 2),
                'Status': np.random.choice(status_opts, p=[0.1, 0.15, 0.2, 0.5, 0.05]),
                'Prazo Entrega': np.random.randint(5, 45),
                'Dias em Atraso': np.random.choice([0, 0, 0, np.random.randint(1, 15)]),
            })
        
        self.df_ordens = pd.DataFrame(ordens_data)
        
        # Gerar movimentações (dados de amostra)
        materiais_mov = ["Óleo Diesel", "Filtro", "Pneu", "Bateria", "Peça Motor", 
                         "Lubrificante", "Graxa", "Correia", "Rolamento", "Vedação"]
        
        mov_data = []
        for i in range(200):
            tipo = np.random.choice(["Entrada", "Saída"])
            mov_data.append({
                'Data': datetime.now() - timedelta(days=np.random.randint(1, 90)),
                'Tipo': tipo,
                'Material': np.random.choice(materiais_mov),
                'Quantidade': np.random.randint(5, 100),
                'Documento': f"{'NF' if tipo == 'Entrada' else 'REQ'}-{np.random.randint(10000, 99999)}",
                'Responsável': np.random.choice(["João Silva", "Maria Santos", "Pedro Costa", "Ana Lima"]),
                'Motivo': np.random.choice(["Produção", "Manutenção", "Reposição", "Projeto"])
            })
        
        self.df_movimentacoes = pd.DataFrame(mov_data)
        
        # Configuração do diretório base de relatórios
        self.base_relatorios = r"G:\SUPRIMENTOS\SUPRIMENTOS\PROJETOS\Relatórios"
        
        # Carregar arquivos reais do mês atual automaticamente
        self._carregar_arquivos_mes_atual()
    
    def _get_mes_ano_atual(self) -> tuple:
        """Retorna o mês e ano atual em português."""
        meses = {
            1: 'janeiro', 2: 'fevereiro', 3: 'marco', 4: 'abril',
            5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
            9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
        }
        agora = datetime.now()
        return meses[agora.month], str(agora.year)
    
    def _buscar_arquivo_na_pasta(self, pasta: str, extensoes: list = ['.xlsx', '.xls']) -> str:
        """Busca o primeiro arquivo com extensão válida na pasta."""
        if not os.path.exists(pasta):
            return None
        
        for arquivo in os.listdir(pasta):
            if any(arquivo.lower().endswith(ext) for ext in extensoes):
                return os.path.join(pasta, arquivo)
        return None
    
    def _carregar_arquivos_mes_atual(self):
        """Carrega automaticamente os arquivos do mês/ano atual."""
        mes, ano = self._get_mes_ano_atual()
        print(f"\n=== Carregando relatórios de {mes.upper()}/{ano} ===")
        
        # Carregar Pagamentos
        pasta_pagamentos = os.path.join(self.base_relatorios, 'Pagamentos', ano, mes)
        arquivo_pagamentos = self._buscar_arquivo_na_pasta(pasta_pagamentos)
        if arquivo_pagamentos:
            print(f"✓ Pagamentos: {arquivo_pagamentos}")
            self.load_payment_file(arquivo_pagamentos)
        else:
            print(f"✗ AVISO: Nenhum arquivo de pagamentos encontrado em {pasta_pagamentos}")
        
        # Carregar Estoque Periódico
        self.df_estoque = None
        pasta_estoque = os.path.join(self.base_relatorios, 'Estoque periodico', ano, mes)
        arquivo_estoque = self._buscar_arquivo_na_pasta(pasta_estoque)
        if arquivo_estoque:
            print(f"✓ Estoque: {arquivo_estoque}")
            self.load_estoque_file(arquivo_estoque)
        else:
            print(f"✗ AVISO: Nenhum arquivo de estoque encontrado em {pasta_estoque}")
        
        # Carregar Geral Contábil (Ordens de Compra)
        self.df_ordens_compra = None
        pasta_ordens = os.path.join(self.base_relatorios, 'Geral contabil', ano, mes)
        arquivo_ordens = self._buscar_arquivo_na_pasta(pasta_ordens)
        if arquivo_ordens:
            print(f"✓ Ordens de Compra: {arquivo_ordens}")
            self.load_ordens_compra_file(arquivo_ordens)
        else:
            print(f"✗ AVISO: Nenhum arquivo de ordens encontrado em {pasta_ordens}")
        
        print("=" * 50 + "\n")
    
    def load_file(self, filepath: str) -> bool:
        """Carrega arquivo Excel ou CSV."""
        try:
            if filepath.endswith('.csv'):
                self.df_pagamentos = pd.read_csv(filepath, encoding='utf-8')
            else:
                self.df_pagamentos = pd.read_excel(filepath)
            return True
        except Exception as e:
            print(f"Erro ao carregar arquivo: {e}")
            return False
    
    def load_estoque_file(self, filepath: str) -> bool:
        """Carrega arquivo de estoque periódico."""
        try:
            if filepath.endswith('.csv'):
                self.df_estoque = pd.read_csv(filepath, encoding='utf-8')
            else:
                self.df_estoque = pd.read_excel(filepath)
            
            # Mapeamento flexível de colunas
            column_mapping = {
                'Nome do Item': ['Nome do Item', 'Item', 'Código', 'Codigo', 'Material', 'SKU'],
                'Descrição do Item': ['Descrição do Item', 'Descricao do Item', 'Descrição', 'Descricao', 'Nome'],
                'Nome do Subinventário': ['Nome do Subinventário', 'Nome do Subinventario', 'Subinventário', 'Subinventario', 'Local', 'Armazém'],
                'Custo Total': ['Custo Total', 'Valor Total', 'Total', 'Valor'],
                'Quantidade': ['Quantidade', 'Qtd', 'Qtde', 'Estoque'],
                'Custo Unitário': ['Custo Unitário', 'Custo Unitario', 'Preço', 'Preco', 'Valor Unit'],
                'Organização': ['Organização', 'Organizacao', 'Empresa', 'Filial', 'Org'],
                'Nome da Organização': ['Nome da Organização', 'Nome da Organizacao', 'Nome Organização', 'Razão Social']
            }
            
            # Aplicar mapeamento
            for col_padrao, variacoes in column_mapping.items():
                if col_padrao not in self.df_estoque.columns:
                    for variacao in variacoes:
                        if variacao in self.df_estoque.columns:
                            self.df_estoque.rename(columns={variacao: col_padrao}, inplace=True)
                            print(f"Coluna '{variacao}' renomeada para '{col_padrao}'")
                            break
            
            # Verificar colunas críticas
            colunas_criticas = ['Nome do Item', 'Custo Total']
            colunas_faltando = [col for col in colunas_criticas if col not in self.df_estoque.columns]
            
            if colunas_faltando:
                print(f"AVISO: Colunas críticas faltando: {colunas_faltando}")
                print(f"Colunas disponíveis: {list(self.df_estoque.columns)}")
                return False
            
            # Garantir que Custo Total seja numérico
            if 'Custo Total' in self.df_estoque.columns:
                self.df_estoque['Custo Total'] = pd.to_numeric(self.df_estoque['Custo Total'], errors='coerce').fillna(0)
            
            # Garantir que Custo Unitário seja numérico
            if 'Custo Unitário' in self.df_estoque.columns:
                self.df_estoque['Custo Unitário'] = pd.to_numeric(self.df_estoque['Custo Unitário'], errors='coerce').fillna(0)
            
            # Garantir que Quantidade seja numérica
            if 'Quantidade' in self.df_estoque.columns:
                self.df_estoque['Quantidade'] = pd.to_numeric(self.df_estoque['Quantidade'], errors='coerce').fillna(0)
            
            print(f"Estoque carregado: {len(self.df_estoque)} linhas")
            print(f"Colunas encontradas: {list(self.df_estoque.columns)}")
            return True
            
        except Exception as e:
            print(f"Erro ao carregar arquivo de estoque: {e}")
            traceback.print_exc()
            return False
    
    def load_payment_file(self, filepath: str) -> bool:
        """Carrega arquivo de pagamentos (Relatório de Títulos Pagos)."""
        try:
            # Ler Excel pulando as primeiras 2 linhas (cabeçalho)
            df = pd.read_excel(filepath, skiprows=2)
            
            # A primeira linha contém os nomes das colunas
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
            
            # Remover linhas totalmente vazias
            df = df.dropna(how='all')
            
            self.df_pagamentos = df
            
            # Garantir que colunas numéricas sejam numéricas
            colunas_numericas = ['Valor Total', 'Valor Pago', 'ISS', 'IR', 'INSS', 'PIS', 'COFINS', 'CSLL', 'Valor do JUROS']
            for col in colunas_numericas:
                if col in self.df_pagamentos.columns:
                    self.df_pagamentos[col] = pd.to_numeric(self.df_pagamentos[col], errors='coerce').fillna(0)
            
            # Converter datas
            colunas_data = ['Data NFF', 'Data de Vencimento', 'Data do Pagamento', 'Data GL', 'Data Criação']
            for col in colunas_data:
                if col in self.df_pagamentos.columns:
                    self.df_pagamentos[col] = pd.to_datetime(self.df_pagamentos[col], errors='coerce')
            
            print(f"Pagamentos carregados: {len(self.df_pagamentos)} linhas")
            print(f"Colunas encontradas: {list(self.df_pagamentos.columns)}")
            return True
            
        except Exception as e:
            print(f"Erro ao carregar arquivo de estoque: {e}")
            traceback.print_exc()
            return False
    
    def load_payment_file(self, filepath: str) -> bool:
        """Carrega arquivo de relatório de pagamentos (Títulos Pagos)."""
        try:
            # Ler planilha com cabeçalhos na linha 3 (skiprows=2, então usa linha index 0 como header)
            df = pd.read_excel(filepath, skiprows=2)
            
            # A primeira linha contém os nomes das colunas
            df.columns = df.iloc[0]
            self.df_pagamentos = df[1:].reset_index(drop=True)
            
            # Converter colunas numéricas
            colunas_numericas = ['Valor Total', 'Valor Pago', 'ISS', 'IR', 'INSS', 'PIS', 'COFINS', 'CSLL']
            for col in colunas_numericas:
                if col in self.df_pagamentos.columns:
                    self.df_pagamentos[col] = pd.to_numeric(self.df_pagamentos[col], errors='coerce').fillna(0)
            
            # Processar datas
            colunas_data = ['Data NFF', 'Data de Vencimento', 'Data do Pagamento', 'Data GL', 'Data Criação']
            for col in colunas_data:
                if col in self.df_pagamentos.columns:
                    self.df_pagamentos[col] = pd.to_datetime(self.df_pagamentos[col], errors='coerce')
            
            # Filtrar usuários indesejados da coluna "Criado por"
            if 'Criado por' in self.df_pagamentos.columns:
                linhas_antes = len(self.df_pagamentos)
                
                # Emails específicos a excluir
                usuarios_excluidos = [
                    'tresende@viasolo.com.br',
                    'gcalisto@essencis.com.br',
                    'dcosta@essencis.com.br'
                ]
                
                # Filtrar emails específicos E todos que contenham @solvi.com ou oic.interface.prd
                mascara_excluir = (
                    self.df_pagamentos['Criado por'].isin(usuarios_excluidos) |
                    self.df_pagamentos['Criado por'].str.contains('@solvi.com', case=False, na=False) |
                    self.df_pagamentos['Criado por'].str.contains('oic.interface.prd', case=False, na=False)
                )
                
                self.df_pagamentos = self.df_pagamentos[~mascara_excluir].reset_index(drop=True)
                
                linhas_removidas = linhas_antes - len(self.df_pagamentos)
                if linhas_removidas > 0:
                    print(f"✓ Removidos {linhas_removidas} registros de usuários filtrados (@solvi.com, oic.interface.prd e outros)")
            
            print(f"Pagamentos carregados: {len(self.df_pagamentos)} linhas")
            print(f"Colunas encontradas: {list(self.df_pagamentos.columns)}")
            return True
            
        except Exception as e:
            print(f"Erro ao carregar arquivo de pagamentos: {e}")
            traceback.print_exc()
            return False
    
    def _classificar_grupo_estoque(self, codigo_item: str) -> str:
        """Classifica item de estoque em grupos baseado no código."""
        if pd.isna(codigo_item):
            return "NÃO CLASSIFICADO"
        
        codigo = str(codigo_item).strip()
        
        # Combustíveis
        if codigo.startswith('01.01.01'):
            return "COMBUSTÍVEIS"
        # Lubrificantes
        elif codigo.startswith('01.01.02'):
            return "LUBRIFICANTES"
        # Peças de frota
        elif codigo.startswith('01.01.03'):
            return "PEÇAS DE FROTA"
        # Componentes e implementos (01.01.04, 01.01.05, 01.01.10)
        elif codigo.startswith(('01.01.04', '01.01.05', '01.01.10')):
            return "COMPONENTES E IMPLEMENTOS"
        # Materiais administrativos
        elif codigo.startswith(('01.01.06', '01.01.07', '01.01.11', '01.01.21', '01.01.22', '01.01.23', '01.01.24')):
            return "MATERIAIS ADMINISTRATIVOS"
        # EPI/UNIFORMES
        elif codigo.startswith('01.01.08'):
            return "EPI/UNIFORMES"
        # Insumos
        elif codigo.startswith(('01.01.09', '01.01.12', '01.01.13', '01.01.14', '01.01.15', 
                               '01.01.16', '01.01.17', '01.01.18', '01.01.19', '01.01.20')):
            return "INSUMOS"
        else:
            return "OUTROS"
    
    def get_analise_estoque_por_grupo(self, organizacao: str = None) -> Dict[str, Any]:
        """Analisa estoque por grupos de materiais."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return {'resumo_grupos': [], 'top5_por_grupo': {}, 'total_geral': 0, 'total_itens': 0}
        
        # Verificar colunas obrigatórias
        if 'Nome do Item' not in self.df_estoque.columns:
            print("ERRO: Coluna 'Nome do Item' não encontrada")
            return {'resumo_grupos': [], 'top5_por_grupo': {}, 'total_geral': 0, 'total_itens': 0}
        
        if 'Custo Total' not in self.df_estoque.columns:
            print("ERRO: Coluna 'Custo Total' não encontrada")
            return {'resumo_grupos': [], 'top5_por_grupo': {}, 'total_geral': 0, 'total_itens': 0}
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        # Filtrar por organização se especificado
        if organizacao and 'Organização' in df.columns:
            df = df[df['Organização'] == organizacao]
        
        # Classificar grupos
        df['Grupo'] = df['Nome do Item'].apply(self._classificar_grupo_estoque)
        
        # Análise por grupo
        grupos_resumo = df.groupby('Grupo').agg({
            'Custo Total': 'sum',
            'Nome do Item': 'count'
        }).reset_index()
        grupos_resumo.columns = ['Grupo', 'Valor Total', 'Quantidade Itens']
        grupos_resumo = grupos_resumo.sort_values('Valor Total', ascending=False)
        
        # Top 5 itens mais caros por grupo
        top5_por_grupo = {}
        for grupo in df['Grupo'].unique():
            df_grupo = df[df['Grupo'] == grupo].nlargest(5, 'Custo Total')
            
            # Selecionar colunas disponíveis
            colunas_disponiveis = ['Nome do Item', 'Custo Total']
            if 'Descrição do Item' in df_grupo.columns:
                colunas_disponiveis.insert(1, 'Descrição do Item')
            if 'Quantidade' in df_grupo.columns:
                colunas_disponiveis.append('Quantidade')
            if 'Custo Unitário' in df_grupo.columns:
                colunas_disponiveis.append('Custo Unitário')
            
            top5_por_grupo[grupo] = df_grupo[colunas_disponiveis].to_dict('records')
        
        return {
            'resumo_grupos': grupos_resumo.to_dict('records'),
            'top5_por_grupo': top5_por_grupo,
            'total_geral': df['Custo Total'].sum(),
            'total_itens': len(df),
            'df_classificado': df  # Retornar DataFrame para uso posterior
        }
    
    def get_materiais_por_grupo(self, grupo: str, organizacao: str = None) -> List[Dict]:
        """Retorna todos os materiais de um grupo específico."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return []
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        # Filtrar por organização se especificado
        if organizacao and 'Organização' in df.columns:
            df = df[df['Organização'] == organizacao]
        
        # Classificar grupos
        df['Grupo'] = df['Nome do Item'].apply(self._classificar_grupo_estoque)
        
        # Filtrar pelo grupo específico
        df_grupo = df[df['Grupo'] == grupo].copy()
        
        # Ordenar por Custo Total decrescente
        df_grupo = df_grupo.sort_values('Custo Total', ascending=False)
        
        # Selecionar colunas para retornar
        colunas = ['Nome do Item', 'Descrição do Item', 'Quantidade', 'Custo Unitário', 'Custo Total']
        colunas_disponiveis = [col for col in colunas if col in df_grupo.columns]
        
        if 'Organização' in df_grupo.columns:
            colunas_disponiveis.insert(0, 'Organização')
        if 'Nome do Subinventário' in df_grupo.columns:
            colunas_disponiveis.append('Nome do Subinventário')
        
        return df_grupo[colunas_disponiveis].to_dict('records')
    
    def get_estoque_stage(self) -> Dict[str, Any]:
        """Retorna materiais classificados como STAGE."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return {'itens': [], 'resumo_grupos': [], 'total': 0, 'quantidade_itens': 0}
        
        df = self.df_estoque.copy()
        
        if 'Nome do Subinventário' not in df.columns or 'Nome do Item' not in df.columns:
            return {'itens': [], 'resumo_grupos': [], 'total': 0, 'quantidade_itens': 0}
        
        df_stage = df[df['Nome do Subinventário'] == 'STAGE'].copy()
        
        if len(df_stage) == 0:
            return {'itens': [], 'resumo_grupos': [], 'total': 0, 'quantidade_itens': 0}
        
        df_stage['Grupo'] = df_stage['Nome do Item'].apply(self._classificar_grupo_estoque)
        
        resumo = df_stage.groupby('Grupo').agg({
            'Custo Total': 'sum',
            'Nome do Item': 'count'
        }).reset_index()
        resumo.columns = ['Grupo', 'Valor Total', 'Quantidade']
        
        return {
            'itens': df_stage.to_dict('records'),
            'resumo_grupos': resumo.to_dict('records'),
            'total': df_stage['Custo Total'].sum(),
            'quantidade_itens': len(df_stage)
        }
    
    def get_estoque_por_organizacao(self) -> Dict[str, Any]:
        """Retorna valor de estoque por organização/empresa."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return {'dados': [], 'total': 0, 'organizacoes': []}
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        if 'Organização' not in df.columns or 'Custo Total' not in df.columns:
            return {'dados': [], 'total': 0, 'organizacoes': []}
        
        resumo = df.groupby(['Organização', 'Nome da Organização']).agg({
            'Custo Total': 'sum',
            'Nome do Item': 'count'
        }).reset_index()
        resumo.columns = ['Organização', 'Nome Organização', 'Valor Total', 'Quantidade Itens']
        resumo = resumo.sort_values('Valor Total', ascending=False)
        
        return {
            'dados': resumo.to_dict('records'),
            'total': resumo['Valor Total'].sum(),
            'organizacoes': resumo['Organização'].tolist()
        }
    
    def get_curva_abc_estoque(self, organizacao: str = None) -> Dict[str, Any]:
        """Calcula Curva ABC do estoque por valor."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return {'dados': [], 'classe_a': {'qtd': 0, 'valor': 0}, 'classe_b': {'qtd': 0, 'valor': 0}, 'classe_c': {'qtd': 0, 'valor': 0}, 'resumo_grupo_classe': []}
        
        if 'Nome do Item' not in self.df_estoque.columns or 'Custo Total' not in self.df_estoque.columns:
            return {'dados': [], 'classe_a': {'qtd': 0, 'valor': 0}, 'classe_b': {'qtd': 0, 'valor': 0}, 'classe_c': {'qtd': 0, 'valor': 0}, 'resumo_grupo_classe': []}
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        # Filtrar por organização se especificado
        if organizacao and 'Organização' in df.columns:
            df = df[df['Organização'] == organizacao]
        
        # Classificar grupos
        df['Grupo'] = df['Nome do Item'].apply(self._classificar_grupo_estoque)
        
        # Calcular ABC por item
        df_sorted = df.sort_values('Custo Total', ascending=False).copy()
        df_sorted['Acumulado'] = df_sorted['Custo Total'].cumsum()
        df_sorted['Percentual'] = (df_sorted['Acumulado'] / df_sorted['Custo Total'].sum()) * 100
        
        df_sorted['Classificação ABC'] = df_sorted['Percentual'].apply(
            lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C')
        )
        
        # Resumo por classe
        classe_a = df_sorted[df_sorted['Classificação ABC'] == 'A']
        classe_b = df_sorted[df_sorted['Classificação ABC'] == 'B']
        classe_c = df_sorted[df_sorted['Classificação ABC'] == 'C']
        
        # Resumo por grupo e classe
        resumo_grupo_classe = df_sorted.groupby(['Grupo', 'Classificação ABC']).agg({
            'Custo Total': 'sum',
            'Nome do Item': 'count'
        }).reset_index()
        
        # Calcular % de representatividade no estoque total
        total_estoque = df_sorted['Custo Total'].sum()
        df_sorted['% Representatividade'] = (df_sorted['Custo Total'] / total_estoque) * 100
        
        # Preparar listas detalhadas para exibição em tabela
        colunas_abc = ['Nome do Item', 'Descrição do Item', 'Quantidade', 'Custo Total', 'Percentual']
        
        itens_classe_a = classe_a[colunas_abc].copy()
        itens_classe_a['% Acumulado'] = itens_classe_a['Percentual']
        itens_classe_a['% Representatividade'] = (itens_classe_a['Custo Total'] / total_estoque) * 100
        
        itens_classe_b = classe_b[colunas_abc].copy()
        itens_classe_b['% Acumulado'] = itens_classe_b['Percentual']
        itens_classe_b['% Representatividade'] = (itens_classe_b['Custo Total'] / total_estoque) * 100
        
        itens_classe_c = classe_c[colunas_abc].copy()
        itens_classe_c['% Acumulado'] = itens_classe_c['Percentual']
        itens_classe_c['% Representatividade'] = (itens_classe_c['Custo Total'] / total_estoque) * 100
        
        return {
            'dados': df_sorted[['Nome do Item', 'Descrição do Item', 'Grupo', 'Custo Total', 
                                'Percentual', 'Classificação ABC']].to_dict('records'),
            'classe_a': {'qtd': len(classe_a), 'valor': classe_a['Custo Total'].sum()},
            'classe_b': {'qtd': len(classe_b), 'valor': classe_b['Custo Total'].sum()},
            'classe_c': {'qtd': len(classe_c), 'valor': classe_c['Custo Total'].sum()},
            'itens_classe_a': itens_classe_a.to_dict('records'),
            'itens_classe_b': itens_classe_b.to_dict('records'),
            'itens_classe_c': itens_classe_c.to_dict('records'),
            'resumo_grupo_classe': resumo_grupo_classe.to_dict('records')
        }
    
    def get_top10_estoque(self, organizacao: str = None) -> List[Dict]:
        """Retorna os 10 itens mais caros do estoque."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return []
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        # Filtrar por organização se especificado
        if organizacao and 'Organização' in df.columns:
            df = df[df['Organização'] == organizacao]
        
        # Classificar grupos
        df['Grupo'] = df['Nome do Item'].apply(self._classificar_grupo_estoque)
        
        # Top 10 por valor
        df_top10 = df.nlargest(10, 'Custo Total')
        
        colunas = ['Nome do Item', 'Descrição do Item', 'Grupo', 'Quantidade', 
                   'Custo Unitário', 'Custo Total', 'Organização', 'Nome do Subinventário']
        colunas_disponiveis = [col for col in colunas if col in df_top10.columns]
        
        return df_top10[colunas_disponiveis].to_dict('records')
    
    def get_analise_subinventarios(self, organizacao: str = None) -> Dict[str, Any]:
        """Análise detalhada por subinventário."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return {'dados': [], 'total': 0}
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        # Filtrar por organização se especificado
        if organizacao and 'Organização' in df.columns:
            df = df[df['Organização'] == organizacao]
        
        if 'Nome do Subinventário' not in df.columns:
            return {'dados': [], 'total': 0}
        
        # Agrupar por subinventário
        resumo = df.groupby('Nome do Subinventário').agg({
            'Custo Total': 'sum',
            'Nome do Item': 'count',
            'Quantidade': 'sum'
        }).reset_index()
        resumo.columns = ['Subinventário', 'Valor Total', 'Qtd SKUs', 'Qtd Total Itens']
        resumo = resumo.sort_values('Valor Total', ascending=False)
        
        total = resumo['Valor Total'].sum()
        resumo['% do Total'] = (resumo['Valor Total'] / total * 100).round(2)
        
        return {
            'dados': resumo.to_dict('records'),
            'total': total
        }
    
    def get_kpis_estoque(self, organizacao: str = None) -> Dict[str, Any]:
        """KPIs detalhados do estoque."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return {
                'valor_total': 0, 'qtd_skus': 0, 'valor_medio': 0,
                'qtd_total_itens': 0, 'maior_item_valor': 0, 'maior_item_nome': '',
                'itens_zerados': 0, 'concentracao_top10': 0
            }
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        # Filtrar por organização se especificado
        if organizacao and 'Organização' in df.columns:
            df = df[df['Organização'] == organizacao]
        
        valor_total = df['Custo Total'].sum()
        qtd_skus = df['Nome do Item'].nunique()
        valor_medio = valor_total / qtd_skus if qtd_skus > 0 else 0
        qtd_total_itens = df['Quantidade'].sum() if 'Quantidade' in df.columns else 0
        
        # Maior item
        if len(df) > 0:
            idx_maior = df['Custo Total'].idxmax()
            maior_item_valor = df.loc[idx_maior, 'Custo Total']
            maior_item_nome = df.loc[idx_maior, 'Nome do Item'] if 'Nome do Item' in df.columns else ''
        else:
            maior_item_valor = 0
            maior_item_nome = ''
        
        # Itens zerados (quantidade = 0 mas com valor)
        itens_zerados = 0
        if 'Quantidade' in df.columns:
            itens_zerados = len(df[(df['Quantidade'] == 0) & (df['Custo Total'] > 0)])
        
        # Concentração top 10 (% do valor nos 10 maiores itens)
        top10_valor = df.nlargest(10, 'Custo Total')['Custo Total'].sum()
        concentracao_top10 = (top10_valor / valor_total * 100) if valor_total > 0 else 0
        
        return {
            'valor_total': valor_total,
            'qtd_skus': qtd_skus,
            'valor_medio': valor_medio,
            'qtd_total_itens': qtd_total_itens,
            'maior_item_valor': maior_item_valor,
            'maior_item_nome': str(maior_item_nome)[:30],
            'itens_zerados': itens_zerados,
            'concentracao_top10': concentracao_top10
        }
    
    def get_comparativo_organizacoes(self) -> List[Dict]:
        """Tabela comparativa entre organizações."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return []
        
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        if 'Organização' not in df.columns:
            return []
        
        # Agrupar por organização
        resumo = df.groupby(['Organização', 'Nome da Organização']).agg({
            'Custo Total': 'sum',
            'Nome do Item': 'nunique',
            'Quantidade': 'sum'
        }).reset_index()
        resumo.columns = ['Código', 'Organização', 'Valor Total', 'Qtd SKUs', 'Qtd Itens']
        resumo = resumo.sort_values('Valor Total', ascending=False)
        
        total = resumo['Valor Total'].sum()
        resumo['% do Total'] = (resumo['Valor Total'] / total * 100).round(2)
        resumo['Valor Médio/SKU'] = (resumo['Valor Total'] / resumo['Qtd SKUs']).round(2)
        
        return resumo.to_dict('records')
    
    # ═══════════════════════════════════════════════════════════════════════════════
    # MÉTODOS DE ORDENS DE COMPRA
    # ═══════════════════════════════════════════════════════════════════════════════
    
    def load_ordens_compra_file(self, filepath: str) -> bool:
        """Carrega arquivo de ordens de compra (Relatório de Distribuição por Conta)."""
        try:
            if filepath.endswith('.csv'):
                self.df_ordens_compra = pd.read_csv(filepath, encoding='utf-8')
            else:
                self.df_ordens_compra = pd.read_excel(filepath)
            
            # Garantir que colunas numéricas sejam numéricas
            for col in ['Crédito', 'Débito', 'Quantidade']:
                if col in self.df_ordens_compra.columns:
                    self.df_ordens_compra[col] = pd.to_numeric(self.df_ordens_compra[col], errors='coerce').fillna(0)
            
            # Converter data
            if 'Data da Transação' in self.df_ordens_compra.columns:
                self.df_ordens_compra['Data da Transação'] = pd.to_datetime(self.df_ordens_compra['Data da Transação'], errors='coerce')
            
            print(f"Ordens de compra carregadas: {len(self.df_ordens_compra)} linhas")
            return True
            
        except Exception as e:
            print(f"Erro ao carregar arquivo de ordens: {e}")
            traceback.print_exc()
            return False
    
    def get_ordens_compra_analise(self, organizacao: str = None) -> Dict[str, Any]:
        """Retorna análise completa das ordens de compra."""
        if self.df_ordens_compra is None or len(self.df_ordens_compra) == 0:
            return {}
        
        # Filtrar apenas "Ordem de compra"
        df = self.df_ordens_compra[self.df_ordens_compra['Tipo de Referência'] == 'Ordem de compra'].copy()
        
        if len(df) == 0:
            return {}
        
        # Filtrar por organização se especificado
        if organizacao and 'Organização de Estoque' in df.columns:
            df = df[df['Organização de Estoque'] == organizacao]
        
        # Calcular valor - usar apenas Débito (o Crédito é a contrapartida contábil)
        df['Valor'] = df['Débito'].fillna(0).abs()
        
        # KPIs principais
        valor_total = df['Valor'].sum()
        qtd_pedidos = df['Documento de Referência'].nunique()
        qtd_itens = df['Item'].nunique()
        qtd_transacoes = len(df)
        valor_medio_pedido = valor_total / qtd_pedidos if qtd_pedidos > 0 else 0
        
        # Por organização
        org_resumo = df.groupby('Organização de Estoque').agg({
            'Valor': 'sum',
            'Documento de Referência': 'nunique',
            'Item': 'nunique'
        }).reset_index()
        org_resumo.columns = ['Organização', 'Valor Total', 'Qtd Pedidos', 'Qtd Itens']
        org_resumo = org_resumo.sort_values('Valor Total', ascending=False)
        
        # Top 10 itens mais caros
        top10_itens = df.groupby(['Item', 'Descrição do Item']).agg({
            'Valor': 'sum',
            'Quantidade': 'sum',
            'Organização de Estoque': 'first'
        }).reset_index()
        top10_itens = top10_itens.nlargest(10, 'Valor')
        top10_itens.columns = ['Código', 'Descrição', 'Valor Total', 'Quantidade', 'Organização']
        
        # Por tipo de transação
        por_tipo = df.groupby('Tipo de Transação').agg({
            'Valor': 'sum',
            'Documento de Referência': 'nunique'
        }).reset_index()
        por_tipo.columns = ['Tipo', 'Valor Total', 'Qtd Pedidos']
        por_tipo = por_tipo.sort_values('Valor Total', ascending=False)
        
        # Por mês (tendência)
        if 'Data da Transação' in df.columns:
            df['Mes'] = df['Data da Transação'].dt.to_period('M')
            por_mes = df.groupby('Mes').agg({
                'Valor': 'sum',
                'Documento de Referência': 'nunique'
            }).reset_index()
            por_mes['Mes'] = por_mes['Mes'].astype(str)
            por_mes.columns = ['Mês', 'Valor', 'Qtd Pedidos']
            tendencia = por_mes.tail(12).to_dict('records')
        else:
            tendencia = []
        
        # Lista de pedidos únicos com detalhes
        pedidos = df.groupby('Documento de Referência').agg({
            'Valor': 'sum',
            'Item': 'nunique',
            'Quantidade': 'sum',
            'Organização de Estoque': 'first',
            'Data da Transação': 'first'
        }).reset_index()
        pedidos.columns = ['Nº Pedido', 'Valor Total', 'Qtd Itens', 'Qtd Total', 'Organização', 'Data']
        # Filtrar pedidos com valor maior que zero
        pedidos = pedidos[pedidos['Valor Total'] > 0]
        pedidos = pedidos.sort_values('Valor Total', ascending=False)
        
        # Top 10 pedidos com detalhes dos itens
        top10_pedidos_list = pedidos.head(10).to_dict('records')
        top10_pedidos = []
        for pedido in top10_pedidos_list:
            num_pedido = pedido['Nº Pedido']
            # Buscar itens deste pedido
            itens_pedido = df[df['Documento de Referência'] == num_pedido][['Item', 'Descrição do Item', 'Quantidade', 'Valor']].copy()
            itens_pedido.columns = ['Código', 'Descrição', 'Quantidade', 'Valor']
            itens_pedido = itens_pedido.sort_values('Valor', ascending=False)
            pedido['itens'] = itens_pedido.to_dict('records')
            top10_pedidos.append(pedido)
        
        # Análise de itens repetidos no mês (itens comprados mais de uma vez no mesmo mês)
        itens_repetidos = []
        if 'Data da Transação' in df.columns:
            df['Mes_Ano'] = df['Data da Transação'].dt.strftime('%Y-%m')
            # Contar quantas vezes cada item aparece em cada mês (em pedidos diferentes)
            itens_por_mes = df.groupby(['Item', 'Descrição do Item', 'Mes_Ano']).agg({
                'Documento de Referência': 'nunique',  # Quantidade de pedidos diferentes
                'Quantidade': 'sum',
                'Valor': 'sum'
            }).reset_index()
            itens_por_mes.columns = ['Código', 'Descrição', 'Mês', 'Qtd Pedidos', 'Qtd Total', 'Valor Total']
            # Filtrar apenas itens que aparecem em mais de 1 pedido no mesmo mês
            itens_repetidos_df = itens_por_mes[itens_por_mes['Qtd Pedidos'] > 1].copy()
            itens_repetidos_df = itens_repetidos_df.sort_values(['Mês', 'Qtd Pedidos'], ascending=[False, False])
            # Pegar top 20 itens mais repetidos
            itens_repetidos = itens_repetidos_df.head(20).to_dict('records')
        
        # Organizações únicas para filtro
        organizacoes = sorted(df['Organização de Estoque'].dropna().unique().tolist())
        
        return {
            'valor_total': valor_total,
            'qtd_pedidos': qtd_pedidos,
            'qtd_itens': qtd_itens,
            'qtd_transacoes': qtd_transacoes,
            'valor_medio_pedido': valor_medio_pedido,
            'por_organizacao': org_resumo.to_dict('records'),
            'top10_itens': top10_itens.to_dict('records'),
            'por_tipo': por_tipo.to_dict('records'),
            'tendencia': tendencia,
            'top10_pedidos': top10_pedidos,
            'itens_repetidos': itens_repetidos,
            'organizacoes': organizacoes
        }

    def _classificar_grupo_item(self, codigo_item: str) -> str:
        """Classifica item em grupo baseado no código."""
        if not codigo_item or pd.isna(codigo_item):
            return None
        
        codigo = str(codigo_item).strip()
        
        # Mapeamento por prefixo do código
        grupos = {
            '01.01.01': 'Combustíveis',
            '01.01.02': 'Lubrificantes',
            '01.01.03': 'Peças de Frota',
            '01.01.04': 'Componentes e Implementos',
            '01.01.05': 'Componentes e Implementos',
            '01.01.06': 'Materiais Administrativos',
            '01.01.07': 'Materiais Administrativos',
            '01.01.08': 'EPIs/Uniformes',
            '01.01.09': 'Insumos',
            '01.01.10': 'Componentes e Implementos',
            '01.01.11': 'Materiais Administrativos',
            '01.01.12': 'Insumos',
            '01.01.13': 'Insumos',
            '01.01.14': 'Insumos',
            '01.01.15': 'Insumos',
            '01.01.16': 'Insumos',
            '01.01.17': 'Insumos',
            '01.01.18': 'Insumos',
            '01.01.19': 'Insumos',
            '01.01.20': 'Insumos',
            '01.01.21': 'Materiais Administrativos',
            '01.01.22': 'Materiais Administrativos',
            '01.01.23': 'Materiais Administrativos',
        }
        
        for prefixo, grupo in grupos.items():
            if codigo.startswith(prefixo):
                return grupo
        
        # Retorna None para itens fora do mapeamento (serão filtrados)
        return None

    def get_movimentacoes_analise(self, organizacao: str = None) -> Dict[str, Any]:
        """Retorna análise de movimentações (saídas de estoque)."""
        if self.df_ordens_compra is None or len(self.df_ordens_compra) == 0:
            return {}
        
        df = self.df_ordens_compra.copy()
        
        # === TRANSFERÊNCIAS PARA DESPESA ===
        df_transf = df[
            (df['Tipo de Referência'] == 'Ordem de transferência') & 
            (df['Tipo de Fluxo'] == 'Emissão de Ordem de Transferência Intraorganizações para Despesa')
        ].copy()
        
        # === BAIXAS POR ORDEM DE SERVIÇO ===
        df_os = df[
            (df['Tipo de Referência'] == 'Ordem de serviço') & 
            (df['Tipo de Transação'] == 'Emissão do Material em Processo')
        ].copy()
        
        # === CONSUMO DE COMBUSTÍVEL ===
        df_comb = df[
            df['Tipo de Transação'] == 'GTFrota - Consumo de Combustivel'
        ].copy()
        
        # Filtrar por organização se especificado
        if organizacao:
            if 'Organização de Estoque' in df_transf.columns:
                df_transf = df_transf[df_transf['Organização de Estoque'] == organizacao]
            if 'Organização de Estoque' in df_os.columns:
                df_os = df_os[df_os['Organização de Estoque'] == organizacao]
            if 'Organização de Estoque' in df_comb.columns:
                df_comb = df_comb[df_comb['Organização de Estoque'] == organizacao]
        
        # Calcular valores - usar apenas Débito (valor de saída)
        for dframe in [df_transf, df_os, df_comb]:
            if len(dframe) > 0:
                dframe['Valor'] = dframe['Débito'].fillna(0).abs()
                dframe['Grupo'] = dframe['Item'].apply(self._classificar_grupo_item)
        
        # Filtrar itens sem grupo válido (None)
        if len(df_transf) > 0:
            df_transf = df_transf[df_transf['Grupo'].notna()].copy()
        if len(df_os) > 0:
            df_os = df_os[df_os['Grupo'].notna()].copy()
        if len(df_comb) > 0:
            df_comb = df_comb[df_comb['Grupo'].notna()].copy()
        
        # === ANÁLISE TRANSFERÊNCIAS ===
        transf_total = df_transf['Valor'].sum() if len(df_transf) > 0 else 0
        transf_qtd = len(df_transf)
        transf_itens = df_transf['Item'].nunique() if len(df_transf) > 0 else 0
        
        # Por grupo - Transferências
        if len(df_transf) > 0:
            transf_por_grupo = df_transf.groupby('Grupo').agg({
                'Valor': 'sum',
                'Item': 'nunique',
                'Quantidade': 'sum'
            }).reset_index()
            transf_por_grupo.columns = ['Grupo', 'Valor Total', 'Qtd Itens', 'Quantidade']
            transf_por_grupo = transf_por_grupo.sort_values('Valor Total', ascending=False)
            transf_por_grupo_list = transf_por_grupo.to_dict('records')
            
            # Top 10 transferências
            transf_top10 = df_transf.groupby(['Item', 'Descrição do Item']).agg({
                'Valor': 'sum',
                'Quantidade': 'sum',
                'Organização de Estoque': 'first'
            }).reset_index().nlargest(10, 'Valor')
            transf_top10.columns = ['Código', 'Descrição', 'Valor Total', 'Quantidade', 'Organização']
            transf_top10_list = transf_top10.to_dict('records')
        else:
            transf_por_grupo_list = []
            transf_top10_list = []
        
        # === ANÁLISE ORDEM DE SERVIÇO ===
        os_total = df_os['Valor'].sum() if len(df_os) > 0 else 0
        os_qtd = len(df_os)
        os_itens = df_os['Item'].nunique() if len(df_os) > 0 else 0
        
        # Por grupo - OS
        if len(df_os) > 0:
            os_por_grupo = df_os.groupby('Grupo').agg({
                'Valor': 'sum',
                'Item': 'nunique',
                'Quantidade': 'sum'
            }).reset_index()
            os_por_grupo.columns = ['Grupo', 'Valor Total', 'Qtd Itens', 'Quantidade']
            os_por_grupo = os_por_grupo.sort_values('Valor Total', ascending=False)
            os_por_grupo_list = os_por_grupo.to_dict('records')
            
            # Top 10 OS
            os_top10 = df_os.groupby(['Item', 'Descrição do Item']).agg({
                'Valor': 'sum',
                'Quantidade': 'sum',
                'Organização de Estoque': 'first'
            }).reset_index().nlargest(10, 'Valor')
            os_top10.columns = ['Código', 'Descrição', 'Valor Total', 'Quantidade', 'Organização']
            os_top10_list = os_top10.to_dict('records')
        else:
            os_por_grupo_list = []
            os_top10_list = []
        
        # === ANÁLISE CONSUMO DE COMBUSTÍVEL ===
        comb_total = df_comb['Valor'].sum() if len(df_comb) > 0 else 0
        comb_qtd = len(df_comb)
        comb_itens = df_comb['Item'].nunique() if len(df_comb) > 0 else 0
        
        # Por grupo - Combustível
        if len(df_comb) > 0:
            comb_por_grupo = df_comb.groupby('Grupo').agg({
                'Valor': 'sum',
                'Item': 'nunique',
                'Quantidade': 'sum'
            }).reset_index()
            comb_por_grupo.columns = ['Grupo', 'Valor Total', 'Qtd Itens', 'Quantidade']
            comb_por_grupo = comb_por_grupo.sort_values('Valor Total', ascending=False)
            comb_por_grupo_list = comb_por_grupo.to_dict('records')
            
            # Top 10 Combustível
            comb_top10 = df_comb.groupby(['Item', 'Descrição do Item']).agg({
                'Valor': 'sum',
                'Quantidade': 'sum',
                'Organização de Estoque': 'first'
            }).reset_index().nlargest(10, 'Valor')
            comb_top10.columns = ['Código', 'Descrição', 'Valor Total', 'Quantidade', 'Organização']
            comb_top10_list = comb_top10.to_dict('records')
        else:
            comb_por_grupo_list = []
            comb_top10_list = []
        
        # Organizações únicas
        org_list = []
        if 'Organização de Estoque' in df.columns:
            org_list = sorted(df['Organização de Estoque'].dropna().unique().tolist())
        
        # Detalhamento de materiais por grupo
        materiais_transf_dict = {}
        if len(df_transf) > 0:
            for grupo in df_transf['Grupo'].unique():
                df_grupo = df_transf[df_transf['Grupo'] == grupo]
                materiais = df_grupo.groupby(['Item', 'Descrição do Item']).agg({
                    'Quantidade': 'sum',
                    'Valor': 'sum'
                }).reset_index()
                materiais.columns = ['Codigo', 'Material', 'Quantidade', 'Valor']
                materiais = materiais.sort_values('Valor', ascending=False).head(20)
                materiais_transf_dict[grupo] = materiais.to_dict('records')
        
        materiais_os_dict = {}
        if len(df_os) > 0:
            for grupo in df_os['Grupo'].unique():
                df_grupo = df_os[df_os['Grupo'] == grupo]
                materiais = df_grupo.groupby(['Item', 'Descrição do Item']).agg({
                    'Quantidade': 'sum',
                    'Valor': 'sum'
                }).reset_index()
                materiais.columns = ['Codigo', 'Material', 'Quantidade', 'Valor']
                materiais = materiais.sort_values('Valor', ascending=False).head(20)
                materiais_os_dict[grupo] = materiais.to_dict('records')
        
        return {
            # Transferências
            'transf_total': transf_total,
            'transf_qtd': transf_qtd,
            'transf_itens': transf_itens,
            'transf_por_grupo': transf_por_grupo_list,
            'transf_top10': transf_top10_list,
            # Ordem de Serviço
            'os_total': os_total,
            'os_qtd': os_qtd,
            'os_itens': os_itens,
            'os_por_grupo': os_por_grupo_list,
            'os_top10': os_top10_list,
            # Combustível
            'comb_total': comb_total,
            'comb_qtd': comb_qtd,
            'comb_itens': comb_itens,
            'comb_por_grupo': comb_por_grupo_list,
            'comb_top10': comb_top10_list,
            # Geral
            'total_geral': transf_total + os_total + comb_total,
            'organizacoes': org_list,
            # Materiais detalhados por grupo
            'materiais_por_grupo_transferencia': materiais_transf_dict,
            'materiais_por_grupo_os': materiais_os_dict
        }

    def get_kpis(self, organizacao: str = None) -> Dict[str, Any]:
        """Calcula KPIs principais do dashboard."""
        if self.df_pagamentos is None or len(self.df_pagamentos) == 0:
            return {
                'total_pago': 0,
                'fornecedores_ativos': 0,
                'ticket_medio': 0,
                'maior_pagamento': 0,
                'maior_fornecedor': 'N/A',
                'variacao_mensal': 0,
                'valor_estoque': 0,
                'itens_criticos': 0,
                'ordens_abertas': 0,
                'ordens_atrasadas': 0
            }
        
        df = self.df_pagamentos.copy()
        
        # Verificar se as colunas necessárias existem
        if 'Valor Pago' not in df.columns or 'Fornecedor' not in df.columns or 'Data do Pagamento' not in df.columns:
            return {
                'total_pago': 0,
                'fornecedores_ativos': 0,
                'ticket_medio': 0,
                'maior_pagamento': 0,
                'maior_fornecedor': 'N/A',
                'variacao_mensal': 0,
                'valor_estoque': 0,
                'itens_criticos': 0,
                'ordens_abertas': 0,
                'ordens_atrasadas': 0
            }
        
        # KPIs de pagamentos
        total_pago = df['Valor Pago'].sum()
        fornecedores_ativos = df['Fornecedor'].nunique()
        ticket_medio = df['Valor Pago'].mean()
        maior_pagamento = df['Valor Pago'].max()
        maior_fornecedor = df.loc[df['Valor Pago'].idxmax(), 'Fornecedor'] if len(df) > 0 else 'N/A'
        
        # Variação mensal
        df['Mes'] = pd.to_datetime(df['Data do Pagamento']).dt.to_period('M')
        gastos_mensais = df.groupby('Mes')['Valor Pago'].sum()
        
        if len(gastos_mensais) >= 2:
            variacao_mensal = ((gastos_mensais.iloc[-1] - gastos_mensais.iloc[-2]) / gastos_mensais.iloc[-2]) * 100
        else:
            variacao_mensal = 0
        
        # KPIs de estoque (com filtro de organização)
        if self.df_estoque is not None and len(self.df_estoque) > 0:
            df_est = self.df_estoque.copy()
            
            # Filtrar subinventários excluídos
            subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
            if 'Nome do Subinventário' in df_est.columns:
                df_est = df_est[~df_est['Nome do Subinventário'].isin(subinv_excluidos)]
            
            # Filtrar por organização se especificado
            if organizacao and 'Organização' in df_est.columns:
                df_est = df_est[df_est['Organização'] == organizacao]
            
            if 'Custo Total' in df_est.columns:
                valor_estoque = df_est['Custo Total'].sum()
            else:
                valor_estoque = 0
            itens_criticos = 0  # Será implementado quando tivermos estoque mín/máx
        else:
            valor_estoque = 0
            itens_criticos = 0
        
        # KPIs de ordens
        ordens_abertas = len(self.df_ordens[self.df_ordens['Status'].isin(['Emitida', 'Aprovada', 'Em Trânsito'])]) if self.df_ordens is not None and len(self.df_ordens) > 0 else 0
        ordens_atrasadas = len(self.df_ordens[self.df_ordens['Dias em Atraso'] > 0]) if self.df_ordens is not None and len(self.df_ordens) > 0 else 0
        
        return {
            'total_pago': total_pago,
            'fornecedores_ativos': fornecedores_ativos,
            'ticket_medio': ticket_medio,
            'maior_pagamento': maior_pagamento,
            'maior_fornecedor': maior_fornecedor,
            'variacao_mensal': variacao_mensal,
            'valor_estoque': valor_estoque,
            'itens_criticos': itens_criticos,
            'ordens_abertas': ordens_abertas,
            'ordens_atrasadas': ordens_atrasadas,
            'total_notas': len(df),
            'novos_fornecedores': 8,
        }
    
    def get_curva_abc(self) -> Dict[str, Any]:
        """Calcula a Curva ABC dos fornecedores."""
        if self.df_pagamentos is None:
            return {}
        
        df = self.df_pagamentos.groupby('Fornecedor')['Valor Pago'].sum().reset_index()
        df = df.sort_values('Valor Pago', ascending=False)
        df['Acumulado'] = df['Valor Pago'].cumsum()
        total = df['Valor Pago'].sum()
        df['Percentual Acumulado'] = (df['Acumulado'] / total) * 100
        df['Percentual Individual'] = (df['Valor Pago'] / total) * 100
        
        df['Classificação'] = df['Percentual Acumulado'].apply(
            lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C')
        )
        
        classe_a = df[df['Classificação'] == 'A']
        classe_b = df[df['Classificação'] == 'B']
        classe_c = df[df['Classificação'] == 'C']
        
        # Tabelas detalhadas por classe
        tabela_classe_a = classe_a[['Fornecedor', 'Valor Pago', 'Percentual Individual', 'Percentual Acumulado', 'Classificação']].to_dict('records')
        tabela_classe_b = classe_b[['Fornecedor', 'Valor Pago', 'Percentual Individual', 'Percentual Acumulado', 'Classificação']].to_dict('records')
        tabela_classe_c = classe_c[['Fornecedor', 'Valor Pago', 'Percentual Individual', 'Percentual Acumulado', 'Classificação']].to_dict('records')
        
        return {
            'dados': df.to_dict('records'),
            'tabela_completa': df[['Fornecedor', 'Valor Pago', 'Percentual Individual', 'Percentual Acumulado', 'Classificação']].to_dict('records'),
            'tabela_classe_a': tabela_classe_a,
            'tabela_classe_b': tabela_classe_b,
            'tabela_classe_c': tabela_classe_c,
            'classe_a': {'qtd': len(classe_a), 'valor': classe_a['Valor Pago'].sum(), 'percentual': (classe_a['Valor Pago'].sum() / total * 100)},
            'classe_b': {'qtd': len(classe_b), 'valor': classe_b['Valor Pago'].sum(), 'percentual': (classe_b['Valor Pago'].sum() / total * 100)},
            'classe_c': {'qtd': len(classe_c), 'valor': classe_c['Valor Pago'].sum(), 'percentual': (classe_c['Valor Pago'].sum() / total * 100)},
            'total_fornecedores': len(df),
            'valor_total': total
        }
    
    def get_ranking_fornecedores(self, top_n: int = 20) -> List[Dict]:
        """Retorna ranking dos principais fornecedores."""
        if self.df_pagamentos is None:
            return []
        
        df = self.df_pagamentos.groupby(['Fornecedor', 'Nr do Fornecedor']).agg({
            'Valor Pago': 'sum',
            'Número da NFF': 'count'
        }).reset_index()
        
        df.columns = ['Fornecedor', 'CNPJ', 'Total Pago', 'Qtd NFs']
        df = df.sort_values('Total Pago', ascending=False).head(top_n)
        df['Ticket Médio'] = df['Total Pago'] / df['Qtd NFs']
        df['% Total'] = (df['Total Pago'] / df['Total Pago'].sum()) * 100
        
        return df.to_dict('records')
    
    def get_evolucao_mensal(self) -> Dict[str, Any]:
        """Retorna dados de evolução mensal dos gastos."""
        if self.df_pagamentos is None:
            return {}
        
        df = self.df_pagamentos.copy()
        df['Mes'] = pd.to_datetime(df['Data do Pagamento']).dt.strftime('%Y-%m')
        
        evolucao = df.groupby('Mes')['Valor Pago'].sum().reset_index()
        evolucao.columns = ['Mes', 'Valor']
        
        # Top 5 fornecedores evolução
        top_fornecedores = df.groupby('Fornecedor')['Valor Pago'].sum().nlargest(5).index.tolist()
        
        evolucao_fornecedores = {}
        for forn in top_fornecedores:
            dados_forn = df[df['Fornecedor'] == forn].groupby('Mes')['Valor Pago'].sum().reset_index()
            evolucao_fornecedores[forn] = dados_forn.to_dict('records')
        
        return {
            'total': evolucao.to_dict('records'),
            'por_fornecedor': evolucao_fornecedores,
            'meses': evolucao['Mes'].tolist()
        }
    
    def get_analise_risco(self) -> Dict[str, Any]:
        """Analisa concentração e risco de dependência de fornecedores."""
        if self.df_pagamentos is None or len(self.df_pagamentos) == 0:
            return {'dados': [], 'alertas': [], 'concentracao_top5': 0}
        
        # Verificar se as colunas necessárias existem
        if 'Fornecedor' not in self.df_pagamentos.columns or 'Valor Pago' not in self.df_pagamentos.columns:
            return {'dados': [], 'alertas': [], 'concentracao_top5': 0}
        
        df = self.df_pagamentos.groupby('Fornecedor')['Valor Pago'].sum().reset_index()
        total = df['Valor Pago'].sum()
        
        if total == 0:
            return {'dados': [], 'alertas': [], 'concentracao_top5': 0}
        
        df['Percentual'] = (df['Valor Pago'] / total) * 100
        df = df.sort_values('Percentual', ascending=False)
        
        def classificar_risco(pct):
            if pct > 30:
                return 'CRÍTICO'
            elif pct > 15:
                return 'ALTO'
            elif pct > 8:
                return 'MÉDIO'
            else:
                return 'BAIXO'
        
        df['Risco'] = df['Percentual'].apply(classificar_risco)
        
        alertas = []
        criticos = df[df['Risco'] == 'CRÍTICO']
        for _, row in criticos.iterrows():
            alertas.append({
                'tipo': 'CRÍTICO',
                'mensagem': f"Fornecedor {row['Fornecedor']} representa {row['Percentual']:.1f}% do gasto total",
                'icone': '🔴'
            })
        
        return {
            'dados': df.to_dict('records'),
            'alertas': alertas,
            'concentracao_top5': df.head(5)['Percentual'].sum()
        }
    
    def get_tempo_entrega_fornecedores(self) -> List[Dict]:
        """Calcula tempo médio de entrega por fornecedor."""
        if self.df_ordens is None or len(self.df_ordens) == 0:
            return []
        
        # Verificar se as colunas necessárias existem
        required_cols = ['Fornecedor', 'Prazo Entrega', 'Dias em Atraso', 'Número OC']
        if not all(col in self.df_ordens.columns for col in required_cols):
            return []
        
        df = self.df_ordens.copy()
        
        analise = df.groupby('Fornecedor').agg({
            'Prazo Entrega': 'mean',
            'Dias em Atraso': 'mean',
            'Número OC': 'count'
        }).reset_index()
        
        analise.columns = ['Fornecedor', 'Prazo Médio', 'Atraso Médio', 'Qtd Pedidos']
        analise['Score'] = 100 - (analise['Atraso Médio'] / analise['Prazo Médio'] * 100).clip(0, 100)
        
        return analise.sort_values('Score', ascending=False).to_dict('records')
    
    def get_evolucao_estoque_mensal(self) -> Dict[str, Any]:
        """Retorna evolução mensal do valor em estoque - VALORES REAIS."""
        if self.df_estoque is None or len(self.df_estoque) == 0:
            return {'meses': [], 'valores': []}
        
        from datetime import datetime
        
        # Pegar valor total atual do estoque (REAL)
        df = self.df_estoque.copy()
        
        # Filtrar subinventários excluídos
        subinv_excluidos = ['ATK1', 'ATK2', 'MANU1', 'MANU3']
        if 'Nome do Subinventário' in df.columns:
            df = df[~df['Nome do Subinventário'].isin(subinv_excluidos)]
        
        valor_atual = df['Custo Total'].sum() if 'Custo Total' in df.columns else 0
        
        # Retornar apenas o valor real do mês atual
        # Para histórico completo, seria necessário ter snapshots mensais salvos
        data_atual = datetime.now()
        mes_atual = data_atual.strftime('%Y-%m')
        
        return {
            'meses': [mes_atual],
            'valores': [valor_atual]
        }
    
    def get_alertas(self) -> List[Dict]:
        """Gera lista de alertas do sistema."""
        alertas = []
        
        # Alertas de estoque - desabilitados por enquanto (dados de amostra não têm min/max)
        # Será implementado quando tivermos dados reais com estoque mínimo/máximo
        
        # Alertas de ordens
        if self.df_ordens is not None and len(self.df_ordens) > 0 and 'Dias em Atraso' in self.df_ordens.columns:
            atrasadas = self.df_ordens[self.df_ordens['Dias em Atraso'] > 5]
            for _, ordem in atrasadas.head(5).iterrows():
                alertas.append({
                    'tipo': 'ATENÇÃO',
                    'categoria': 'Ordens',
                    'mensagem': f"OC {ordem['Número OC']} atrasada em {ordem['Dias em Atraso']} dias",
                    'icone': '🟡',
                    'acao': 'Contatar fornecedor'
                })
        
        # Alertas de concentração
        risco = self.get_analise_risco()
        alertas.extend(risco.get('alertas', []))
        
        return alertas[:10]
    
    def get_ficha_fornecedor(self, nome_fornecedor: str) -> Dict[str, Any]:
        """Retorna ficha detalhada de um fornecedor."""
        if self.df_pagamentos is None:
            return {}
        
        df = self.df_pagamentos[self.df_pagamentos['Fornecedor'] == nome_fornecedor]
        
        if df.empty:
            return {}
        
        historico = df.sort_values('Data do Pagamento', ascending=False).head(20).to_dict('records')
        
        return {
            'nome': nome_fornecedor,
            'cnpj': df['Nr do Fornecedor'].iloc[0] if 'Nr do Fornecedor' in df.columns else '',
            'total_pago': df['Valor Pago'].sum(),
            'qtd_notas': len(df),
            'ticket_medio': df['Valor Pago'].mean(),
            'maior_pagamento': df['Valor Pago'].max(),
            'menor_pagamento': df['Valor Pago'].min(),
            'primeiro_pagamento': df['Data do Pagamento'].min(),
            'ultimo_pagamento': df['Data do Pagamento'].max(),
            'historico': historico
        }
    
    def exportar_excel(self, filepath: str) -> bool:
        """Exporta dados para Excel."""
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                if self.df_pagamentos is not None:
                    self.df_pagamentos.to_excel(writer, sheet_name='Pagamentos', index=False)
                if self.df_estoque is not None:
                    self.df_estoque.to_excel(writer, sheet_name='Estoque', index=False)
                if self.df_ordens is not None:
                    self.df_ordens.to_excel(writer, sheet_name='Ordens', index=False)
                if self.df_movimentacoes is not None:
                    self.df_movimentacoes.to_excel(writer, sheet_name='Movimentações', index=False)
            return True
        except Exception as e:
            print(f"Erro ao exportar: {e}")
            return False

    def exportar_pdf_executivo(self, filepath: str, organizacao: str = None) -> bool:
        """
        Exporta relatório executivo profissional em PDF com gráficos.
        """
        try:
            # Configurar documento PDF em paisagem para melhor visualização
            doc = SimpleDocTemplate(
                filepath,
                pagesize=landscape(A4),
                rightMargin=1.5*cm,
                leftMargin=1.5*cm,
                topMargin=1.5*cm,
                bottomMargin=1.5*cm
            )
            
            # Estilos personalizados
            styles = getSampleStyleSheet()
            
            # Estilo do título principal
            title_style = ParagraphStyle(
                'TitleExecutive',
                parent=styles['Heading1'],
                fontSize=24,
                textColor=colors.HexColor('#2d3748'),
                spaceAfter=20,
                alignment=TA_CENTER,
                fontName='Helvetica-Bold'
            )
            
            # Estilo do subtítulo
            subtitle_style = ParagraphStyle(
                'SubtitleExecutive',
                parent=styles['Normal'],
                fontSize=12,
                textColor=colors.HexColor('#718096'),
                spaceAfter=30,
                alignment=TA_CENTER
            )
            
            # Estilo de seção
            section_style = ParagraphStyle(
                'SectionTitle',
                parent=styles['Heading2'],
                fontSize=16,
                textColor=colors.HexColor('#2d3748'),
                spaceBefore=25,
                spaceAfter=15,
                fontName='Helvetica-Bold',
                borderColor=colors.HexColor('#2d3748'),
                borderWidth=1,
                borderPadding=5
            )
            
            # Estilo para texto normal
            normal_style = ParagraphStyle(
                'NormalExecutive',
                parent=styles['Normal'],
                fontSize=10,
                textColor=colors.HexColor('#4a5568'),
                spaceAfter=8
            )
            
            # Estilo para KPIs
            kpi_value_style = ParagraphStyle(
                'KPIValue',
                parent=styles['Normal'],
                fontSize=18,
                textColor=colors.HexColor('#2d3748'),
                fontName='Helvetica-Bold',
                alignment=TA_CENTER
            )
            
            kpi_label_style = ParagraphStyle(
                'KPILabel',
                parent=styles['Normal'],
                fontSize=9,
                textColor=colors.HexColor('#718096'),
                alignment=TA_CENTER
            )
            
            elements = []
            
            # ═══════════════════════════════════════════════════════════════════
            # CAPA DO RELATÓRIO
            # ═══════════════════════════════════════════════════════════════════
            
            elements.append(Spacer(1, 3*cm))
            elements.append(Paragraph("📊 RELATÓRIO", title_style))
            elements.append(Paragraph("Painel Inteligente de Gestão", subtitle_style))
            elements.append(Spacer(1, 1*cm))
            
            # Informações do relatório
            data_geracao = datetime.now().strftime("%d/%m/%Y às %H:%M")
            org_info = f"Organização: {organizacao}" if organizacao else "Todas as Organizações"
            
            info_table_data = [
                ['Data de Geração:', data_geracao],
                ['Escopo:', org_info],
                ['Período:', 'Ano Atual'],
            ]
            
            info_table = Table(info_table_data, colWidths=[4*cm, 8*cm])
            info_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),
                ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#2d3748')),
                ('TEXTCOLOR', (1, 0), (1, -1), colors.HexColor('#4a5568')),
                ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))
            elements.append(info_table)
            elements.append(PageBreak())
            
            # ═══════════════════════════════════════════════════════════════════
            # SUMÁRIO EXECUTIVO - KPIs PRINCIPAIS
            # ═══════════════════════════════════════════════════════════════════
            
            elements.append(Paragraph("📈 SUMÁRIO EXECUTIVO", section_style))
            
            kpis = self.get_kpis(organizacao)
            
            def fmt_currency(val):
                try:
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        return "R$ 0,00"
                    return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except:
                    return "R$ 0,00"
            
            # Tabela de KPIs principais
            kpi_data = [
                [
                    Paragraph(fmt_currency(kpis.get('total_pago', 0)), kpi_value_style),
                    Paragraph(str(kpis.get('fornecedores_ativos', 0)), kpi_value_style),
                    Paragraph(fmt_currency(kpis.get('ticket_medio', 0)), kpi_value_style),
                    Paragraph(str(kpis.get('total_notas', 0)), kpi_value_style),
                ],
                [
                    Paragraph('TOTAL PAGO', kpi_label_style),
                    Paragraph('FORNECEDORES ATIVOS', kpi_label_style),
                    Paragraph('TICKET MÉDIO', kpi_label_style),
                    Paragraph('TOTAL DE NFs', kpi_label_style),
                ]
            ]
            
            kpi_table = Table(kpi_data, colWidths=[6*cm, 6*cm, 6*cm, 6*cm])
            kpi_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f7fafc')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
                ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('TOPPADDING', (0, 0), (-1, -1), 15),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
            ]))
            elements.append(kpi_table)
            elements.append(Spacer(1, 0.5*cm))
            
            # Segunda linha de KPIs
            kpi_data2 = [
                [
                    Paragraph(fmt_currency(kpis.get('valor_estoque', 0)), kpi_value_style),
                    Paragraph(str(kpis.get('itens_criticos', 0)), kpi_value_style),
                    Paragraph(str(kpis.get('ordens_abertas', 0)), kpi_value_style),
                    Paragraph(str(kpis.get('ordens_atrasadas', 0)), kpi_value_style),
                ],
                [
                    Paragraph('VALOR EM ESTOQUE', kpi_label_style),
                    Paragraph('ITENS CRÍTICOS', kpi_label_style),
                    Paragraph('ORDENS ABERTAS', kpi_label_style),
                    Paragraph('ORDENS ATRASADAS', kpi_label_style),
                ]
            ]
            
            kpi_table2 = Table(kpi_data2, colWidths=[6*cm, 6*cm, 6*cm, 6*cm])
            kpi_table2.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f7fafc')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
                ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('TOPPADDING', (0, 0), (-1, -1), 15),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
            ]))
            elements.append(kpi_table2)
            elements.append(Spacer(1, 1*cm))
            
            # ═══════════════════════════════════════════════════════════════════
            # GRÁFICO: TOP 10 FORNECEDORES
            # ═══════════════════════════════════════════════════════════════════
            
            elements.append(Paragraph("🏆 TOP 10 FORNECEDORES", section_style))
            
            ranking = self.get_ranking_fornecedores(top_n=10)
            
            if ranking:
                # Criar gráfico com matplotlib
                fig, ax = plt.subplots(figsize=(12, 5))
                
                fornecedores = [r.get('Fornecedor', '')[:25] for r in ranking[:10]]
                valores = [r.get('Total Pago', 0) for r in ranking[:10]]
                
                # Cores gradient
                cores = ['#2d3748', '#3d4a5c', '#4d5d70', '#5d7084', '#6d8398', 
                         '#7d96ac', '#8da9c0', '#9dbcd4', '#adcfe8', '#bde2fc']
                
                bars = ax.barh(fornecedores[::-1], valores[::-1], color=cores[::-1])
                
                ax.set_xlabel('Valor Total Pago (R$)', fontsize=10, fontweight='bold')
                ax.set_title('Ranking de Fornecedores por Volume de Pagamentos', fontsize=12, fontweight='bold', pad=15)
                
                # Adicionar valores nas barras
                for i, (bar, val) in enumerate(zip(bars, valores[::-1])):
                    ax.text(bar.get_width() + max(valores)*0.01, bar.get_y() + bar.get_height()/2, 
                            fmt_currency(val), va='center', fontsize=8)
                
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                plt.tight_layout()
                
                # Salvar gráfico em buffer
                img_buffer = io.BytesIO()
                plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight', 
                           facecolor='white', edgecolor='none')
                img_buffer.seek(0)
                plt.close(fig)
                
                img = Image(img_buffer, width=24*cm, height=10*cm)
                elements.append(img)
                elements.append(Spacer(1, 0.5*cm))
                
                # Tabela detalhada do ranking
                ranking_header = ['#', 'Fornecedor', 'Total Pago', '% Total', 'Qtd NFs', 'Ticket Médio', 'Classe']
                ranking_data = [ranking_header]
                
                for i, forn in enumerate(ranking[:10], 1):
                    pct = forn.get('% Total', 0)
                    classe = 'A' if pct > 10 else ('B' if pct > 5 else 'C')
                    ranking_data.append([
                        str(i),
                        forn.get('Fornecedor', '')[:35],
                        fmt_currency(forn.get('Total Pago', 0)),
                        f"{pct:.1f}%",
                        str(forn.get('Qtd NFs', 0)),
                        fmt_currency(forn.get('Ticket Médio', 0)),
                        f"Classe {classe}"
                    ])
                
                ranking_table = Table(ranking_data, colWidths=[1*cm, 8*cm, 4*cm, 2*cm, 2*cm, 4*cm, 2.5*cm])
                ranking_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2d3748')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f7fafc')]),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ]))
                elements.append(ranking_table)
            
            elements.append(PageBreak())
            
            # ═══════════════════════════════════════════════════════════════════
            # CURVA ABC
            # ═══════════════════════════════════════════════════════════════════
            
            elements.append(Paragraph("📊 ANÁLISE CURVA ABC", section_style))
            
            curva_abc = self.get_curva_abc()
            
            if curva_abc:
                # Criar gráfico de pizza
                fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
                
                # Gráfico de pizza
                labels = ['Classe A (80%)', 'Classe B (15%)', 'Classe C (5%)']
                valores_abc = [
                    curva_abc.get('classe_a', {}).get('valor', 0),
                    curva_abc.get('classe_b', {}).get('valor', 0),
                    curva_abc.get('classe_c', {}).get('valor', 0)
                ]
                cores_abc = ['#f56565', '#ed8936', '#48bb78']
                explode = (0.05, 0, 0)
                
                ax1.pie(valores_abc, labels=labels, colors=cores_abc, explode=explode,
                       autopct='%1.1f%%', startangle=90, shadow=True)
                ax1.set_title('Distribuição de Valor por Classe', fontsize=12, fontweight='bold')
                
                # Gráfico de barras com quantidades
                qtds_abc = [
                    curva_abc.get('classe_a', {}).get('qtd', 0),
                    curva_abc.get('classe_b', {}).get('qtd', 0),
                    curva_abc.get('classe_c', {}).get('qtd', 0)
                ]
                
                bars = ax2.bar(['Classe A', 'Classe B', 'Classe C'], qtds_abc, color=cores_abc)
                ax2.set_ylabel('Quantidade de Fornecedores', fontsize=10)
                ax2.set_title('Quantidade de Fornecedores por Classe', fontsize=12, fontweight='bold')
                
                for bar, val in zip(bars, qtds_abc):
                    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                            str(int(val)), ha='center', va='bottom', fontweight='bold')
                
                ax2.spines['top'].set_visible(False)
                ax2.spines['right'].set_visible(False)
                
                plt.tight_layout()
                
                img_buffer = io.BytesIO()
                plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight',
                           facecolor='white', edgecolor='none')
                img_buffer.seek(0)
                plt.close(fig)
                
                img = Image(img_buffer, width=24*cm, height=9*cm)
                elements.append(img)
                elements.append(Spacer(1, 0.5*cm))
                
                # Tabela resumo ABC
                abc_summary = [
                    ['Classe', 'Valor Total', '% do Valor', 'Qtd Fornecedores', 'Característica'],
                    ['A', fmt_currency(curva_abc.get('classe_a', {}).get('valor', 0)), '80%', 
                     str(curva_abc.get('classe_a', {}).get('qtd', 0)), 'Poucos fornecedores, alto valor'],
                    ['B', fmt_currency(curva_abc.get('classe_b', {}).get('valor', 0)), '15%', 
                     str(curva_abc.get('classe_b', {}).get('qtd', 0)), 'Volume médio'],
                    ['C', fmt_currency(curva_abc.get('classe_c', {}).get('valor', 0)), '5%', 
                     str(curva_abc.get('classe_c', {}).get('qtd', 0)), 'Muitos fornecedores, baixo valor'],
                ]
                
                abc_table = Table(abc_summary, colWidths=[2*cm, 5*cm, 3*cm, 4*cm, 8*cm])
                abc_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2d3748')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('BACKGROUND', (0, 1), (0, 1), colors.HexColor('#fed7d7')),
                    ('BACKGROUND', (0, 2), (0, 2), colors.HexColor('#feebc8')),
                    ('BACKGROUND', (0, 3), (0, 3), colors.HexColor('#c6f6d5')),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('ALIGN', (-1, 1), (-1, -1), 'LEFT'),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                    ('TOPPADDING', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ]))
                elements.append(abc_table)
            
            elements.append(PageBreak())
            
            # ═══════════════════════════════════════════════════════════════════
            # EVOLUÇÃO MENSAL
            # ═══════════════════════════════════════════════════════════════════
            
            elements.append(Paragraph("📈 EVOLUÇÃO MENSAL DE GASTOS", section_style))
            
            evolucao = self.get_evolucao_mensal()
            
            if evolucao and evolucao.get('total'):
                fig, ax = plt.subplots(figsize=(14, 5))
                
                meses = [e.get('Mes', '') for e in evolucao.get('total', [])]
                valores_mensais = [e.get('Valor', 0) for e in evolucao.get('total', [])]
                
                # Linha de evolução
                ax.plot(meses, valores_mensais, marker='o', linewidth=2.5, 
                       markersize=8, color='#2d3748', markerfacecolor='white',
                       markeredgewidth=2)
                
                # Área preenchida
                ax.fill_between(meses, valores_mensais, alpha=0.3, color='#2d3748')
                
                # Adicionar valores nos pontos
                for i, (mes, val) in enumerate(zip(meses, valores_mensais)):
                    ax.annotate(fmt_currency(val), (i, val), textcoords="offset points", 
                               xytext=(0, 10), ha='center', fontsize=8, fontweight='bold')
                
                ax.set_xlabel('Mês', fontsize=10, fontweight='bold')
                ax.set_ylabel('Valor (R$)', fontsize=10, fontweight='bold')
                ax.set_title('Evolução dos Gastos ao Longo do Ano', fontsize=12, fontweight='bold', pad=15)
                ax.grid(True, alpha=0.3)
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                
                img_buffer = io.BytesIO()
                plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight',
                           facecolor='white', edgecolor='none')
                img_buffer.seek(0)
                plt.close(fig)
                
                img = Image(img_buffer, width=24*cm, height=9*cm)
                elements.append(img)
            
            elements.append(Spacer(1, 1*cm))
            
            # ═══════════════════════════════════════════════════════════════════
            # ALERTAS E RECOMENDAÇÕES
            # ═══════════════════════════════════════════════════════════════════
            
            elements.append(Paragraph("⚠️ ALERTAS E RECOMENDAÇÕES", section_style))
            
            alertas = self.get_alertas()
            
            if alertas:
                for alerta in alertas[:5]:
                    tipo = alerta.get('tipo', 'INFO')
                    cor_fundo = colors.HexColor('#fff5f5') if tipo == 'CRÍTICO' else colors.HexColor('#fffaf0')
                    cor_borda = colors.HexColor('#c53030') if tipo == 'CRÍTICO' else colors.HexColor('#c05621')
                    
                    alerta_data = [[
                        Paragraph(f"<b>{tipo}</b> - {alerta.get('categoria', '')}", normal_style),
                        Paragraph(alerta.get('mensagem', ''), normal_style)
                    ]]
                    
                    alerta_table = Table(alerta_data, colWidths=[6*cm, 18*cm])
                    alerta_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, -1), cor_fundo),
                        ('BOX', (0, 0), (-1, -1), 2, cor_borda),
                        ('LEFTPADDING', (0, 0), (-1, -1), 12),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
                        ('TOPPADDING', (0, 0), (-1, -1), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                    ]))
                    elements.append(alerta_table)
                    elements.append(Spacer(1, 0.3*cm))
            else:
                elements.append(Paragraph("✅ Nenhum alerta crítico identificado.", normal_style))
            
            elements.append(PageBreak())
            
            # ═══════════════════════════════════════════════════════════════════
            # ANÁLISE DE ESTOQUE (se disponível)
            # ═══════════════════════════════════════════════════════════════════
            
            if self.df_estoque is not None and len(self.df_estoque) > 0:
                elements.append(Paragraph("📦 ANÁLISE DE ESTOQUE", section_style))
                
                estoque_org = self.get_estoque_por_organizacao()
                analise_grupos = self.get_analise_grupos()
                
                # Gráfico de estoque por grupo
                if analise_grupos and analise_grupos.get('resumo_grupos'):
                    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
                    
                    grupos = [g['Grupo'][:20] for g in analise_grupos.get('resumo_grupos', [])[:8]]
                    valores_grupos = [g['Valor Total'] for g in analise_grupos.get('resumo_grupos', [])[:8]]
                    
                    cores_grupos = plt.cm.Blues(np.linspace(0.3, 0.9, len(grupos)))
                    
                    bars = ax1.barh(grupos[::-1], valores_grupos[::-1], color=cores_grupos[::-1])
                    ax1.set_xlabel('Valor Total (R$)', fontsize=10)
                    ax1.set_title('Top 8 Grupos por Valor em Estoque', fontsize=12, fontweight='bold')
                    ax1.spines['top'].set_visible(False)
                    ax1.spines['right'].set_visible(False)
                    
                    # Gráfico por organização
                    if estoque_org and estoque_org.get('dados'):
                        orgs = [o['Nome Organização'][:15] for o in estoque_org.get('dados', [])[:6]]
                        valores_orgs = [o['Valor Total'] for o in estoque_org.get('dados', [])[:6]]
                        cores_orgs = ['#2d3748', '#1a202c', '#f56565', '#ed8936', '#48bb78', '#4299e1']
                        
                        ax2.pie(valores_orgs, labels=orgs, colors=cores_orgs[:len(orgs)],
                               autopct='%1.1f%%', startangle=90)
                        ax2.set_title('Distribuição por Organização', fontsize=12, fontweight='bold')
                    
                    plt.tight_layout()
                    
                    img_buffer = io.BytesIO()
                    plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight',
                               facecolor='white', edgecolor='none')
                    img_buffer.seek(0)
                    plt.close(fig)
                    
                    img = Image(img_buffer, width=24*cm, height=9*cm)
                    elements.append(img)
                
                elements.append(Spacer(1, 0.5*cm))
                
                # KPIs de estoque
                total_estoque = analise_grupos.get('total_geral', 0) if analise_grupos else 0
                total_itens = analise_grupos.get('total_itens', 0) if analise_grupos else 0
                
                estoque_kpi_data = [
                    [
                        Paragraph(fmt_currency(total_estoque), kpi_value_style),
                        Paragraph(str(total_itens), kpi_value_style),
                        Paragraph(str(len(analise_grupos.get('resumo_grupos', [])) if analise_grupos else 0), kpi_value_style),
                    ],
                    [
                        Paragraph('VALOR TOTAL EM ESTOQUE', kpi_label_style),
                        Paragraph('TOTAL DE ITENS', kpi_label_style),
                        Paragraph('GRUPOS DE MATERIAIS', kpi_label_style),
                    ]
                ]
                
                estoque_kpi_table = Table(estoque_kpi_data, colWidths=[8*cm, 8*cm, 8*cm])
                estoque_kpi_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f7fafc')),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                    ('TOPPADDING', (0, 0), (-1, -1), 15),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
                ]))
                elements.append(estoque_kpi_table)
            
            # ═══════════════════════════════════════════════════════════════════
            # RODAPÉ
            # ═══════════════════════════════════════════════════════════════════
            
            elements.append(Spacer(1, 2*cm))
            
            footer_style = ParagraphStyle(
                'Footer',
                parent=styles['Normal'],
                fontSize=8,
                textColor=colors.HexColor('#a0aec0'),
                alignment=TA_CENTER
            )
            
            elements.append(Paragraph("─" * 80, footer_style))
            elements.append(Paragraph(
                f"Relatório gerado automaticamente pelo Painel Inteligente de Gestão | {data_geracao}",
                footer_style
            ))
            elements.append(Paragraph("Documento confidencial - Uso interno", footer_style))
            
            # Construir PDF
            doc.build(elements)
            
            return True
            
        except Exception as e:
            print(f"Erro ao exportar PDF: {e}")
            traceback.print_exc()
            return False


# ═══════════════════════════════════════════════════════════════════════════════
# CLASSE DE ANÁLISE DE PAGAMENTOS
# ═══════════════════════════════════════════════════════════════════════════════

class PaymentAnalyzer:
    """Classe para análise avançada de dados de pagamentos."""
    
    def __init__(self, df_pagamentos: pd.DataFrame = None):
        self.df = df_pagamentos
        self._processar_dados()
    
    def _processar_dados(self):
        """Processa e limpa dados de pagamentos."""
        if self.df is None or len(self.df) == 0:
            return
        
        # Converter colunas numéricas
        colunas_numericas = ['Valor Total', 'Valor Pago', 'ISS', 'IR', 'INSS', 'PIS', 'COFINS', 'CSLL']
        for col in colunas_numericas:
            if col in self.df.columns:
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce').fillna(0)
        
        # Processar condição de pagamento (extrair número de dias)
        if 'Condição de Pagamento' in self.df.columns:
            self.df['Dias Pagamento'] = self.df['Condição de Pagamento'].apply(self._extrair_dias_pagamento)
        
        # Processar datas
        colunas_data = ['Data NFF', 'Data de Vencimento', 'Data do Pagamento', 'Data GL', 'Data Criação']
        for col in colunas_data:
            if col in self.df.columns:
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
    
    def _extrair_dias_pagamento(self, condicao: str) -> int:
        """Extrai número de dias da condição de pagamento."""
        if pd.isna(condicao):
            return 0
        
        import re
        match = re.search(r'(\d+)', str(condicao))
        if match:
            return int(match.group(1))
        return 0
    
    def get_fornecedores_mais_faturados(self, top_n: int = 10) -> pd.DataFrame:
        """Retorna os fornecedores que mais faturaram."""
        if self.df is None or len(self.df) == 0:
            return pd.DataFrame()
        
        resultado = self.df.groupby('Fornecedor').agg({
            'Valor Pago': 'sum',
            'Número da NFF': 'count'
        }).reset_index()
        
        resultado.columns = ['Fornecedor', 'Valor Total Pago', 'Quantidade NFFs']
        resultado = resultado.sort_values('Valor Total Pago', ascending=False).head(top_n)
        
        return resultado
    
    def get_resumo_formas_pagamento(self) -> Dict[str, Any]:
        """Retorna resumo das formas de pagamento."""
        if self.df is None or len(self.df) == 0:
            return {'resumo': pd.DataFrame(), 'boleto': [], 'deposito': []}
        
        # Criar coluna simplificada: Boleto ou Depósito (case-insensitive)
        self.df['Tipo Pagamento'] = self.df['Método de Pagamento'].apply(
            lambda x: 'Boleto' if pd.notna(x) and 'BOLETO' in str(x).upper() else 'Depósito'
        )
        
        # Agrupar por tipo simplificado
        resumo = self.df.groupby('Tipo Pagamento').agg({
            'Fornecedor': 'count',
            'Valor Pago': 'sum'
        }).reset_index()
        resumo.columns = ['Método', 'Quantidade', 'Valor Total']
        
        # Separar fornecedores por boleto e depósito (único por fornecedor)
        boleto_fornecedores = self.df[
            self.df['Tipo Pagamento'] == 'Boleto'
        ]['Fornecedor'].unique().tolist()
        
        deposito_fornecedores = self.df[
            self.df['Tipo Pagamento'] == 'Depósito'
        ]['Fornecedor'].unique().tolist()
        
        return {
            'resumo': resumo,
            'boleto': boleto_fornecedores,
            'deposito': deposito_fornecedores
        }
    
    def get_curva_abc(self) -> Dict[str, Any]:
        """Calcula Curva ABC por valor financeiro."""
        if self.df is None or len(self.df) == 0:
            return {'tabela': pd.DataFrame(), 'classe_a': [], 'classe_b': [], 'classe_c': []}
        
        # Agrupar por fornecedor
        fornecedores = self.df.groupby('Fornecedor').agg({
            'Valor Pago': 'sum'
        }).reset_index()
        
        fornecedores = fornecedores.sort_values('Valor Pago', ascending=False)
        fornecedores['Percentual'] = (fornecedores['Valor Pago'] / fornecedores['Valor Pago'].sum()) * 100
        fornecedores['Percentual Acumulado'] = fornecedores['Percentual'].cumsum()
        
        # Classificar em A, B, C
        def classificar_abc(perc_acum):
            if perc_acum <= 80:
                return 'A'
            elif perc_acum <= 95:
                return 'B'
            else:
                return 'C'
        
        fornecedores['Classe'] = fornecedores['Percentual Acumulado'].apply(classificar_abc)
        fornecedores.rename(columns={'Valor Pago': 'Valor Total'}, inplace=True)
        
        return {
            'tabela': fornecedores,
            'classe_a': fornecedores[fornecedores['Classe'] == 'A']['Fornecedor'].tolist(),
            'classe_b': fornecedores[fornecedores['Classe'] == 'B']['Fornecedor'].tolist(),
            'classe_c': fornecedores[fornecedores['Classe'] == 'C']['Fornecedor'].tolist()
        }
    
    def get_total_impostos(self) -> Dict[str, float]:
        """Calcula total de impostos."""
        if self.df is None or len(self.df) == 0:
            return {
                'ISS': 0, 'IR': 0, 'INSS': 0, 'PIS': 0, 'COFINS': 0, 'CSLL': 0, 'Total': 0
            }
        
        impostos = {
            'ISS': self.df['ISS'].sum() if 'ISS' in self.df.columns else 0,
            'IR': self.df['IR'].sum() if 'IR' in self.df.columns else 0,
            'INSS': self.df['INSS'].sum() if 'INSS' in self.df.columns else 0,
            'PIS': self.df['PIS'].sum() if 'PIS' in self.df.columns else 0,
            'COFINS': self.df['COFINS'].sum() if 'COFINS' in self.df.columns else 0,
            'CSLL': self.df['CSLL'].sum() if 'CSLL' in self.df.columns else 0,
        }
        
        impostos['Total'] = sum(impostos.values())
        
        return impostos
    
    def get_fornecedores_condicao_menor_28_dias(self) -> pd.DataFrame:
        """Lista fornecedores com condição de pagamento menor que 28 dias."""
        if self.df is None or len(self.df) == 0 or 'Dias Pagamento' not in self.df.columns:
            return pd.DataFrame()
        
        df_filtrado = self.df[self.df['Dias Pagamento'] < 28]
        
        resultado = df_filtrado.groupby(['Fornecedor', 'Condição de Pagamento']).agg({
            'Valor Pago': 'sum',
            'Número da NFF': 'count',
            'Dias Pagamento': 'first'
        }).reset_index()
        
        resultado = resultado.sort_values('Valor Pago', ascending=False)
        resultado.columns = ['Fornecedor', 'Condição Pagamento', 'Valor Total', 'Qtd NFFs', 'Dias']
        
        return resultado
    
    def buscar_fornecedor(self, termo_busca: str) -> Dict[str, Any]:
        """Busca fornecedor por nome ou CNPJ."""
        if self.df is None or len(self.df) == 0 or not termo_busca:
            return {
                'encontrado': False,
                'mensagem': 'Nenhum fornecedor encontrado'
            }
        
        termo = termo_busca.strip().upper()
        
        # Buscar por nome ou CNPJ
        mask = (
            self.df['Fornecedor'].str.upper().str.contains(termo, na=False) |
            self.df['Nr do Fornecedor'].astype(str).str.contains(termo, na=False)
        )
        
        df_fornecedor = self.df[mask]
        
        if len(df_fornecedor) == 0:
            return {
                'encontrado': False,
                'mensagem': f'Nenhum fornecedor encontrado com o termo "{termo_busca}"'
            }
        
        # Agrupar dados do fornecedor
        fornecedor_info = df_fornecedor.groupby('Fornecedor').agg({
            'Nr do Fornecedor': 'first',
            'Valor Pago': 'sum',
            'Número da NFF': 'count',
            'Data do Pagamento': ['min', 'max']
        }).reset_index()
        
        # Achatarcolumns
        fornecedor_info.columns = ['Fornecedor', 'CNPJ', 'Valor Total Pago', 'Qtd Faturamentos', 
                                    'Primeira Data', 'Última Data']
        
        # Pegar histórico de pagamentos
        historico = df_fornecedor.sort_values('Data do Pagamento', ascending=False)[[
            'Número da NFF', 'Data do Pagamento', 'Valor Pago', 'Método de Pagamento', 'Status do Pagamento'
        ]].head(20).to_dict('records')
        
        # Dados consolidados
        resultado = []
        for _, row in fornecedor_info.iterrows():
            resultado.append({
                'nome': row['Fornecedor'],
                'cnpj': row['CNPJ'],
                'valor_total_pago': row['Valor Total Pago'],
                'qtd_faturamentos': int(row['Qtd Faturamentos']),
                'primeira_data': row['Primeira Data'],
                'ultima_data': row['Última Data'],
                'historico': historico
            })
        
        return {
            'encontrado': True,
            'fornecedores': resultado
        }


# ═══════════════════════════════════════════════════════════════════════════════
# BRIDGE PARA COMUNICAÇÃO PYTHON-JAVASCRIPT
# ═══════════════════════════════════════════════════════════════════════════════


class WebBridge(QObject):
    """Ponte de comunicação entre JavaScript e Python."""
    
    showSupplierDetailSignal = pyqtSignal(str)
    applyFiltersSignal = pyqtSignal()
    exportReportSignal = pyqtSignal()
    goBackSignal = pyqtSignal()
    generateReportSignal = pyqtSignal(str)
    filtrarOrganizacaoSignal = pyqtSignal(str)
    filtrarOrganizacaoOrdensSignal = pyqtSignal(str)
    filtrarOrganizacaoMovimentacoesSignal = pyqtSignal(str)
    importarEstoqueSignal = pyqtSignal()
    mostrarMateriaisGrupoSignal = pyqtSignal(str)
    loadAtividadeSubmenuSignal = pyqtSignal(str)
    # Sinais Pipefy
    carregarDadosPipefySignal = pyqtSignal()
    exportarPdfPipefySignal = pyqtSignal()
    atualizarDadosPipefySignal = pyqtSignal()
    enviarEmailPipefySignal = pyqtSignal()
    
    @pyqtSlot(str)
    def showSupplierDetail(self, supplier_name: str):
        self.showSupplierDetailSignal.emit(supplier_name)
    
    @pyqtSlot()
    def applyFilters(self):
        self.applyFiltersSignal.emit()
    
    @pyqtSlot()
    def exportReport(self):
        self.exportReportSignal.emit()
    
    @pyqtSlot()
    def exportRanking(self):
        self.exportReportSignal.emit()
    
    @pyqtSlot()
    def refreshData(self):
        self.applyFiltersSignal.emit()
    
    @pyqtSlot()
    def goBack(self):
        self.goBackSignal.emit()
    
    @pyqtSlot(str)
    def generateReport(self, report_type: str):
        self.generateReportSignal.emit(report_type)
    
    @pyqtSlot(str)
    def filtrarOrganizacao(self, organizacao: str):
        self.filtrarOrganizacaoSignal.emit(organizacao)
    
    @pyqtSlot()
    def importarEstoque(self):
        self.importarEstoqueSignal.emit()
    
    @pyqtSlot(str)
    def mostrarMateriaisGrupo(self, grupo: str):
        self.mostrarMateriaisGrupoSignal.emit(grupo)
    
    @pyqtSlot(str)
    def filtrarOrganizacaoOrdens(self, organizacao: str):
        self.filtrarOrganizacaoOrdensSignal.emit(organizacao)
    
    @pyqtSlot(str)
    def filtrarOrganizacaoMovimentacoes(self, organizacao: str):
        self.filtrarOrganizacaoMovimentacoesSignal.emit(organizacao)
    
    @pyqtSlot(str)
    def loadAtividadeSubmenu(self, submenu: str):
        self.loadAtividadeSubmenuSignal.emit(submenu)
    
    @pyqtSlot()
    def carregarDadosPipefy(self):
        self.carregarDadosPipefySignal.emit()
    
    @pyqtSlot()
    def exportarPdfPipefy(self):
        self.exportarPdfPipefySignal.emit()
    
    @pyqtSlot()
    def atualizarDadosPipefy(self):
        self.atualizarDadosPipefySignal.emit()
    
    @pyqtSlot()
    def enviarEmailPipefy(self):
        self.enviarEmailPipefySignal.emit()


# ═══════════════════════════════════════════════════════════════════════════════
# TEMPLATES HTML/CSS
# ═══════════════════════════════════════════════════════════════════════════════

def get_base_html() -> str:
    """Retorna o template HTML base com estilos modernos."""
    return """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Painel de Gestão</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        :root {
            --primary: #2d3748;
            --primary-dark: #5a67d8;
            --secondary: #1a202c;
            --success: #48bb78;
            --warning: #ed8936;
            --danger: #f56565;
            --info: #4299e1;
            --dark: #1a202c;
            --light: #f7fafc;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.12);
            --shadow-md: 0 4px 6px rgba(0,0,0,0.1);
            --shadow-lg: 0 10px 25px rgba(0,0,0,0.1);
            --radius: 16px;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
            color: var(--dark);
            line-height: 1.6;
            min-height: 100vh;
        }
        
        .dashboard-container {
            padding: 24px;
            max-width: 1800px;
            margin: 0 auto;
        }
        
        .header {
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            border-radius: var(--radius);
            padding: 32px;
            margin-bottom: 24px;
            color: white;
            box-shadow: var(--shadow-lg);
            position: relative;
            overflow: hidden;
        }
        
        .header::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -20%;
            width: 60%;
            height: 200%;
            background: rgba(255,255,255,0.1);
            transform: rotate(25deg);
        }
        
        .header h1 {
            font-size: 2rem;
            font-weight: 800;
            margin-bottom: 8px;
        }
        
        .header .subtitle {
            opacity: 0.9;
            font-size: 1rem;
        }
        
        .filter-bar {
            background: white;
            border-radius: var(--radius);
            padding: 20px 24px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-md);
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
            align-items: center;
        }
        
        .filter-select {
            padding: 10px 16px;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            font-size: 0.9rem;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        
        .filter-select:hover {
            border-color: var(--primary);
        }
        
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }
        
        .kpi-card {
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
            transition: all 0.3s ease;
            border-left: 4px solid var(--primary);
        }
        
        .kpi-card:hover {
            transform: translateY(-4px);
            box-shadow: var(--shadow-lg);
        }
        
        .kpi-card.success { border-left-color: var(--success); }
        .kpi-card.warning { border-left-color: var(--warning); }
        .kpi-card.danger { border-left-color: var(--danger); }
        
        .kpi-icon {
            width: 56px;
            height: 56px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.5rem;
            color: white;
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            margin-bottom: 16px;
        }
        
        .kpi-card.success .kpi-icon {
            background: linear-gradient(135deg, var(--success) 0%, #38a169 100%);
        }
        
        .kpi-card.warning .kpi-icon {
            background: linear-gradient(135deg, var(--warning) 0%, #dd6b20 100%);
        }
        
        .kpi-card.danger .kpi-icon {
            background: linear-gradient(135deg, var(--danger) 0%, #e53e3e 100%);
        }
        
        .kpi-label {
            font-size: 0.85rem;
            font-weight: 600;
            color: #718096;
            text-transform: uppercase;
            margin-bottom: 8px;
        }
        
        .kpi-value {
            font-size: 2rem;
            font-weight: 800;
            color: var(--dark);
            line-height: 1.2;
        }
        
        .kpi-change {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            font-size: 0.85rem;
            font-weight: 600;
            margin-top: 8px;
            padding: 4px 10px;
            border-radius: 20px;
        }
        
        .kpi-change.positive {
            color: #22543d;
            background: #c6f6d5;
        }
        
        .kpi-change.negative {
            color: #742a2a;
            background: #fed7d7;
        }
        
        .chart-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 24px;
            margin-bottom: 24px;
            width: 100%;
            max-width: 100%;
            overflow: hidden;
        }
        
        @media (max-width: 1200px) {
            .chart-grid {
                grid-template-columns: 1fr;
            }
        }
        
        .chart-card {
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
            min-width: 0;
            overflow: hidden;
        }
        
        .chart-card.full-width {
            grid-column: span 2;
        }
        
        @media (max-width: 1200px) {
            .chart-card.full-width {
                grid-column: span 1;
            }
        }
        
        .chart-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 16px;
            border-bottom: 1px solid #e2e8f0;
        }
        
        .chart-title {
            font-size: 1.1rem;
            font-weight: 700;
            color: var(--dark);
        }
        
        .chart-container {
            position: relative;
            height: 300px;
            width: 100%;
            max-width: 100%;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }
        
        .data-table thead {
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            color: white;
        }
        
        .data-table th {
            padding: 14px 16px;
            text-align: left;
            font-weight: 600;
            font-size: 0.8rem;
            text-transform: uppercase;
        }
        
        .data-table td {
            padding: 14px 16px;
            border-bottom: 1px solid #e2e8f0;
        }
        
        .data-table tbody tr:hover {
            background: #f7fafc;
            cursor: pointer;
        }
        
        .badge {
            display: inline-flex;
            align-items: center;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
        }
        
        .badge.class-a { background: #fed7d7; color: #c53030; }
        .badge.class-b { background: #feebc8; color: #c05621; }
        .badge.class-c { background: #c6f6d5; color: #276749; }
        
        .alert-card {
            display: flex;
            gap: 16px;
            padding: 16px 20px;
            border-radius: 8px;
            margin-bottom: 12px;
            border-left: 4px solid;
        }
        
        .alert-card.critical {
            background: #fff5f5;
            border-color: #c53030;
        }
        
        .alert-card.warning {
            background: #fffaf0;
            border-color: #c05621;
        }
        
        .alert-icon {
            font-size: 1.2rem;
        }
        
        .alert-content {
            flex: 1;
        }
        
        .alert-title {
            font-weight: 600;
            font-size: 0.9rem;
        }
        
        .alert-message {
            font-size: 0.85rem;
            color: #4a5568;
        }
        
        .animate-fade-in {
            animation: fadeIn 0.5s ease-out;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        /* Estilos para página de estoque */
        .kpis-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }
        
        .stat-card {
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
            transition: all 0.3s ease;
            border-left: 4px solid var(--primary);
        }
        
        .stat-card:hover {
            transform: translateY(-4px);
            box-shadow: var(--shadow-lg);
        }
        
        .stat-card.highlight {
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            color: white;
            border: none;
        }
        
        .stat-card.highlight .stat-header h3,
        .stat-card.highlight .stat-value,
        .stat-card.highlight .stat-label {
            color: white;
        }
        
        .stat-header {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 12px;
        }
        
        .stat-icon {
            font-size: 28px;
        }
        
        .stat-header h3, .stat-header h4 {
            margin: 0;
            font-size: 1rem;
            font-weight: 600;
            color: var(--dark);
        }
        
        .stat-value {
            font-size: 1.8rem;
            font-weight: 800;
            color: var(--dark);
            margin: 8px 0;
        }
        
        .stat-label {
            font-size: 0.85rem;
            color: #718096;
            font-weight: 500;
        }
        
        .export-btn {
            padding: 10px 20px;
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .export-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
        }
    </style>
</head>
<body>
    <div class="dashboard-container" id="dashboard">
        <!-- Conteúdo será injetado aqui -->
    </div>
    <script src="qrc:///qtwebchannel/qwebchannel.js"></script>
    <script>
        new QWebChannel(qt.webChannelTransport, function(channel) {
            window.pyBridge = channel.objects.pyBridge;
        });
        
        function formatCurrency(value) {
            return new Intl.NumberFormat('pt-BR', { 
                style: 'currency', 
                currency: 'BRL' 
            }).format(value);
        }
    </script>
</body>
</html>"""


def generate_dashboard_html(kpis: Dict, ranking: List[Dict], curva_abc: Dict, 
                            evolucao: Dict, alertas: List[Dict], evolucao_estoque: Dict,
                            estoque_org: Dict = None, organizacao_selecionada: str = None) -> str:
    """Gera HTML do dashboard."""
    
    def fmt_currency(val):
        try:
            if val is None or pd.isna(val):
                return "R$ 0,00"
            val_float = float(val)
            return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"
    
    ranking_json = json.dumps(ranking[:10])
    abc_data = json.dumps(curva_abc.get('dados', [])[:15])
    evolucao_total = json.dumps(evolucao.get('total', []))
    evolucao_estoque_json = json.dumps(evolucao_estoque)
    
    var_mensal = kpis.get('variacao_mensal', 0)
    var_class = "positive" if var_mensal >= 0 else "negative"
    var_icon = "↑" if var_mensal >= 0 else "↓"
    
    # Opções para select de organização (se estoque disponível)
    organizacoes_options = ""
    if estoque_org and estoque_org.get('organizacoes'):
        for org in estoque_org.get('organizacoes', []):
            selected = 'selected' if org == organizacao_selecionada else ''
            organizacoes_options += f'<option value="{org}" {selected}>{org}</option>'
    
    # Gerar linhas da tabela
    ranking_rows = ""
    for i, forn in enumerate(ranking[:10], 1):
        classe = "A" if forn.get('% Total', 0) > 10 else ("B" if forn.get('% Total', 0) > 5 else "C")
        ranking_rows += f"""<tr onclick="showSupplierDetail('{forn.get('Fornecedor', '')}')">
            <td><strong>{i}</strong></td>
            <td>{forn.get('Fornecedor', '')}</td>
            <td>{fmt_currency(forn.get('Total Pago', 0))}</td>
            <td>{forn.get('% Total', 0):.1f}%</td>
            <td>{forn.get('Qtd NFs', 0)}</td>
            <td>{fmt_currency(forn.get('Ticket Médio', 0))}</td>
            <td><span class="badge class-{classe.lower()}">Classe {classe}</span></td>
        </tr>"""
    
    # Gerar alertas
    alertas_html = ""
    for alerta in alertas[:5]:
        tipo_class = "critical" if alerta.get('tipo') == 'CRÍTICO' else "warning"
        alertas_html += f"""<div class="alert-card {tipo_class}">
            <div class="alert-icon">{alerta.get('icone', '<i class="fas fa-exclamation-triangle"></i>')}</div>
            <div class="alert-content">
                <div class="alert-title">{alerta.get('tipo', '')} - {alerta.get('categoria', '')}</div>
                <div class="alert-message">{alerta.get('mensagem', '')}</div>
            </div>
        </div>"""
    
    return f"""
    <div class="header animate-fade-in">
        <h1><i class="fas fa-chart-line"></i> Painel Inteligente de Gestão</h1>
        <p class="subtitle">Análise de estoque, fornecedores e ordens de compra</p>
    </div>
    
    <div class="filter-bar animate-fade-in">
        <select class="filter-select">
            <option>Mês Atual</option>
            <option selected>Ano Atual</option>
            <option>Último Trimestre</option>
        </select>
        {f'''
        <select class="filter-select" id="orgSelect" onchange="filtrarOrganizacao()">
            <option value="">Todas as Organizações</option>
            {organizacoes_options}
        </select>
        ''' if organizacoes_options else ''}
        <button onclick="exportReport()" style="padding: 10px 20px; background: var(--primary); color: white; border: none; border-radius: 8px; cursor: pointer; font-weight: 600;">
            <i class='fas fa-download'></i> Exportar
        </button>
    </div>
    
    <div class="kpi-grid">
        <div class="kpi-card animate-fade-in">
            <div class="kpi-icon"><i class="fas fa-dollar-sign"></i></div>
            <div class="kpi-label">Total Pago</div>
            <div class="kpi-value">{fmt_currency(kpis.get('total_pago', 0))}</div>
            <div class="kpi-change {var_class}">{var_icon} {abs(var_mensal):.1f}%</div>
        </div>
        <div class="kpi-card success animate-fade-in">
            <div class="kpi-icon"><i class="fas fa-building"></i></div>
            <div class="kpi-label">Fornecedores</div>
            <div class="kpi-value">{kpis.get('fornecedores_ativos', 0)}</div>
            <div style="font-size: 0.8rem; color: #718096; margin-top: 4px;">Novos: {kpis.get('novos_fornecedores', 0)}</div>
        </div>
        <div class="kpi-card animate-fade-in">
            <div class="kpi-icon"><i class="fas fa-receipt"></i></div>
            <div class="kpi-label">Ticket Médio</div>
            <div class="kpi-value">{fmt_currency(kpis.get('ticket_medio', 0))}</div>
            <div style="font-size: 0.8rem; color: #718096; margin-top: 4px;">{kpis.get('total_notas', 0)} NFs</div>
        </div>
        <div class="kpi-card warning animate-fade-in">
            <div class="kpi-icon"><i class="fas fa-trophy"></i></div>
            <div class="kpi-label">Maior Pagamento</div>
            <div class="kpi-value">{fmt_currency(kpis.get('maior_pagamento', 0))}</div>
            <div style="font-size: 0.8rem; color: #718096; margin-top: 4px;">{kpis.get('maior_fornecedor', '')[:20]}</div>
        </div>
        <div class="kpi-card success animate-fade-in">
            <div class="kpi-icon"><i class="fas fa-boxes"></i></div>
            <div class="kpi-label">Valor em Estoque</div>
            <div class="kpi-value">{fmt_currency(kpis.get('valor_estoque', 0))}</div>
        </div>
    </div>

    
    <div class="chart-grid">
        <div class="chart-card animate-fade-in">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-chart-bar"></i> Top Fornecedores</div>
            </div>
            <div class="chart-container">
                <canvas id="chartTopFornecedores"></canvas>
            </div>
        </div>
        <div class="chart-card animate-fade-in">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-chart-line"></i> Evolução Mensal do Estoque</div>
            </div>
            <div class="chart-container">
                <canvas id="chartEstoqueMensal"></canvas>
            </div>
        </div>
    </div>

    
    <div class="chart-card full-width animate-fade-in">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-list"></i> Ranking de Fornecedores</div>
        </div>
        <table class="data-table">
            <thead>
                <tr>
                    <th>#</th>
                    <th>Fornecedor</th>
                    <th>Total Pago</th>
                    <th>% Total</th>
                    <th>Qtd NFs</th>
                    <th>Ticket Médio</th>
                    <th>Classificação</th>
                </tr>
            </thead>
            <tbody>
                {ranking_rows}
            </tbody>
        </table>
    </div>
    
    <script>
        const rankingData = {ranking_json};
        const evolucaoEstoqueData = {evolucao_estoque_json};
        
        // Chart: Top Fornecedores
        new Chart(document.getElementById('chartTopFornecedores'), {{
            type: 'bar',
            data: {{
                labels: rankingData.map(d => d.Fornecedor.substring(0, 15)),
                datasets: [{{
                    label: 'Total',
                    data: rankingData.map(d => d['Total Pago']),
                    backgroundColor: 'rgba(102, 126, 234, 0.7)',
                    borderRadius: 6
                }}]
            }},
            options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
        }});
        
        // Chart: Evolução Mensal do Estoque
        new Chart(document.getElementById('chartEstoqueMensal'), {{
            type: 'line',
            data: {{
                labels: evolucaoEstoqueData.meses || [],
                datasets: [{{
                    label: 'Valor em Estoque',
                    data: evolucaoEstoqueData.valores || [],
                    borderColor: '#667eea',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 5,
                    pointBackgroundColor: '#667eea',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointHoverRadius: 7
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        display: true,
                        position: 'top'
                    }},
                    tooltip: {{
                        callbacks: {{
                            label: function(context) {{
                                return 'Estoque: R$ ' + context.parsed.y.toLocaleString('pt-BR', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
                            }}
                        }}
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: false,
                        ticks: {{
                            callback: function(value) {{
                                return 'R$ ' + value.toLocaleString('pt-BR');
                            }}
                        }}
                    }}
                }}
            }}
        }});
        
        function showSupplierDetail(name) {{
            if (window.pyBridge) window.pyBridge.showSupplierDetail(name);
        }}
        
        function exportReport() {{
            if (window.pyBridge) window.pyBridge.exportReport();
        }}
        
        function filtrarOrganizacao() {{
            const org = document.getElementById('orgSelect').value;
            if (window.pyBridge) {{
                window.pyBridge.filtrarOrganizacao(org);
            }}
        }}
    </script>
    """


def generate_estoque_html(analise_grupos: Dict, estoque_stage: Dict, 
                          estoque_org: Dict, curva_abc: Dict, organizacao_selecionada: str = None,
                          top10: List[Dict] = None, subinventarios: Dict = None,
                          kpis_estoque: Dict = None, comparativo_org: List[Dict] = None) -> str:
    """Gera HTML da página de estoque."""
    
    # Valores padrão
    top10 = top10 or []
    subinventarios = subinventarios or {'dados': [], 'total': 0}
    kpis_estoque = kpis_estoque or {}
    comparativo_org = comparativo_org or []
    
    def fmt_currency(val):
        try:
            if val is None or pd.isna(val):
                return "R$ 0,00"
            val_float = float(val)
            return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"
    
    def fmt_numero(val):
        try:
            if val is None or pd.isna(val):
                return "0"
            return f"{float(val):,.0f}".replace(",", ".")
        except:
            return "0"
    
    # Dados para gráficos
    grupos_resumo = analise_grupos.get('resumo_grupos', [])
    grupos_labels = [g['Grupo'] for g in grupos_resumo]
    grupos_valores = [g['Valor Total'] for g in grupos_resumo]
    
    org_dados = estoque_org.get('dados', [])
    org_labels = [o['Nome Organização'] for o in org_dados]
    org_valores = [o['Valor Total'] for o in org_dados]
    
    # Opções para select de organização
    organizacoes_options = ""
    for org in estoque_org.get('organizacoes', []):
        selected = 'selected' if org == organizacao_selecionada else ''
        organizacoes_options += f'<option value="{org}" {selected}>{org}</option>'
    
    # Cards de resumo por grupo - CLICÁVEIS
    grupos_cards = ""
    for grupo in grupos_resumo[:8]:  # Top 8 grupos
        grupo_nome = grupo['Grupo'].replace("'", "\\'")
        grupos_cards += f"""
        <div class="stat-card animate-fade-in clickable-card" onclick="mostrarMateriaisGrupo('{grupo_nome}')">
            <div class="stat-header">
                <span class="stat-icon"><i class='fas fa-chart-bar'></i></span>
                <h3>{grupo['Grupo']}</h3>
            </div>
            <div class="stat-value">{fmt_currency(grupo['Valor Total'])}</div>
            <div class="stat-label">{grupo['Quantidade Itens']} itens • Clique para ver lista</div>
        </div>
        """
    
    # Top 5 por grupo - em CARDS ao invés de tabelas
    top5_html = ""
    for grupo, itens in analise_grupos.get('top5_por_grupo', {}).items():
        if not itens:
            continue
        
        items_cards = ""
        for idx, item in enumerate(itens[:5], 1):
            descricao = item.get('Descrição do Item', 'N/A')
            if descricao and len(str(descricao)) > 60:
                descricao = str(descricao)[:60] + '...'
            elif not descricao:
                descricao = 'Sem descrição'
            
            quantidade = item.get('Quantidade', 0)
            custo_unit = item.get('Custo Unitário', 0)
            custo_total = item.get('Custo Total', 0)
            
            # Badge de posição
            badge_color = '#f56565' if idx == 1 else ('#ed8936' if idx == 2 else '#2d3748')
            
            items_cards += f"""
            <div class="item-card">
                <div class="item-rank" style="background: {badge_color};">#{idx}</div>
                <div class="item-details">
                    <div class="item-code">{item.get('Nome do Item', 'N/A')}</div>
                    <div class="item-desc">{descricao}</div>
                    <div class="item-stats">
                        <div class="item-stat">
                            <span class="stat-label">Qtd:</span>
                            <span class="stat-value">{quantidade:.0f}</span>
                        </div>
                        <div class="item-stat">
                            <span class="stat-label">Unit:</span>
                            <span class="stat-value">{fmt_currency(custo_unit)}</span>
                        </div>
                        <div class="item-stat highlight-stat">
                            <span class="stat-label">Total:</span>
                            <span class="stat-value">{fmt_currency(custo_total)}</span>
                        </div>
                    </div>
                </div>
            </div>
            """
        
        top5_html += f"""
        <div class="grupo-card">
            <div class="grupo-header">
                <i class="fas fa-boxes"></i>
                <h3>{grupo}</h3>
            </div>
            <div class="items-grid">
                {items_cards}
            </div>
        </div>
        """
    
    # Card STAGE
    stage_card = f"""
    <div class="stage-card">
        <div class="header" style="margin: 0 0 20px 0; background: linear-gradient(135deg, #1a202c 0%, #2d3748 100%); border-radius: 12px;">
            <h2><i class="fas fa-warehouse"></i> Materiais STAGE</h2>
            <p class="subtitle">Total: {fmt_currency(estoque_stage.get('total', 0))} | Itens: {estoque_stage.get('quantidade_itens', 0)}</p>
        </div>
        <div class="grid-stage">
    """
    
    for grupo_stage in estoque_stage.get('resumo_grupos', []):
        stage_card += f"""
            <div class="stat-card stage-item">
                <div class="stat-header">
                    <span class="stat-icon"><i class='fas fa-box'></i></span>
                    <h4>{grupo_stage['Grupo']}</h4>
                </div>
                <div class="stat-value">{fmt_currency(grupo_stage['Valor Total'])}</div>
                <div class="stat-label">{grupo_stage['Quantidade']} itens</div>
            </div>
        """
    
    stage_card += """
        </div>
    </div>
    """
    
    # Curva ABC
    abc_classe_a = curva_abc.get('classe_a', {})
    abc_classe_b = curva_abc.get('classe_b', {})
    abc_classe_c = curva_abc.get('classe_c', {})
    
    abc_html = f"""
    <div class="abc-section">
        <h3 style="margin-bottom: 20px; color: var(--primary);">
            <i class="fas fa-chart-pie"></i> Curva ABC do Estoque
        </h3>
        <div class="grid-3">
            <div class="stat-card class-a-card clickable-card" onclick="mostrarClasseABC('A')" style="cursor: pointer;">
                <div class="stat-header">
                    <span class="stat-icon">🔴</span>
                    <h4>Classe A</h4>
                </div>
                <div class="stat-value">{fmt_currency(abc_classe_a.get('valor', 0))}</div>
                <div class="stat-label">{abc_classe_a.get('qtd', 0)} itens | 80% do valor</div>
                <div style="margin-top: 8px; font-size: 0.8rem; color: #718096;">📋 Clique para ver lista</div>
            </div>
            <div class="stat-card class-b-card clickable-card" onclick="mostrarClasseABC('B')" style="cursor: pointer;">
                <div class="stat-header">
                    <span class="stat-icon">🟡</span>
                    <h4>Classe B</h4>
                </div>
                <div class="stat-value">{fmt_currency(abc_classe_b.get('valor', 0))}</div>
                <div class="stat-label">{abc_classe_b.get('qtd', 0)} itens | 15% do valor</div>
                <div style="margin-top: 8px; font-size: 0.8rem; color: #718096;">📋 Clique para ver lista</div>
            </div>
            <div class="stat-card class-c-card clickable-card" onclick="mostrarClasseABC('C')" style="cursor: pointer;">
                <div class="stat-header">
                    <span class="stat-icon">🟢</span>
                    <h4>Classe C</h4>
                </div>
                <div class="stat-value">{fmt_currency(abc_classe_c.get('valor', 0))}</div>
                <div class="stat-label">{abc_classe_c.get('qtd', 0)} itens | 5% do valor</div>
                <div style="margin-top: 8px; font-size: 0.8rem; color: #718096;">📋 Clique para ver lista</div>
            </div>
        </div>
        <div style="height: 350px; margin-top: 24px; background: white; border-radius: var(--radius); padding: 20px;">
            <canvas id="chartABCEstoque"></canvas>
        </div>
    </div>
    """
    
    # KPIs Adicionais de Estoque
    kpis_cards = f"""
    <div class="kpis-section">
        <h3 style="margin-bottom: 20px; color: var(--primary);">
            <i class="fas fa-tachometer-alt"></i> Indicadores de Estoque
        </h3>
        <div class="kpi-grid">
            <div class="kpi-card animate-fade-in">
                <div class="kpi-icon"><i class="fas fa-cubes"></i></div>
                <div class="kpi-label">SKUs Únicos</div>
                <div class="kpi-value">{fmt_numero(kpis_estoque.get('qtd_skus', 0))}</div>
            </div>
            <div class="kpi-card animate-fade-in">
                <div class="kpi-icon"><i class="fas fa-calculator"></i></div>
                <div class="kpi-label">Valor Médio/SKU</div>
                <div class="kpi-value">{fmt_currency(kpis_estoque.get('valor_medio', 0))}</div>
            </div>
            <div class="kpi-card warning animate-fade-in">
                <div class="kpi-icon"><i class="fas fa-crown"></i></div>
                <div class="kpi-label">Concentração Top 10</div>
                <div class="kpi-value">{kpis_estoque.get('concentracao_top10', 0):.1f}%</div>
                <div style="font-size: 0.75rem; color: #718096;">do valor total</div>
            </div>
            <div class="kpi-card danger animate-fade-in">
                <div class="kpi-icon"><i class="fas fa-exclamation-triangle"></i></div>
                <div class="kpi-label">Itens Zerados</div>
                <div class="kpi-value">{kpis_estoque.get('itens_zerados', 0)}</div>
                <div style="font-size: 0.75rem; color: #718096;">Qtd = 0 com valor</div>
            </div>
        </div>
    </div>
    """
    
    # Top 10 Itens Mais Caros
    top10_rows = ""
    for idx, item in enumerate(top10, 1):
        descricao = item.get('Descrição do Item', 'N/A')
        if descricao and len(str(descricao)) > 50:
            descricao = str(descricao)[:50] + '...'
        
        badge_color = '#f56565' if idx <= 3 else ('#ed8936' if idx <= 5 else '#2d3748')
        top10_rows += f"""
        <tr>
            <td><span style="display: inline-flex; align-items: center; justify-content: center; width: 30px; height: 30px; border-radius: 50%; background: {badge_color}; color: white; font-weight: 700; font-size: 0.85rem;">#{idx}</span></td>
            <td><div class="codigo-badge">{item.get('Nome do Item', 'N/A')}</div></td>
            <td>{descricao}</td>
            <td>{item.get('Grupo', 'N/A')}</td>
            <td>{item.get('Organização', 'N/A')}</td>
            <td style="text-align: right;"><strong>{fmt_currency(item.get('Custo Total', 0))}</strong></td>
        </tr>
        """
    
    top10_html = f"""
    <div class="chart-card full-width">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-trophy"></i> Top 10 Itens Mais Caros</div>
        </div>
        <table class="data-table">
            <thead>
                <tr>
                    <th style="width: 60px;">#</th>
                    <th style="width: 140px;">Código</th>
                    <th>Descrição</th>
                    <th style="width: 150px;">Grupo</th>
                    <th style="width: 150px;">Organização</th>
                    <th style="width: 140px; text-align: right;">Valor Total</th>
                </tr>
            </thead>
            <tbody>
                {top10_rows}
            </tbody>
        </table>
    </div>
    """
    
    # Análise por Subinventário
    subinv_rows = ""
    for subinv in subinventarios.get('dados', []):
        subinv_rows += f"""
        <tr>
            <td><strong>{subinv.get('Subinventário', 'N/A')}</strong></td>
            <td style="text-align: right;">{fmt_currency(subinv.get('Valor Total', 0))}</td>
            <td style="text-align: right;">{subinv.get('% do Total', 0):.2f}%</td>
            <td style="text-align: right;">{fmt_numero(subinv.get('Qtd SKUs', 0))}</td>
            <td style="text-align: right;">{fmt_numero(subinv.get('Qtd Total Itens', 0))}</td>
        </tr>
        """
    
    subinv_html = f"""
    <div class="chart-card full-width">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-warehouse"></i> Análise por Subinventário</div>
        </div>
        <table class="data-table">
            <thead>
                <tr>
                    <th>Subinventário</th>
                    <th style="text-align: right;">Valor Total</th>
                    <th style="text-align: right;">% do Total</th>
                    <th style="text-align: right;">Qtd SKUs</th>
                    <th style="text-align: right;">Qtd Itens</th>
                </tr>
            </thead>
            <tbody>
                {subinv_rows}
            </tbody>
        </table>
    </div>
    """
    
    # Tabela Comparativa entre Organizações
    comp_rows = ""
    for org in comparativo_org:
        comp_rows += f"""
        <tr>
            <td><strong>{org.get('Código', '')}</strong></td>
            <td>{org.get('Organização', 'N/A')}</td>
            <td style="text-align: right;">{fmt_currency(org.get('Valor Total', 0))}</td>
            <td style="text-align: right;">{org.get('% do Total', 0):.2f}%</td>
            <td style="text-align: right;">{fmt_numero(org.get('Qtd SKUs', 0))}</td>
            <td style="text-align: right;">{fmt_currency(org.get('Valor Médio/SKU', 0))}</td>
        </tr>
        """
    
    comparativo_html = f"""
    <div class="chart-card full-width">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-balance-scale"></i> Comparativo entre Organizações</div>
        </div>
        <table class="data-table">
            <thead>
                <tr>
                    <th style="width: 80px;">Código</th>
                    <th>Organização</th>
                    <th style="text-align: right;">Valor Total</th>
                    <th style="text-align: right;">% do Total</th>
                    <th style="text-align: right;">Qtd SKUs</th>
                    <th style="text-align: right;">Valor Médio/SKU</th>
                </tr>
            </thead>
            <tbody>
                {comp_rows}
            </tbody>
        </table>
    </div>
    """
    
    return f"""
    <div class="header animate-fade-in">
        <h1><i class="fas fa-warehouse"></i> Gestão de Estoque</h1>
        <p class="subtitle">Análise completa de estoque por grupos e organizações</p>
    </div>
    
    <div class="filter-bar animate-fade-in">
        <select class="filter-select" id="orgSelect" onchange="filtrarOrganizacao()">
            <option value="">Todas as Organizações</option>
            {organizacoes_options}
        </select>
        <button class="export-btn" onclick="exportReport()">
            <i class="fas fa-file-pdf"></i> Exportar PDF
        </button>
    </div>
    
    <!-- Barra de Navegação Horizontal -->
    <div class="nav-tabs-container animate-fade-in">
        <div class="nav-tabs">
            <button class="nav-tab active" onclick="showTab('visao-geral')" data-tab="visao-geral">
                <i class="fas fa-tachometer-alt"></i> Visão Geral
            </button>
            <button class="nav-tab" onclick="showTab('indicadores')" data-tab="indicadores">
                <i class="fas fa-chart-line"></i> Indicadores
            </button>
            <button class="nav-tab" onclick="showTab('curva-abc')" data-tab="curva-abc">
                <i class="fas fa-chart-pie"></i> Curva ABC
            </button>
            <button class="nav-tab" onclick="showTab('top-itens')" data-tab="top-itens">
                <i class="fas fa-trophy"></i> Top Itens
            </button>
            <button class="nav-tab" onclick="showTab('subinventarios')" data-tab="subinventarios">
                <i class="fas fa-warehouse"></i> Subinventários
            </button>
            <button class="nav-tab" onclick="showTab('comparativo')" data-tab="comparativo">
                <i class="fas fa-balance-scale"></i> Comparativo
            </button>
        </div>
    </div>
    
    <!-- Tab: Visão Geral -->
    <div id="tab-visao-geral" class="tab-content active">
        <div class="kpis-grid animate-fade-in">
            <div class="stat-card highlight">
                <div class="stat-header">
                    <span class="stat-icon"><i class='fas fa-dollar-sign'></i></span>
                    <h3>Valor Total Estoque</h3>
                </div>
                <div class="stat-value">{fmt_currency(analise_grupos.get('total_geral', 0))}</div>
                <div class="stat-label">{analise_grupos.get('total_itens', 0)} itens totais</div>
            </div>
            
            {grupos_cards}
        </div>
        
        <div class="grid-2 animate-fade-in" style="margin-top: 24px;">
            <div class="chart-card">
                <h3><i class="fas fa-chart-bar"></i> Valor por Grupo</h3>
                <div style="height: 350px; margin-top: 16px;">
                    <canvas id="chartGrupos"></canvas>
                </div>
            </div>
            
            <div class="chart-card">
                <h3><i class="fas fa-building"></i> Valor por Organização</h3>
                <div style="height: 350px; margin-top: 16px;">
                    <canvas id="chartOrganizacoes"></canvas>
                </div>
            </div>
        </div>
        
        <div class="animate-fade-in" style="margin-top: 24px;">
            {stage_card}
        </div>
    </div>
    
    <!-- Tab: Indicadores -->
    <div id="tab-indicadores" class="tab-content">
        {kpis_cards}
    </div>
    
    <!-- Tab: Curva ABC -->
    <div id="tab-curva-abc" class="tab-content">
        {abc_html}
    </div>
    
    <!-- Tab: Top Itens -->
    <div id="tab-top-itens" class="tab-content">
        <div style="margin-bottom: 24px;">
            {top10_html}
        </div>
        {top5_html}
    </div>
    
    <!-- Tab: Subinventários -->
    <div id="tab-subinventarios" class="tab-content">
        {subinv_html}
    </div>
    
    <!-- Tab: Comparativo -->
    <div id="tab-comparativo" class="tab-content">
        {comparativo_html}
    </div>
    
    <style>
        /* Nav Tabs */
        .nav-tabs-container {{
            background: white;
            border-radius: var(--radius);
            padding: 8px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-md);
            overflow-x: auto;
        }}
        
        .nav-tabs {{
            display: flex;
            gap: 8px;
            min-width: max-content;
        }}
        
        .nav-tab {{
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 12px 20px;
            border: none;
            background: transparent;
            color: #718096;
            font-size: 0.95rem;
            font-weight: 600;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.3s ease;
            white-space: nowrap;
        }}
        
        .nav-tab:hover {{
            background: #f7fafc;
            color: var(--primary);
        }}
        
        .nav-tab.active {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
            box-shadow: 0 4px 15px rgba(45, 55, 72, 0.4);
        }}
        
        .nav-tab i {{
            font-size: 1rem;
        }}
        
        /* Tab Content */
        .tab-content {{
            display: none;
            animation: fadeIn 0.3s ease;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        .grupo-card {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: var(--shadow-md);
            transition: transform 0.2s;
        }}
        
        .grupo-card:hover {{
            transform: translateY(-2px);
            box-shadow: var(--shadow-lg);
        }}
        
        .grupo-header {{
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 2px solid #f0f0f0;
        }}
        
        .grupo-header i {{
            color: var(--primary);
            font-size: 24px;
        }}
        
        .grupo-header h3 {{
            margin: 0;
            color: var(--dark);
            font-size: 20px;
            font-weight: 600;
        }}
        
        .items-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
            gap: 16px;
        }}
        
        .item-card {{
            position: relative;
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            border-radius: 12px;
            padding: 16px;
            border: 2px solid #e0e0e0;
            transition: all 0.3s ease;
        }}
        
        .item-card:hover {{
            border-color: var(--primary);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.15);
            transform: translateY(-2px);
        }}
        
        .item-rank {{
            position: absolute;
            top: -10px;
            right: -10px;
            background: var(--primary);
            color: white;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 14px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        }}
        
        .item-details {{
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}
        
        .item-code {{
            font-family: 'Courier New', monospace;
            font-size: 13px;
            font-weight: 700;
            color: var(--primary);
            padding: 4px 8px;
            background: white;
            border-radius: 6px;
            display: inline-block;
            align-self: flex-start;
        }}
        
        .item-desc {{
            font-size: 14px;
            color: #555;
            line-height: 1.4;
            margin-bottom: 8px;
        }}
        
        .item-stats {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 8px;
            margin-top: 8px;
        }}
        
        .item-stat {{
            background: white;
            padding: 8px;
            border-radius: 8px;
            text-align: center;
            border: 1px solid #e0e0e0;
        }}
        
        .item-stat.highlight-stat {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            border: none;
        }}
        
        .item-stat.highlight-stat .stat-label,
        .item-stat.highlight-stat .stat-value {{
            color: white;
        }}
        
        .stat-label {{
            display: block;
            font-size: 11px;
            color: #888;
            font-weight: 500;
            margin-bottom: 2px;
        }}
        
        .stat-value {{
            display: block;
            font-size: 13px;
            font-weight: 700;
            color: var(--dark);
        }}
        
        .stage-card {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-lg);
        }}
        
        .class-a-card {{ border-left: 4px solid #f56565; }}
        .class-b-card {{ border-left: 4px solid #ed8936; }}
        .class-c-card {{ border-left: 4px solid #48bb78; }}
        
        .abc-section {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
        }}
        
        .grid-2 {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(450px, 1fr));
            gap: 24px;
        }}
        
        .grid-3 {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
        }}
        
        .grid-stage {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 16px;
        }}
        
        .stage-item {{
            border-left: 4px solid #1a202c;
            transition: all 0.3s;
        }}
        
        .stage-item:hover {{
            transform: translateY(-4px);
            box-shadow: 0 6px 20px rgba(118, 75, 162, 0.2);
        }}
        
        .clickable-card {{
            cursor: pointer;
            transition: all 0.3s ease;
        }}
        
        .clickable-card:hover {{
            transform: translateY(-4px) scale(1.02);
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.3);
        }}
        
        .clickable-card:active {{
            transform: translateY(-2px) scale(1.01);
        }}
        
        /* KPI Cards Adicionais */
        .kpis-section {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
        }}
        
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
        }}
        
        .kpi-card {{
            background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%);
            border-radius: 16px;
            padding: 24px;
            text-align: center;
            border: 2px solid #e8e8e8;
            transition: all 0.3s ease;
        }}
        
        .kpi-card:hover {{
            transform: translateY(-4px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.1);
            border-color: #2d3748;
        }}
        
        .kpi-card.warning {{
            border-color: #ed8936;
            background: linear-gradient(135deg, #fffaf0 0%, #ffffff 100%);
        }}
        
        .kpi-card.danger {{
            border-color: #f56565;
            background: linear-gradient(135deg, #fff5f5 0%, #ffffff 100%);
        }}
        
        .kpi-icon {{
            width: 50px;
            height: 50px;
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 16px;
            color: white;
            font-size: 20px;
        }}
        
        .kpi-card.warning .kpi-icon {{
            background: linear-gradient(135deg, #ed8936 0%, #dd6b20 100%);
        }}
        
        .kpi-card.danger .kpi-icon {{
            background: linear-gradient(135deg, #f56565 0%, #e53e3e 100%);
        }}
        
        .kpi-label {{
            font-size: 0.9rem;
            color: #718096;
            font-weight: 500;
            margin-bottom: 8px;
        }}
        
        .kpi-value {{
            font-size: 1.5rem;
            font-weight: 700;
            color: var(--dark);
        }}
        
        /* Data Table */
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }}
        
        .data-table thead tr {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
        }}
        
        .data-table th {{
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.8rem;
            letter-spacing: 0.5px;
        }}
        
        .data-table td {{
            padding: 14px 12px;
            border-bottom: 1px solid #e8e8e8;
        }}
        
        .data-table tbody tr:hover {{
            background: #f8f9fa;
        }}
        
        .data-table tbody tr:nth-child(even) {{
            background: #fafafa;
        }}
        
        .data-table tbody tr:nth-child(even):hover {{
            background: #f0f0f0;
        }}
        
        .codigo-badge {{
            display: inline-block;
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
            padding: 4px 10px;
            border-radius: 6px;
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            font-weight: 600;
        }}
        
        .full-width {{
            grid-column: 1 / -1;
        }}
    </style>
    
    <script>
        const gruposLabels = {json.dumps(grupos_labels)};
        const gruposValores = {json.dumps(grupos_valores)};
        const orgLabels = {json.dumps(org_labels)};
        const orgValores = {json.dumps(org_valores)};
        
        // Dados ABC
        const itensClasseA = {json.dumps(curva_abc.get('itens_classe_a', []))};
        const itensClasseB = {json.dumps(curva_abc.get('itens_classe_b', []))};
        const itensClasseC = {json.dumps(curva_abc.get('itens_classe_c', []))};
        
        // Chart: Grupos - COM ONCLICK
        new Chart(document.getElementById('chartGrupos'), {{
            type: 'bar',
            data: {{
                labels: gruposLabels,
                datasets: [{{
                    label: 'Valor Total',
                    data: gruposValores,
                    backgroundColor: 'rgba(102, 126, 234, 0.7)',
                    borderColor: '#2d3748',
                    borderWidth: 2,
                    borderRadius: 8
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                onClick: (event, elements) => {{
                    if (elements.length > 0) {{
                        const index = elements[0].index;
                        const grupo = gruposLabels[index];
                        mostrarMateriaisGrupo(grupo);
                    }}
                }},
                plugins: {{
                    legend: {{ display: false }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        ticks: {{
                            callback: function(value) {{
                                return 'R$ ' + value.toLocaleString('pt-BR');
                            }}
                        }}
                    }}
                }}
            }}
        }});
        
        // Chart: Organizações
        new Chart(document.getElementById('chartOrganizacoes'), {{
            type: 'doughnut',
            data: {{
                labels: orgLabels,
                datasets: [{{
                    data: orgValores,
                    backgroundColor: [
                        '#2d3748', '#1a202c', '#f56565', '#ed8936', '#48bb78',
                        '#4299e1', '#9f7aea', '#38b2ac', '#f6ad55', '#fc8181'
                    ],
                    borderWidth: 2,
                    borderColor: '#fff'
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        position: 'right'
                    }}
                }}
            }}
        }});
        
        // Chart: ABC
        new Chart(document.getElementById('chartABCEstoque'), {{
            type: 'doughnut',
            data: {{
                labels: ['Classe A (80%)', 'Classe B (15%)', 'Classe C (5%)'],
                datasets: [{{
                    data: [{abc_classe_a.get('valor', 0)}, {abc_classe_b.get('valor', 0)}, {abc_classe_c.get('valor', 0)}],
                    backgroundColor: ['#f56565', '#ed8936', '#48bb78'],
                    borderWidth: 3,
                    borderColor: '#fff'
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                cutout: '60%',
                plugins: {{
                    legend: {{
                        position: 'bottom'
                    }}
                }}
            }}
        }});
        
        function filtrarOrganizacao() {{
            const org = document.getElementById('orgSelect').value;
            if (window.pyBridge) {{
                window.pyBridge.filtrarOrganizacao(org);
            }}
        }}
        
        function exportReport() {{
            if (window.pyBridge) window.pyBridge.exportReport();
        }}
        
        function mostrarMateriaisGrupo(grupo) {{
            if (window.pyBridge) {{
                window.pyBridge.mostrarMateriaisGrupo(grupo);
            }}
        }}
        
        function mostrarClasseABC(classe) {{
            let itens = [];
            let titulo = '';
            let cor = '';
            
            if (classe === 'A') {{
                itens = itensClasseA;
                titulo = 'Classe A - Itens de Alto Valor (80% do estoque)';
                cor = '#f56565';
            }} else if (classe === 'B') {{
                itens = itensClasseB;
                titulo = 'Classe B - Itens de Médio Valor (15% do estoque)';
                cor = '#ed8936';
            }} else if (classe === 'C') {{
                itens = itensClasseC;
                titulo = 'Classe C - Itens de Baixo Valor (5% do estoque)';
                cor = '#48bb78';
            }}
            
            if (itens.length === 0) {{
                alert('Nenhum item encontrado para esta classe.');
                return;
            }}
            
            // Criar modal
            const modal = document.createElement('div');
            modal.id = 'modalClasseABC';
            modal.style.cssText = `
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0,0,0,0.7);
                display: flex;
                align-items: center;
                justify-content: center;
                z-index: 9999;
            `;
            
            let linhasTabela = '';
            itens.forEach((item, index) => {{
                const codigo = item['Nome do Item'] || 'N/A';
                const descricao = item['Descrição do Item'] || 'Sem descrição';
                const quantidade = item['Quantidade'] !== undefined ? Number(item['Quantidade']).toFixed(0) : '0';
                const valor = item['Custo Total'] !== undefined ? formatCurrency(item['Custo Total']) : 'R$ 0,00';
                const percentual = item['% Representatividade'] !== undefined ? Number(item['% Representatividade']).toFixed(2) + '%' : '0%';
                
                const bgColor = index % 2 === 0 ? '#f7fafc' : 'white';
                linhasTabela += `
                    <tr style="background: ${{bgColor}};">
                        <td style="padding: 12px; border-bottom: 1px solid #e2e8f0;">${{index + 1}}</td>
                        <td style="padding: 12px; border-bottom: 1px solid #e2e8f0; font-family: monospace; font-weight: 600;">${{codigo}}</td>
                        <td style="padding: 12px; border-bottom: 1px solid #e2e8f0;">${{descricao}}</td>
                        <td style="padding: 12px; border-bottom: 1px solid #e2e8f0; text-align: right;">${{quantidade}}</td>
                        <td style="padding: 12px; border-bottom: 1px solid #e2e8f0; text-align: right; font-weight: 600;">${{valor}}</td>
                        <td style="padding: 12px; border-bottom: 1px solid #e2e8f0; text-align: right;">${{percentual}}</td>
                    </tr>
                `;
            }});
            
            modal.innerHTML = `
                <div style="background: white; border-radius: 12px; width: 95%; max-width: 1200px; max-height: 90vh; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);">
                    <div style="background: linear-gradient(135deg, ${{cor}} 0%, ${{cor}}dd 100%); color: white; padding: 24px; border-radius: 12px 12px 0 0;">
                        <h2 style="margin: 0; font-size: 1.5rem; display: flex; align-items: center; gap: 12px;">
                            <i class="fas fa-list-ul"></i>
                            ${{titulo}}
                        </h2>
                        <p style="margin: 8px 0 0 0; opacity: 0.9;">${{itens.length}} itens encontrados</p>
                    </div>
                    <div style="padding: 24px; overflow-y: auto; max-height: calc(90vh - 180px);">
                        <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
                            <thead>
                                <tr style="background: #2d3748; color: white;">
                                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #4a5568;">#</th>
                                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #4a5568;">Código</th>
                                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #4a5568;">Descrição</th>
                                    <th style="padding: 12px; text-align: right; border-bottom: 2px solid #4a5568;">Quantidade</th>
                                    <th style="padding: 12px; text-align: right; border-bottom: 2px solid #4a5568;">Valor Total</th>
                                    <th style="padding: 12px; text-align: right; border-bottom: 2px solid #4a5568;">% do Estoque</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${{linhasTabela}}
                            </tbody>
                        </table>
                    </div>
                    <div style="padding: 16px 24px; border-top: 2px solid #e2e8f0; display: flex; justify-content: flex-end;">
                        <button onclick="fecharModalABC()" style="background: #2d3748; color: white; border: none; padding: 12px 32px; border-radius: 8px; font-size: 1rem; cursor: pointer; font-weight: 600; transition: all 0.2s;" onmouseover="this.style.background='#1a202c'" onmouseout="this.style.background='#2d3748'">
                            <i class="fas fa-times"></i> Fechar
                        </button>
                    </div>
                </div>
            `;
            
            document.body.appendChild(modal);
            
            // Fechar ao clicar fora
            modal.addEventListener('click', function(e) {{
                if (e.target === modal) {{
                    fecharModalABC();
                }}
            }});
        }}
        
        function fecharModalABC() {{
            const modal = document.getElementById('modalClasseABC');
            if (modal) {{
                modal.remove();
            }}
        }}
        
        function formatCurrency(value) {{
            if (value === null || value === undefined) return 'R$ 0,00';
            return 'R$ ' + Number(value).toLocaleString('pt-BR', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
        }}
        
        // Navegação por Abas
        function showTab(tabName) {{
            // Ocultar todas as tabs
            document.querySelectorAll('.tab-content').forEach(tab => {{
                tab.classList.remove('active');
            }});
            
            // Remover active de todos os botões
            document.querySelectorAll('.nav-tab').forEach(btn => {{
                btn.classList.remove('active');
            }});
            
            // Mostrar tab selecionada
            const tabContent = document.getElementById('tab-' + tabName);
            if (tabContent) {{
                tabContent.classList.add('active');
            }}
            
            // Ativar botão selecionado
            const tabBtn = document.querySelector('.nav-tab[data-tab="' + tabName + '"]');
            if (tabBtn) {{
                tabBtn.classList.add('active');
            }}
        }}
    </script>
    """


def generate_lista_materiais_html(grupo: str, materiais: List[Dict]) -> str:
    """Gera HTML da lista de materiais de um grupo em formato de tabela com paginação."""
    
    def fmt_currency(val):
        try:
            if val is None or pd.isna(val):
                return "R$ 0,00"
            val_float = float(val)
            return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"
    
    def fmt_numero(val):
        try:
            if val is None or pd.isna(val):
                return "0,00"
            val_float = float(val)
            return f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "0,00"
    
    # Linhas da tabela - todas as linhas com data-page para paginação
    items_per_page = 15
    tabela_linhas = ""
    for idx, mat in enumerate(materiais, 1):
        descricao = mat.get('Descrição do Item', 'Sem descrição')
        if descricao and len(str(descricao)) > 100:
            descricao = str(descricao)[:100] + '...'
        
        # Badge de posição para os top 3
        rank_badge = ""
        if idx <= 3:
            badge_colors = {1: '#f56565', 2: '#ed8936', 3: '#48bb78'}
            rank_badge = f'<span class="rank-badge" style="background: {badge_colors[idx]};">#{idx}</span>'
        else:
            rank_badge = f'<span class="rank-number">#{idx}</span>'
        
        org = mat.get('Organização', '')
        subinv = mat.get('Nome do Subinventário', '')
        
        # Calcular página deste item
        page_num = (idx - 1) // items_per_page + 1
        display_style = 'table-row' if page_num == 1 else 'none'
        
        tabela_linhas += f"""
        <tr class="material-row" data-page="{page_num}" style="display: {display_style};">
            <td class="rank-col">{rank_badge}</td>
            <td class="codigo-col">
                <div class="codigo-badge">{mat.get('Nome do Item', 'N/A')}</div>
            </td>
            <td class="descricao-col">{descricao}</td>
            <td class="org-col">{org}</td>
            <td class="subinv-col">{subinv}</td>
            <td class="qtd-col">{fmt_numero(mat.get('Quantidade', 0))}</td>
            <td class="valor-col">{fmt_currency(mat.get('Custo Unitário', 0))}</td>
            <td class="total-col"><strong>{fmt_currency(mat.get('Custo Total', 0))}</strong></td>
        </tr>
        """
    
    total_materiais = len(materiais)
    total_valor = sum(mat.get('Custo Total', 0) for mat in materiais)
    total_pages = (total_materiais + items_per_page - 1) // items_per_page
    
    # Gerar botões de paginação
    pagination_buttons = ""
    if total_pages > 1:
        pagination_buttons = f"""
        <div class="pagination-container">
            <button class="pagination-btn" onclick="goToPage(1)" id="btn-first" disabled>
                <i class="fas fa-angle-double-left"></i>
            </button>
            <button class="pagination-btn" onclick="previousPage()" id="btn-prev" disabled>
                <i class="fas fa-angle-left"></i>
            </button>
            <span class="pagination-info">
                Página <span id="current-page">1</span> de {total_pages}
            </span>
            <button class="pagination-btn" onclick="nextPage()" id="btn-next">
                <i class="fas fa-angle-right"></i>
            </button>
            <button class="pagination-btn" onclick="goToPage({total_pages})" id="btn-last">
                <i class="fas fa-angle-double-right"></i>
            </button>
            <span class="pagination-total">| {total_materiais} itens</span>
        </div>
        """
    
    return f"""
    <div class="header animate-fade-in">
        <h1><i class="fas fa-boxes"></i> {grupo}</h1>
        <p class="subtitle">Total: {fmt_currency(total_valor)} | {total_materiais} materiais</p>
    </div>
    
    <div class="filter-bar animate-fade-in">
        <button class="back-btn" onclick="window.pyBridge.goBack()">
            <i class="fas fa-arrow-left"></i> Voltar para Estoque
        </button>
        <button class="export-btn" onclick="exportReport()">
            <i class="fas fa-file-excel"></i> Exportar
        </button>
    </div>
    
    <div class="chart-card full-width animate-fade-in" style="margin: 24px;">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-list"></i> Lista Completa de Materiais</div>
            <div style="color: #718096; font-size: 0.9rem;">
                Mostrando 15 itens por página • Ordenado por valor (maior → menor)
            </div>
        </div>
        <div style="overflow-x: auto;">
            <table class="materiais-table">
                <thead>
                    <tr>
                        <th style="width: 60px;">#</th>
                        <th style="width: 140px;">Código</th>
                        <th style="min-width: 300px;">Descrição</th>
                        <th style="width: 100px;">Org</th>
                        <th style="width: 120px;">Subinventário</th>
                        <th style="width: 100px;">Quantidade</th>
                        <th style="width: 120px;">Custo Unit.</th>
                        <th style="width: 140px;">Custo Total</th>
                    </tr>
                </thead>
                <tbody id="materiais-tbody">
                    {tabela_linhas}
                </tbody>
            </table>
        </div>
        {pagination_buttons}
    </div>
    
    <style>
        .pagination-container {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            background: white;
            padding: 16px;
            margin-top: 16px;
            border-radius: 8px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        }}
        
        .pagination-btn {{
            width: 36px;
            height: 36px;
            border: none;
            background: #f0f4ff;
            color: #2d3748;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            justify-content: center;
        }}
        
        .pagination-btn:hover:not(:disabled) {{
            background: #2d3748;
            color: white;
        }}
        
        .pagination-btn:disabled {{
            opacity: 0.4;
            cursor: not-allowed;
        }}
        
        .pagination-info {{
            font-weight: 600;
            color: #4a5568;
            padding: 0 12px;
        }}
        
        .pagination-total {{
            color: #718096;
            font-size: 0.9rem;
        }}
        
        .materiais-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }}
        
        .materiais-table thead {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            color: white;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        
        .materiais-table th {{
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
            font-size: 0.8rem;
            text-transform: uppercase;
            white-space: nowrap;
        }}
        
        .materiais-table td {{
            padding: 14px 12px;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: middle;
        }}
        
        .material-row {{
            transition: all 0.2s ease;
        }}
        
        .material-row:hover {{
            background: #f7fafc;
            transform: scale(1.005);
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}
        
        .rank-col {{
            text-align: center;
        }}
        
        .rank-badge {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
            box-shadow: 0 2px 6px rgba(0,0,0,0.15);
        }}
        
        .rank-number {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background: #e2e8f0;
            color: #4a5568;
            font-weight: 600;
            font-size: 0.8rem;
        }}
        
        .codigo-badge {{
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            font-weight: 700;
            color: var(--primary);
            background: #f0f4ff;
            padding: 6px 10px;
            border-radius: 6px;
            display: inline-block;
        }}
        
        .descricao-col {{
            color: #2d3748;
            line-height: 1.4;
        }}
        
        .org-col, .subinv-col {{
            font-size: 0.85rem;
            color: #4a5568;
        }}
        
        .qtd-col {{
            text-align: right;
            font-weight: 600;
            color: #2d3748;
        }}
        
        .valor-col {{
            text-align: right;
            color: #4a5568;
        }}
        
        .total-col {{
            text-align: right;
            font-size: 1rem;
            color: var(--primary);
        }}
        
        .back-btn {{
            padding: 10px 20px;
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        
        .back-btn:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
        }}
    </style>
    
    <script>
        let currentPage = 1;
        const totalPages = {total_pages};
        
        function updatePagination() {{
            document.getElementById('current-page').textContent = currentPage;
            document.getElementById('btn-first').disabled = currentPage === 1;
            document.getElementById('btn-prev').disabled = currentPage === 1;
            document.getElementById('btn-next').disabled = currentPage === totalPages;
            document.getElementById('btn-last').disabled = currentPage === totalPages;
        }}
        
        function showPage(page) {{
            const rows = document.querySelectorAll('.material-row');
            rows.forEach(row => {{
                const rowPage = parseInt(row.getAttribute('data-page'));
                row.style.display = rowPage === page ? 'table-row' : 'none';
            }});
            currentPage = page;
            updatePagination();
            // Scroll para o topo da tabela
            document.querySelector('.materiais-table').scrollIntoView({{ behavior: 'smooth', block: 'start' }});
        }}
        
        function goToPage(page) {{
            if (page >= 1 && page <= totalPages) {{
                showPage(page);
            }}
        }}
        
        function previousPage() {{
            if (currentPage > 1) {{
                showPage(currentPage - 1);
            }}
        }}
        
        function nextPage() {{
            if (currentPage < totalPages) {{
                showPage(currentPage + 1);
            }}
        }}
        
        function exportReport() {{
            if (window.pyBridge) window.pyBridge.exportReport();
        }}
    </script>
    """


def generate_movimentacoes_html(analise: Dict, organizacao_selecionada: str = None) -> str:
    """Gera HTML da página de Movimentações (saídas de estoque)."""
    
    def fmt_currency(val):
        try:
            if val is None or pd.isna(val):
                return "R$ 0,00"
            return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return "R$ 0,00"
    
    def fmt_numero(val):
        try:
            if val is None or pd.isna(val):
                return "0"
            return f"{int(val):,}".replace(",", ".")
        except:
            return "0"
    
    # Opções de organizações
    organizacoes = analise.get('organizacoes', [])
    org_options = ""
    for org in organizacoes:
        selected = "selected" if org == organizacao_selecionada else ""
        org_options += f'<option value="{org}" {selected}>{org}</option>'
    
    # Top 10 Transferências
    transf_rows = ""
    for idx, item in enumerate(analise.get('transf_top10', []), 1):
        descricao = item.get('Descrição', 'N/A')
        if descricao and len(str(descricao)) > 40:
            descricao = str(descricao)[:40] + '...'
        
        badge_color = '#f56565' if idx <= 3 else ('#ed8936' if idx <= 5 else '#2d3748')
        transf_rows += f"""
        <tr>
            <td><span style="display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; border-radius: 50%; background: {badge_color}; color: white; font-weight: 700; font-size: 0.8rem;">#{idx}</span></td>
            <td><div class="codigo-badge">{item.get('Código', 'N/A')}</div></td>
            <td>{descricao}</td>
            <td style="text-align: right;">{fmt_numero(item.get('Quantidade', 0))}</td>
            <td style="text-align: right;"><strong>{fmt_currency(item.get('Valor Total', 0))}</strong></td>
        </tr>
        """
    
    # Top 10 OS
    os_rows = ""
    for idx, item in enumerate(analise.get('os_top10', []), 1):
        descricao = item.get('Descrição', 'N/A')
        if descricao and len(str(descricao)) > 40:
            descricao = str(descricao)[:40] + '...'
        
        badge_color = '#48bb78' if idx <= 3 else ('#38b2ac' if idx <= 5 else '#4299e1')
        os_rows += f"""
        <tr>
            <td><span style="display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; border-radius: 50%; background: {badge_color}; color: white; font-weight: 700; font-size: 0.8rem;">#{idx}</span></td>
            <td><div class="codigo-badge">{item.get('Código', 'N/A')}</div></td>
            <td>{descricao}</td>
            <td style="text-align: right;">{fmt_numero(item.get('Quantidade', 0))}</td>
            <td style="text-align: right;"><strong>{fmt_currency(item.get('Valor Total', 0))}</strong></td>
        </tr>
        """
    
    # Por Grupo - Transferências
    transf_grupo_rows = ""
    for grupo in analise.get('transf_por_grupo', []):
        transf_grupo_rows += f"""
        <tr onclick=\"showMateriaisGrupo('transferencia','{grupo.get('Grupo', 'N/A')}')\" style=\"cursor:pointer;\">
            <td><strong>{grupo.get('Grupo', 'N/A')}</strong></td>
            <td style=\"text-align: right;\">{fmt_currency(grupo.get('Valor Total', 0))}</td>
            <td style=\"text-align: right;\">{grupo.get('Qtd Itens', 0)}</td>
            <td style=\"text-align: right;\">{fmt_numero(grupo.get('Quantidade', 0))}</td>
        </tr>
        """
    
    # Por Grupo - OS
    os_grupo_rows = ""
    for grupo in analise.get('os_por_grupo', []):
        os_grupo_rows += f"""
        <tr onclick=\"showMateriaisGrupo('os','{grupo.get('Grupo', 'N/A')}')\" style=\"cursor:pointer;\">
            <td><strong>{grupo.get('Grupo', 'N/A')}</strong></td>
            <td style=\"text-align: right;\">{fmt_currency(grupo.get('Valor Total', 0))}</td>
            <td style=\"text-align: right;\">{grupo.get('Qtd Itens', 0)}</td>
            <td style=\"text-align: right;\">{fmt_numero(grupo.get('Quantidade', 0))}</td>
        </tr>
        """
    
    # Top 10 Combustível
    comb_rows = ""
    for idx, item in enumerate(analise.get('comb_top10', []), 1):
        descricao = item.get('Descrição', 'N/A')
        if descricao and len(str(descricao)) > 40:
            descricao = str(descricao)[:40] + '...'
        
        badge_color = '#ed8936' if idx <= 3 else ('#f6ad55' if idx <= 5 else '#fbd38d')
        comb_rows += f"""
        <tr>
            <td><span style="display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; border-radius: 50%; background: {badge_color}; color: white; font-weight: 700; font-size: 0.8rem;">#{idx}</span></td>
            <td><div class="codigo-badge">{item.get('Código', 'N/A')}</div></td>
            <td>{descricao}</td>
            <td style="text-align: right;">{fmt_numero(item.get('Quantidade', 0))}</td>
            <td style="text-align: right;"><strong>{fmt_currency(item.get('Valor Total', 0))}</strong></td>
        </tr>
        """
    
    # Por Grupo - Combustível
    comb_grupo_rows = ""
    for grupo in analise.get('comb_por_grupo', []):
        comb_grupo_rows += f"""
        <tr>
            <td><strong>{grupo.get('Grupo', 'N/A')}</strong></td>
            <td style="text-align: right;">{fmt_currency(grupo.get('Valor Total', 0))}</td>
            <td style="text-align: right;">{grupo.get('Qtd Itens', 0)}</td>
            <td style="text-align: right;">{fmt_numero(grupo.get('Quantidade', 0))}</td>
        </tr>
        """
    
    # Dados para gráficos - Consolidar todos os grupos
    transf_grupos = analise.get('transf_por_grupo', [])
    os_grupos = analise.get('os_por_grupo', [])
    comb_grupos = analise.get('comb_por_grupo', [])
    
    # Criar dicionário com todos os grupos únicos
    todos_grupos = {}
    for g in transf_grupos:
        grupo = g['Grupo']
        if grupo not in todos_grupos:
            todos_grupos[grupo] = {'transf': 0, 'os': 0, 'comb': 0}
        todos_grupos[grupo]['transf'] = g['Valor Total']
    
    for g in os_grupos:
        grupo = g['Grupo']
        if grupo not in todos_grupos:
            todos_grupos[grupo] = {'transf': 0, 'os': 0, 'comb': 0}
        todos_grupos[grupo]['os'] = g['Valor Total']
    
    for g in comb_grupos:
        grupo = g['Grupo']
        if grupo not in todos_grupos:
            todos_grupos[grupo] = {'transf': 0, 'os': 0, 'comb': 0}
        todos_grupos[grupo]['comb'] = g['Valor Total']
    
    # Ordenar por valor total (soma dos três)
    grupos_ordenados = sorted(
        todos_grupos.items(),
        key=lambda x: x[1]['transf'] + x[1]['os'] + x[1]['comb'],
        reverse=True
    )
    
    # Preparar dados para gráfico
    grupos_labels = [g[0] for g in grupos_ordenados]
    grupos_transf = [g[1]['transf'] for g in grupos_ordenados]
    grupos_os = [g[1]['os'] for g in grupos_ordenados]
    grupos_comb = [g[1]['comb'] for g in grupos_ordenados]
    
    return f"""
    <div class="header animate-fade-in">
        <h1><i class="fas fa-exchange-alt"></i> Movimentações de Estoque</h1>
        <p class="subtitle">Análise de saídas de estoque por transferências, ordens de serviço e combustível</p>
    </div>
    
    <!-- Modal para Materiais -->
    <div id="materiaisModal" style="display:none;position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(0,0,0,0.5);z-index:9999;align-items:center;justify-content:center;">
        <div style="background:white;padding:32px;border-radius:12px;max-width:900px;width:90%;max-height:85vh;overflow:auto;box-shadow:0 8px 32px rgba(0,0,0,0.3);">
            <h2 id="materiaisModalTitle" style="margin-bottom:20px;font-size:1.3rem;color:#2d3748;border-bottom:2px solid #e2e8f0;padding-bottom:12px;"></h2>
            <div id="materiaisModalBody" style="overflow-x:auto;"></div>
            <div style="margin-top:24px;display:flex;justify-content:flex-end;">
                <button onclick="document.getElementById('materiaisModal').style.display='none'" style="padding:10px 24px;background:#2d3748;color:white;border:none;border-radius:6px;cursor:pointer;font-weight:600;transition:background 0.2s;" onmouseover="this.style.background='#1a202c'" onmouseout="this.style.background='#2d3748'">Fechar</button>
            </div>
        </div>
    </div>
    
    <div class="filter-bar animate-fade-in">
        <select class="filter-select" id="orgSelect" onchange="filtrarOrganizacao()">
            <option value="">Todas as Organizações</option>
            {org_options}
        </select>
        <button class="export-btn" onclick="exportReport()">
            <i class="fas fa-file-pdf"></i> Exportar PDF
        </button>
    </div>
    
    <!-- KPIs Principais -->
    <div class="kpis-grid animate-fade-in">
        <div class="stat-card highlight">
            <div class="stat-header">
                <span class="stat-icon">💸</span>
                <h3>Total de Saídas</h3>
            </div>
            <div class="stat-value">{fmt_currency(analise.get('total_geral', 0))}</div>
            <div class="stat-label">Transferências + OS + Combustível</div>
        </div>
        
        <div class="stat-card transf-card">
            <div class="stat-header">
                <span class="stat-icon">📤</span>
                <h3>Transferências</h3>
            </div>
            <div class="stat-value">{fmt_currency(analise.get('transf_total', 0))}</div>
            <div class="stat-label">{analise.get('transf_qtd', 0)} movimentações | {analise.get('transf_itens', 0)} itens</div>
        </div>
        
        <div class="stat-card os-card">
            <div class="stat-header">
                <span class="stat-icon">🔧</span>
                <h3>Ordens de Serviço</h3>
            </div>
            <div class="stat-value">{fmt_currency(analise.get('os_total', 0))}</div>
            <div class="stat-label">{analise.get('os_qtd', 0)} movimentações | {analise.get('os_itens', 0)} itens</div>
        </div>
        
        <div class="stat-card comb-card">
            <div class="stat-header">
                <span class="stat-icon">⛽</span>
                <h3>Combustível</h3>
            </div>
            <div class="stat-value">{fmt_currency(analise.get('comb_total', 0))}</div>
            <div class="stat-label">{analise.get('comb_qtd', 0)} movimentações | {analise.get('comb_itens', 0)} itens</div>
        </div>
    </div>
    
    <!-- Gráfico Consolidado por Grupo -->
    <div class="chart-card full-width animate-fade-in" style="margin-top: 24px;">
        <h3><i class="fas fa-chart-bar" style="color: #2d3748;"></i> Saídas por Grupo (Transferências, OS e Combustível)</h3>
        <div id="chartGruposConsolidado" style="margin-top: 16px; min-height: 400px;">
            <!-- Gráfico será inserido aqui via JavaScript -->
        </div>
    </div>
    
    <!-- Top 10 por Tipo -->
    <div class="grid-2 animate-fade-in" style="margin-top: 24px;">
        <!-- Top 10 Transferências -->
        <div class="chart-card">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-arrow-up" style="color: #f56565;"></i> Top 10 Transferências</div>
            </div>
            <table class="data-table compact">
                <thead>
                    <tr>
                        <th style="width: 40px;">#</th>
                        <th style="width: 120px;">Código</th>
                        <th>Descrição</th>
                        <th style="width: 70px; text-align: right;">Qtd</th>
                        <th style="width: 110px; text-align: right;">Valor</th>
                    </tr>
                </thead>
                <tbody>
                    {transf_rows if transf_rows else '<tr><td colspan="5" style="text-align: center; color: #999;">Nenhum dado encontrado</td></tr>'}
                </tbody>
            </table>
        </div>
        
        <!-- Top 10 OS -->
        <div class="chart-card">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-wrench" style="color: #48bb78;"></i> Top 10 Ordens de Serviço</div>
            </div>
            <table class="data-table compact">
                <thead>
                    <tr>
                        <th style="width: 40px;">#</th>
                        <th style="width: 120px;">Código</th>
                        <th>Descrição</th>
                        <th style="width: 70px; text-align: right;">Qtd</th>
                        <th style="width: 110px; text-align: right;">Valor</th>
                    </tr>
                </thead>
                <tbody>
                    {os_rows if os_rows else '<tr><td colspan="5" style="text-align: center; color: #999;">Nenhum dado encontrado</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- Resumo por Grupo -->
    <div class="grid-2 animate-fade-in" style="margin-top: 24px;">
        <!-- Transferências por Grupo -->
        <div class="chart-card">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-layer-group"></i> Transferências por Grupo</div>
            </div>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Grupo</th>
                        <th style="text-align: right;">Valor Total</th>
                        <th style="text-align: right;">Itens</th>
                        <th style="text-align: right;">Qtd</th>
                    </tr>
                </thead>
                <tbody>
                    {transf_grupo_rows if transf_grupo_rows else '<tr><td colspan="4" style="text-align: center; color: #999;">Nenhum dado encontrado</td></tr>'}
                </tbody>
            </table>
        </div>
        
        <!-- OS por Grupo -->
        <div class="chart-card">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-layer-group"></i> Ordens de Serviço por Grupo</div>
            </div>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Grupo</th>
                        <th style="text-align: right;">Valor Total</th>
                        <th style="text-align: right;">Itens</th>
                        <th style="text-align: right;">Qtd</th>
                    </tr>
                </thead>
                <tbody>
                    {os_grupo_rows if os_grupo_rows else '<tr><td colspan="4" style="text-align: center; color: #999;">Nenhum dado encontrado</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    
    <style>
        .kpis-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
        }}
        
        .stat-card {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
            transition: all 0.3s ease;
        }}
        
        .stat-card:hover {{
            transform: translateY(-4px);
            box-shadow: var(--shadow-lg);
        }}
        
        .stat-card.highlight {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
        }}
        
        .stat-card.highlight h3,
        .stat-card.highlight .stat-label {{
            color: rgba(255,255,255,0.9);
        }}
        
        .stat-card.transf-card {{
            border-left: 4px solid #f56565;
        }}
        
        .stat-card.os-card {{
            border-left: 4px solid #48bb78;
        }}
        
        .stat-card.comb-card {{
            border-left: 4px solid #ed8936;
        }}
        
        .stat-header {{
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 12px;
        }}
        
        .stat-icon {{
            font-size: 28px;
        }}
        
        .stat-header h3 {{
            margin: 0;
            font-size: 1rem;
            font-weight: 600;
            color: #64748b;
        }}
        
        .stat-value {{
            font-size: 1.8rem;
            font-weight: 700;
            color: var(--dark);
        }}
        
        .stat-card.highlight .stat-value {{
            color: white;
        }}
        
        .stat-label {{
            font-size: 0.85rem;
            color: #94a3b8;
            margin-top: 4px;
        }}
        
        .grid-2 {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(450px, 1fr));
            gap: 24px;
        }}
        
        .chart-card {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
        }}
        
        .chart-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 2px solid #f0f0f0;
        }}
        
        .chart-title {{
            font-size: 1.1rem;
            font-weight: 600;
            color: var(--dark);
        }}
        
        .chart-title i {{
            margin-right: 8px;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }}
        
        .data-table.compact {{
            font-size: 0.85rem;
        }}
        
        .data-table thead tr {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
        }}
        
        .data-table th {{
            padding: 12px 10px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.5px;
        }}
        
        .data-table td {{
            padding: 12px 10px;
            border-bottom: 1px solid #e8e8e8;
        }}
        
        .data-table tbody tr:hover {{
            background: #f8f9fa;
        }}
        
        .codigo-badge {{
            display: inline-block;
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
            padding: 3px 8px;
            border-radius: 5px;
            font-family: 'Courier New', monospace;
            font-size: 0.8rem;
            font-weight: 600;
        }}
    </style>
    
    <script>
        var gruposLabels = {json.dumps(grupos_labels)};
        var gruposTransf = {json.dumps(grupos_transf)};
        var gruposOS = {json.dumps(grupos_os)};
        var gruposComb = {json.dumps(grupos_comb)};
        
        console.log('Dados do grafico carregados:', gruposLabels.length, 'grupos');
        
        // Executar assim que possível
        setTimeout(function() {{
            criarGraficoGrupos();
        }}, 100);
        
        function criarGraficoGrupos() {{
            console.log('Criando grafico de grupos...');
            
            if (gruposLabels.length === 0) {{
                console.log('Nenhum dado para o grafico');
                return;
            }}
            
            var container = document.getElementById('chartGruposConsolidado');
            if (!container) {{
                console.error('Container chartGruposConsolidado nao encontrado!');
                return;
            }}
            
            // Encontrar valor máximo para escala
            var maxValor = 0;
            for (var i = 0; i < gruposLabels.length; i++) {{
                var transf = gruposTransf[i] || 0;
                var os = gruposOS[i] || 0;
                var comb = gruposComb[i] || 0;
                maxValor = Math.max(maxValor, transf, os, comb);
            }}
            
            var html = '';
            html += '<div style="padding: 20px; font-family: Arial, sans-serif;">';
            html += '<div style="display: flex; flex-direction: column; gap: 20px;">';
            
            for (var i = 0; i < gruposLabels.length; i++) {{
                var label = gruposLabels[i];
                var transf = gruposTransf[i] || 0;
                var os = gruposOS[i] || 0;
                var comb = gruposComb[i] || 0;
                var total = transf + os + comb;
                
                // Calcular larguras proporcionais ao valor máximo
                var transfWidth = maxValor > 0 ? (transf / maxValor) * 100 : 0;
                var osWidth = maxValor > 0 ? (os / maxValor) * 100 : 0;
                var combWidth = maxValor > 0 ? (comb / maxValor) * 100 : 0;
                
                html += '<div style="display: flex; flex-direction: column; gap: 4px; padding: 12px; background: #f7fafc; border-radius: 8px;">';
                html += '<div style="font-weight: bold; color: #2d3748; margin-bottom: 6px;">' + label + ' <span style="color: #718096; font-size: 0.9rem;">(Total: R$ ' + total.toLocaleString('pt-BR', {{minimumFractionDigits: 2, maximumFractionDigits: 2}}) + ')</span></div>';
                
                // Barra Transferências
                if (transf > 0) {{
                    html += '<div style="display: flex; align-items: center; gap: 8px;">';
                    html += '<div style="width: 100px; font-size: 0.85rem; color: #2d3748;">Transferências</div>';
                    html += '<div style="flex: 1; background: #fee; border-radius: 4px; height: 20px; position: relative; overflow: hidden;">';
                    html += '<div style="width: ' + transfWidth + '%; background-color: #f56565; height: 100%; transition: width 0.3s;"></div>';
                    html += '</div>';
                    html += '<div style="width: 110px; text-align: right; font-weight: 600; color: #f56565; font-size: 0.9rem;">R$ ' + transf.toLocaleString('pt-BR', {{minimumFractionDigits: 2, maximumFractionDigits: 2}}) + '</div>';
                    html += '</div>';
                }}
                
                // Barra Ordens de Serviço
                if (os > 0) {{
                    html += '<div style="display: flex; align-items: center; gap: 8px;">';
                    html += '<div style="width: 100px; font-size: 0.85rem; color: #2d3748;">Ordens Serviço</div>';
                    html += '<div style="flex: 1; background: #efe; border-radius: 4px; height: 20px; position: relative; overflow: hidden;">';
                    html += '<div style="width: ' + osWidth + '%; background-color: #48bb78; height: 100%; transition: width 0.3s;"></div>';
                    html += '</div>';
                    html += '<div style="width: 110px; text-align: right; font-weight: 600; color: #48bb78; font-size: 0.9rem;">R$ ' + os.toLocaleString('pt-BR', {{minimumFractionDigits: 2, maximumFractionDigits: 2}}) + '</div>';
                    html += '</div>';
                }}
                
                // Barra Combustível
                if (comb > 0) {{
                    html += '<div style="display: flex; align-items: center; gap: 8px;">';
                    html += '<div style="width: 100px; font-size: 0.85rem; color: #2d3748;">Combustível</div>';
                    html += '<div style="flex: 1; background: #fef5e7; border-radius: 4px; height: 20px; position: relative; overflow: hidden;">';
                    html += '<div style="width: ' + combWidth + '%; background-color: #ed8936; height: 100%; transition: width 0.3s;"></div>';
                    html += '</div>';
                    html += '<div style="width: 110px; text-align: right; font-weight: 600; color: #ed8936; font-size: 0.9rem;">R$ ' + comb.toLocaleString('pt-BR', {{minimumFractionDigits: 2, maximumFractionDigits: 2}}) + '</div>';
                    html += '</div>';
                }}
                
                html += '</div>';
            }}
            
            html += '</div>';
            html += '<div style="display: flex; justify-content: center; gap: 20px; margin-top: 20px; font-size: 12px; padding: 12px; background: white; border-radius: 8px;">';
            html += '<div style="display: flex; align-items: center; gap: 5px;"><div style="width: 16px; height: 16px; background-color: #f56565; border-radius: 2px;"></div>Transferências</div>';
            html += '<div style="display: flex; align-items: center; gap: 5px;"><div style="width: 16px; height: 16px; background-color: #48bb78; border-radius: 2px;"></div>Ordens de Serviço</div>';
            html += '<div style="display: flex; align-items: center; gap: 5px;"><div style="width: 16px; height: 16px; background-color: #ed8936; border-radius: 2px;"></div>Combustível</div>';
            html += '</div>';
            html += '</div>';
            
            container.innerHTML = html;
            console.log('Grafico criado com sucesso!');
        }}
        
        function filtrarOrganizacao() {{
            var org = document.getElementById('orgSelect').value;
            if (window.pyBridge) {{
                window.pyBridge.filtrarOrganizacaoMovimentacoes(org);
            }}
        }}
        
        function exportReport() {{
            if (window.pyBridge) window.pyBridge.exportReport();
        }}
        
        // Dados dos materiais por grupo
        var materiaisTransferencia = {json.dumps(analise.get('materiais_por_grupo_transferencia', {}))};
        var materiaisOS = {json.dumps(analise.get('materiais_por_grupo_os', {}))};
        
        function showMateriaisGrupo(tipo, grupo) {{
            var lista = [];
            var title = '';
            if (tipo === 'transferencia') {{
                lista = materiaisTransferencia[grupo] || [];
                title = 'Materiais do Grupo de Transferência: ' + grupo;
            }} else if (tipo === 'os') {{
                lista = materiaisOS[grupo] || [];
                title = 'Materiais do Grupo de OS: ' + grupo;
            }}
            
            console.log('Mostrando materiais:', tipo, grupo, lista.length);
            
            var html = '';
            if (lista.length === 0) {{
                html = '<p style="color:#999;text-align:center;padding:20px;">Nenhum material encontrado para este grupo.</p>';
            }} else {{
                html = '<table style="width:100%;font-size:0.9rem;border-collapse:collapse;">';
                html += '<thead><tr style="background:#2d3748;color:white;">';
                html += '<th style="padding:10px;text-align:left;">Código</th>';
                html += '<th style="padding:10px;text-align:left;">Material</th>';
                html += '<th style="padding:10px;text-align:right;">Quantidade</th>';
                html += '<th style="padding:10px;text-align:right;">Valor</th>';
                html += '</tr></thead><tbody>';
                lista.forEach(function(mat, idx) {{
                    var bgColor = idx % 2 === 0 ? '#f7fafc' : 'white';
                    html += '<tr style="background:' + bgColor + ';">';
                    html += '<td style="padding:8px;font-family:monospace;font-size:0.85rem;">' + (mat.Codigo || '-') + '</td>';
                    html += '<td style="padding:8px;">' + (mat.Material || '-') + '</td>';
                    html += '<td style="padding:8px;text-align:right;font-weight:600;">' + (mat.Quantidade ? Number(mat.Quantidade).toLocaleString('pt-BR', {{minimumFractionDigits:0, maximumFractionDigits:2}}) : '-') + '</td>';
                    html += '<td style="padding:8px;text-align:right;font-weight:600;color:#2d3748;">' + (mat.Valor ? ('R$ ' + Number(mat.Valor).toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:2}})) : '-') + '</td>';
                    html += '</tr>';
                }});
                html += '</tbody></table>';
            }}
            document.getElementById('materiaisModalTitle').textContent = title;
            document.getElementById('materiaisModalBody').innerHTML = html;
            document.getElementById('materiaisModal').style.display = 'flex';
        }}
    </script>
    """


def generate_ordens_compra_html(analise: Dict, organizacao_selecionada: str = None) -> str:
    """Gera HTML da página de Ordens de Compra."""
    
    def fmt_currency(val):
        try:
            if val is None or pd.isna(val):
                return "R$ 0,00"
            return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return "R$ 0,00"
    
    def fmt_numero(val):
        try:
            if val is None or pd.isna(val):
                return "0"
            return f"{float(val):,.0f}".replace(",", ".")
        except:
            return "0"
    
    if not analise:
        return """
        <div class="header">
            <h1><i class="fas fa-shopping-cart"></i> Ordens de Compra</h1>
        </div>
        <div style="padding: 40px; text-align: center;">
            <i class="fas fa-inbox" style="font-size: 72px; color: #ccc; margin-bottom: 20px;"></i>
            <h2>Nenhum dado de ordens de compra carregado</h2>
            <p style="color: #666;">Verifique se o arquivo está no caminho correto.</p>
        </div>
        """
    
    # Opções de organizações
    organizacoes = analise.get('organizacoes', [])
    org_options = ""
    for org in organizacoes:
        selected = "selected" if org == organizacao_selecionada else ""
        org_options += f'<option value="{org}" {selected}>{org}</option>'
    
    # Top 10 Itens
    top10_rows = ""
    for idx, item in enumerate(analise.get('top10_itens', []), 1):
        descricao = item.get('Descrição', 'N/A')
        if descricao and len(str(descricao)) > 45:
            descricao = str(descricao)[:45] + '...'
        
        badge_color = '#f56565' if idx <= 3 else ('#ed8936' if idx <= 5 else '#2d3748')
        top10_rows += f"""
        <tr>
            <td><span style="display: inline-flex; align-items: center; justify-content: center; width: 30px; height: 30px; border-radius: 50%; background: {badge_color}; color: white; font-weight: 700; font-size: 0.85rem;">#{idx}</span></td>
            <td><div class="codigo-badge">{item.get('Código', 'N/A')}</div></td>
            <td>{descricao}</td>
            <td style="text-align: right;">{fmt_numero(item.get('Quantidade', 0))}</td>
            <td>{item.get('Organização', 'N/A')}</td>
            <td style="text-align: right;"><strong>{fmt_currency(item.get('Valor Total', 0))}</strong></td>
        </tr>
        """
    
    # Top 10 Pedidos com expansão
    top10_pedidos_rows = ""
    for idx, pedido in enumerate(analise.get('top10_pedidos', []), 1):
        data = pedido.get('Data', '')
        if pd.notna(data):
            try:
                data = pd.to_datetime(data).strftime('%d/%m/%Y')
            except:
                data = str(data)[:10]
        else:
            data = '-'
        
        badge_color = '#48bb78' if idx <= 3 else ('#38b2ac' if idx <= 5 else '#4299e1')
        num_pedido = pedido.get('Nº Pedido', 'N/A')
        
        # Gerar linhas dos itens deste pedido
        itens_pedido = pedido.get('itens', [])
        itens_rows = ""
        for item in itens_pedido:
            descricao_item = str(item.get('Descrição', 'N/A'))[:60]
            if len(str(item.get('Descrição', ''))) > 60:
                descricao_item += '...'
            itens_rows += f"""
            <tr class="item-detail-row" style="background: #f8fafc;">
                <td style="padding-left: 50px; color: #718096;"><i class="fas fa-box-open" style="margin-right: 8px;"></i></td>
                <td style="font-size: 0.85rem;"><span class="codigo-badge" style="font-size: 0.75rem;">{item.get('Código', 'N/A')}</span></td>
                <td colspan="2" style="font-size: 0.85rem;">{descricao_item}</td>
                <td style="text-align: right; font-size: 0.85rem;">{fmt_numero(item.get('Quantidade', 0))}</td>
                <td style="text-align: right; font-size: 0.85rem;">{fmt_currency(item.get('Valor', 0))}</td>
            </tr>
            """
        
        top10_pedidos_rows += f"""
        <tr class="pedido-row" onclick="togglePedidoDetails('pedido-{idx}')" style="cursor: pointer;">
            <td><span style="display: inline-flex; align-items: center; justify-content: center; width: 30px; height: 30px; border-radius: 50%; background: {badge_color}; color: white; font-weight: 700; font-size: 0.85rem;">#{idx}</span></td>
            <td><strong>{num_pedido}</strong></td>
            <td>{data}</td>
            <td>{pedido.get('Organização', 'N/A')}</td>
            <td style="text-align: right;">{pedido.get('Qtd Itens', 0)}</td>
            <td style="text-align: right;">
                <strong>{fmt_currency(pedido.get('Valor Total', 0))}</strong>
                <i class="fas fa-chevron-down expand-icon" id="icon-pedido-{idx}" style="margin-left: 10px; transition: transform 0.3s;"></i>
            </td>
        </tr>
        <tbody id="pedido-{idx}" class="pedido-details" style="display: none;">
            {itens_rows}
        </tbody>
        """
    
    # Itens Repetidos no Mês
    itens_repetidos_rows = ""
    for item in analise.get('itens_repetidos', []):
        mes_formatado = item.get('Mês', 'N/A')
        try:
            ano, mes = mes_formatado.split('-')
            meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            mes_formatado = f"{meses[int(mes)-1]}/{ano}"
        except:
            pass
        descricao = str(item.get('Descrição', 'N/A'))[:45]
        if len(str(item.get('Descrição', ''))) > 45:
            descricao += '...'
        itens_repetidos_rows += f"""
        <tr>
            <td><span class="codigo-badge">{item.get('Código', 'N/A')}</span></td>
            <td title="{item.get('Descrição', '')}">{descricao}</td>
            <td style="text-align: center;"><span style="background: #fed7d7; color: #c53030; padding: 4px 12px; border-radius: 12px; font-weight: 600;">{item.get('Qtd Pedidos', 0)}x</span></td>
            <td style="text-align: right;">{fmt_numero(item.get('Qtd Total', 0))}</td>
            <td style="text-align: right;">{fmt_currency(item.get('Valor Total', 0))}</td>
            <td style="text-align: center;"><span style="background: #e2e8f0; padding: 4px 10px; border-radius: 8px; font-size: 0.85rem;">{mes_formatado}</span></td>
        </tr>
        """
    
    # Por Organização
    org_rows = ""
    for org in analise.get('por_organizacao', []):
        org_rows += f"""
        <tr>
            <td><strong>{org.get('Organização', 'N/A')}</strong></td>
            <td style="text-align: right;">{fmt_currency(org.get('Valor Total', 0))}</td>
            <td style="text-align: right;">{org.get('Qtd Pedidos', 0)}</td>
            <td style="text-align: right;">{org.get('Qtd Itens', 0)}</td>
        </tr>
        """
    
    # Por Tipo de Transação
    tipo_rows = ""
    for tipo in analise.get('por_tipo', []):
        tipo_rows += f"""
        <tr>
            <td><strong>{tipo.get('Tipo', 'N/A')}</strong></td>
            <td style="text-align: right;">{fmt_currency(tipo.get('Valor Total', 0))}</td>
            <td style="text-align: right;">{tipo.get('Qtd Pedidos', 0)}</td>
        </tr>
        """
    
    # Dados para gráfico de tendência
    tendencia = analise.get('tendencia', [])
    tendencia_labels = [t['Mês'] for t in tendencia]
    tendencia_valores = [t['Valor'] for t in tendencia]
    tendencia_pedidos = [t['Qtd Pedidos'] for t in tendencia]
    
    # Dados para gráfico de organizações
    org_data = analise.get('por_organizacao', [])
    org_labels = [o['Organização'] for o in org_data]
    org_valores = [o['Valor Total'] for o in org_data]
    
    return f"""
    <div class="header animate-fade-in">
        <h1><i class="fas fa-shopping-cart"></i> Ordens de Compra</h1>
        <p class="subtitle">Análise de ordens de compra por organização e itens</p>
    </div>
    
    <div class="filter-bar animate-fade-in">
        <select class="filter-select" id="orgSelect" onchange="filtrarOrganizacao()">
            <option value="">Todas as Organizações</option>
            {org_options}
        </select>
        <button class="export-btn" onclick="exportReport()">
            <i class="fas fa-file-pdf"></i> Exportar PDF
        </button>
    </div>
    
    <!-- KPIs Principais -->
    <div class="kpis-grid animate-fade-in">
        <div class="stat-card highlight">
            <div class="stat-header">
                <span class="stat-icon"><i class='fas fa-dollar-sign'></i></span>
                <h3>Valor Total</h3>
            </div>
            <div class="stat-value">{fmt_currency(analise.get('valor_total', 0))}</div>
            <div class="stat-label">Total em ordens de compra</div>
        </div>
        
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-icon"><i class='fas fa-clipboard-list'></i></span>
                <h3>Qtd de Pedidos</h3>
            </div>
            <div class="stat-value">{fmt_numero(analise.get('qtd_pedidos', 0))}</div>
            <div class="stat-label">Pedidos únicos</div>
        </div>
        
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-icon"><i class='fas fa-box'></i></span>
                <h3>Qtd de Itens</h3>
            </div>
            <div class="stat-value">{fmt_numero(analise.get('qtd_itens', 0))}</div>
            <div class="stat-label">Itens diferentes</div>
        </div>
        
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-icon"><i class='fas fa-chart-bar'></i></span>
                <h3>Valor Médio/Pedido</h3>
            </div>
            <div class="stat-value">{fmt_currency(analise.get('valor_medio_pedido', 0))}</div>
            <div class="stat-label">Ticket médio</div>
        </div>
    </div>
    
    <!-- Gráficos -->
    <div class="grid-2 animate-fade-in" style="margin-top: 24px;">
        <div class="chart-card">
            <h3><i class="fas fa-chart-line"></i> Evolução Mensal</h3>
            <div style="height: 350px; margin-top: 16px;">
                <canvas id="chartTendencia"></canvas>
            </div>
        </div>
        
        <div class="chart-card">
            <h3><i class="fas fa-building"></i> Valor por Organização</h3>
            <div style="height: 350px; margin-top: 16px;">
                <canvas id="chartOrganizacoes"></canvas>
            </div>
        </div>
    </div>
    
    <!-- Top 10 Itens -->
    <div class="chart-card full-width animate-fade-in" style="margin-top: 24px;">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-trophy"></i> Top 10 Itens Mais Caros</div>
        </div>
        <table class="data-table">
            <thead>
                <tr>
                    <th style="width: 60px;">#</th>
                    <th style="width: 140px;">Código</th>
                    <th>Descrição</th>
                    <th style="width: 100px; text-align: right;">Qtd</th>
                    <th style="width: 150px;">Organização</th>
                    <th style="width: 140px; text-align: right;">Valor Total</th>
                </tr>
            </thead>
            <tbody>
                {top10_rows}
            </tbody>
        </table>
    </div>
    
    <!-- Top 10 Pedidos -->
    <div class="chart-card full-width animate-fade-in" style="margin-top: 24px;">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-file-invoice-dollar"></i> Top 10 Pedidos por Valor <span style="font-weight: 400; font-size: 0.85rem; color: #718096; margin-left: 10px;">(clique para expandir)</span></div>
        </div>
        <table class="data-table">
            <thead>
                <tr>
                    <th style="width: 60px;">#</th>
                    <th style="width: 180px;">Nº Pedido</th>
                    <th style="width: 120px;">Data</th>
                    <th>Organização</th>
                    <th style="width: 100px; text-align: right;">Qtd Itens</th>
                    <th style="width: 140px; text-align: right;">Valor Total</th>
                </tr>
            </thead>
            <tbody>
                {top10_pedidos_rows}
            </tbody>
        </table>
    </div>
    
    <!-- Itens Repetidos no Mês -->
    <div class="chart-card full-width animate-fade-in" style="margin-top: 24px;">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-redo-alt"></i> Itens Comprados Múltiplas Vezes no Mesmo Mês</div>
            <span style="font-size: 0.85rem; color: #718096;">Itens que aparecem em mais de 1 pedido no mesmo mês</span>
        </div>
        <table class="data-table">
            <thead>
                <tr>
                    <th style="width: 140px;">Código</th>
                    <th>Descrição</th>
                    <th style="width: 100px; text-align: center;">Repetições</th>
                    <th style="width: 100px; text-align: right;">Qtd Total</th>
                    <th style="width: 140px; text-align: right;">Valor Total</th>
                    <th style="width: 100px; text-align: center;">Mês</th>
                </tr>
            </thead>
            <tbody>
                {itens_repetidos_rows}
            </tbody>
        </table>
    </div>
    
    <!-- Análise por Organização e Tipo -->
    <div class="grid-2 animate-fade-in" style="margin-top: 24px;">
        <div class="chart-card">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-building"></i> Resumo por Organização</div>
            </div>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Organização</th>
                        <th style="text-align: right;">Valor Total</th>
                        <th style="text-align: right;">Pedidos</th>
                        <th style="text-align: right;">Itens</th>
                    </tr>
                </thead>
                <tbody>
                    {org_rows}
                </tbody>
            </table>
        </div>
        
        <div class="chart-card">
            <div class="chart-header">
                <div class="chart-title"><i class="fas fa-tags"></i> Por Tipo de Transação</div>
            </div>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Tipo</th>
                        <th style="text-align: right;">Valor Total</th>
                        <th style="text-align: right;">Pedidos</th>
                    </tr>
                </thead>
                <tbody>
                    {tipo_rows}
                </tbody>
            </table>
        </div>
    </div>
    
    <style>
        .kpis-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 20px;
        }}
        
        .stat-card {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
            transition: all 0.3s ease;
        }}
        
        .stat-card:hover {{
            transform: translateY(-4px);
            box-shadow: var(--shadow-lg);
        }}
        
        .stat-card.highlight {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
        }}
        
        .stat-card.highlight h3,
        .stat-card.highlight .stat-label {{
            color: rgba(255,255,255,0.9);
        }}
        
        .stat-header {{
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 12px;
        }}
        
        .stat-icon {{
            font-size: 28px;
        }}
        
        .stat-header h3 {{
            margin: 0;
            font-size: 1rem;
            font-weight: 600;
            color: #64748b;
        }}
        
        .stat-value {{
            font-size: 1.8rem;
            font-weight: 700;
            color: var(--dark);
        }}
        
        .stat-card.highlight .stat-value {{
            color: white;
        }}
        
        .stat-label {{
            font-size: 0.85rem;
            color: #94a3b8;
            margin-top: 4px;
        }}
        
        .grid-2 {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(450px, 1fr));
            gap: 24px;
        }}
        
        .chart-card {{
            background: white;
            border-radius: var(--radius);
            padding: 24px;
            box-shadow: var(--shadow-md);
        }}
        
        .chart-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 2px solid #f0f0f0;
        }}
        
        .chart-title {{
            font-size: 1.1rem;
            font-weight: 600;
            color: var(--dark);
        }}
        
        .chart-title i {{
            color: var(--primary);
            margin-right: 8px;
        }}
        
        .full-width {{
            grid-column: 1 / -1;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }}
        
        .data-table thead tr {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
        }}
        
        .data-table th {{
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.8rem;
            letter-spacing: 0.5px;
        }}
        
        .data-table td {{
            padding: 14px 12px;
            border-bottom: 1px solid #e8e8e8;
        }}
        
        .data-table tbody tr:hover {{
            background: #f8f9fa;
        }}
        
        .data-table tbody tr:nth-child(even) {{
            background: #fafafa;
        }}
        
        .data-table tbody tr:nth-child(even):hover {{
            background: #f0f0f0;
        }}
        
        .pedido-row:hover {{
            background: #eef2ff !important;
        }}
        
        .pedido-row td:last-child {{
            display: flex;
            align-items: center;
            justify-content: flex-end;
        }}
        
        .expand-icon {{
            color: #718096;
            font-size: 0.85rem;
        }}
        
        .item-detail-row td {{
            padding: 10px 12px !important;
            border-bottom: 1px dashed #e2e8f0 !important;
        }}
        
        .codigo-badge {{
            display: inline-block;
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
            padding: 4px 10px;
            border-radius: 6px;
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            font-weight: 600;
        }}
    </style>
    
    <script>
        const tendenciaLabels = {json.dumps(tendencia_labels)};
        const tendenciaValores = {json.dumps(tendencia_valores)};
        const orgLabels = {json.dumps(org_labels)};
        const orgValores = {json.dumps(org_valores)};
        
        // Função para expandir/colapsar detalhes do pedido
        function togglePedidoDetails(id) {{
            const details = document.getElementById(id);
            const icon = document.getElementById('icon-' + id);
            if (details.style.display === 'none') {{
                details.style.display = 'table-row-group';
                icon.style.transform = 'rotate(180deg)';
            }} else {{
                details.style.display = 'none';
                icon.style.transform = 'rotate(0deg)';
            }}
        }}
        
        // Chart: Tendência Mensal
        new Chart(document.getElementById('chartTendencia'), {{
            type: 'line',
            data: {{
                labels: tendenciaLabels,
                datasets: [{{
                    label: 'Valor Total',
                    data: tendenciaValores,
                    borderColor: '#2d3748',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointBackgroundColor: '#2d3748',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointRadius: 5
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        ticks: {{
                            callback: function(value) {{
                                return 'R$ ' + value.toLocaleString('pt-BR');
                            }}
                        }}
                    }}
                }}
            }}
        }});
        
        // Chart: Organizações
        new Chart(document.getElementById('chartOrganizacoes'), {{
            type: 'doughnut',
            data: {{
                labels: orgLabels,
                datasets: [{{
                    data: orgValores,
                    backgroundColor: [
                        '#2d3748', '#1a202c', '#f56565', '#ed8936', '#48bb78',
                        '#4299e1', '#9f7aea', '#38b2ac', '#f6ad55', '#fc8181'
                    ],
                    borderWidth: 2,
                    borderColor: '#fff'
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        position: 'right'
                    }}
                }}
            }}
        }});
        
        function filtrarOrganizacao() {{
            const org = document.getElementById('orgSelect').value;
            if (window.pyBridge) {{
                window.pyBridge.filtrarOrganizacaoOrdens(org);
            }}
        }}
        
        function exportReport() {{
            if (window.pyBridge) window.pyBridge.exportReport();
        }}
    </script>
    """


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
    
    # Curva ABC - TODOS os fornecedores com paginação
    curva_abc_html = ""
    total_fornecedores_abc = 0
    if not tabela_abc.empty:
        total_fornecedores_abc = len(tabela_abc)
        for idx, row in tabela_abc.iterrows():
            classe_badge = f"<span class='badge-{row['Classe'].lower()}'>{row['Classe']}</span>"
            curva_abc_html += f"""
            <tr class='curva-abc-row'>
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
                <div class="kpi-icon" style="background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);">
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
                    <span style="margin-left: 20px; color: #666;">
                        <i class="fas fa-info-circle"></i> Total: {total_fornecedores_abc} fornecedores
                    </span>
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
                
                <!-- Controles de Paginação -->
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; padding: 12px; background: #f7fafc; border-radius: 8px;">
                    <div style="color: #666; font-size: 14px;">
                        Mostrando <span id="abc-showing-start">1</span> - <span id="abc-showing-end">10</span> de <span id="abc-total">{total_fornecedores_abc}</span> fornecedores
                    </div>
                    <div style="display: flex; gap: 8px;">
                        <button id="abc-prev-btn" onclick="changeAbcPage(-1)" style="padding: 8px 16px; border: 1px solid #cbd5e0; background: white; border-radius: 6px; cursor: pointer; font-size: 14px;">
                            <i class="fas fa-chevron-left"></i> Anterior
                        </button>
                        <span style="padding: 8px 16px; background: white; border: 1px solid #cbd5e0; border-radius: 6px; font-size: 14px;">
                            Página <span id="abc-current-page">1</span> de <span id="abc-total-pages">1</span>
                        </span>
                        <button id="abc-next-btn" onclick="changeAbcPage(1)" style="padding: 8px 16px; border: 1px solid #cbd5e0; background: white; border-radius: 6px; cursor: pointer; font-size: 14px;">
                            Próxima <i class="fas fa-chevron-right"></i>
                        </button>
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
                    <tbody id="curva-abc-tbody">
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
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
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
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
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
        
        // ========== PAGINAÇÃO DA CURVA ABC ==========
        let abcCurrentPage = 1;
        const abcItemsPerPage = 10;
        let abcAllRows = [];
        
        // Inicializar paginação quando a página carregar
        window.addEventListener('DOMContentLoaded', function() {{
            initAbcPagination();
        }});
        
        function initAbcPagination() {{
            const tbody = document.getElementById('curva-abc-tbody');
            if (!tbody) return;
            
            abcAllRows = Array.from(tbody.getElementsByClassName('curva-abc-row'));
            const totalRows = abcAllRows.length;
            
            if (totalRows === 0) return;
            
            const totalPages = Math.ceil(totalRows / abcItemsPerPage);
            document.getElementById('abc-total').textContent = totalRows;
            document.getElementById('abc-total-pages').textContent = totalPages;
            
            showAbcPage(1);
        }}
        
        function showAbcPage(page) {{
            const totalPages = Math.ceil(abcAllRows.length / abcItemsPerPage);
            
            // Validar página
            if (page < 1) page = 1;
            if (page > totalPages) page = totalPages;
            
            abcCurrentPage = page;
            
            // Esconder todas as linhas
            abcAllRows.forEach(row => row.style.display = 'none');
            
            // Mostrar apenas as linhas da página atual
            const start = (page - 1) * abcItemsPerPage;
            const end = start + abcItemsPerPage;
            
            for (let i = start; i < end && i < abcAllRows.length; i++) {{
                abcAllRows[i].style.display = '';
            }}
            
            // Atualizar informações da paginação
            document.getElementById('abc-current-page').textContent = page;
            document.getElementById('abc-showing-start').textContent = start + 1;
            document.getElementById('abc-showing-end').textContent = Math.min(end, abcAllRows.length);
            
            // Habilitar/desabilitar botões
            const prevBtn = document.getElementById('abc-prev-btn');
            const nextBtn = document.getElementById('abc-next-btn');
            
            prevBtn.disabled = page === 1;
            nextBtn.disabled = page === totalPages;
            
            prevBtn.style.opacity = page === 1 ? '0.5' : '1';
            nextBtn.style.opacity = page === totalPages ? '0.5' : '1';
            prevBtn.style.cursor = page === 1 ? 'not-allowed' : 'pointer';
            nextBtn.style.cursor = page === totalPages ? 'not-allowed' : 'pointer';
        }}
        
        function changeAbcPage(delta) {{
            showAbcPage(abcCurrentPage + delta);
        }}
    </script>
    </body>
    </html>
    """
    
    return html


def generate_curva_abc_page_html(curva_abc: Dict, estoque_org: Dict, organizacao_selecionada: str = None) -> str:
    """Gera HTML da página dedicada à Curva ABC em formato de tabela."""
    
    def fmt_currency(val):
        try:
            if val is None or pd.isna(val):
                return "R$ 0,00"
            val_float = float(val)
            return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"
    
    # Dados das classes
    abc_classe_a = curva_abc.get('classe_a', {})
    abc_classe_b = curva_abc.get('classe_b', {})
    abc_classe_c = curva_abc.get('classe_c', {})
    itens_classe_a = curva_abc.get('itens_classe_a', [])
    itens_classe_b = curva_abc.get('itens_classe_b', [])
    itens_classe_c = curva_abc.get('itens_classe_c', [])
    
    # Opções para select de organização
    organizacoes_options = ""
    for org in estoque_org.get('organizacoes', []):
        selected = 'selected' if org == organizacao_selecionada else ''
        organizacoes_options += f'<option value="{org}" {selected}>{org}</option>'
    
    # Gerar linhas da tabela para Classe A
    linhas_classe_a = ""
    for idx, item in enumerate(itens_classe_a, 1):
        descricao = item.get('Descrição do Item', 'Sem descrição')
        if descricao and len(str(descricao)) > 80:
            descricao = str(descricao)[:80] + '...'
        
        linhas_classe_a += f"""
        <tr class="abc-row classe-a-row">
            <td class="rank-col"><span class="rank-badge-a">#{idx}</span></td>
            <td class="codigo-col"><div class="codigo-badge">{item.get('Nome do Item', 'N/A')}</div></td>
            <td class="descricao-col">{descricao}</td>
            <td class="valor-col">{fmt_currency(item.get('Custo Total', 0))}</td>
            <td class="percent-col">{item.get('% Representatividade', 0):.2f}%</td>
            <td class="percent-col"><strong>{item.get('% Acumulado', 0):.2f}%</strong></td>
        </tr>
        """
    
    # Gerar linhas da tabela para Classe B
    linhas_classe_b = ""
    for idx, item in enumerate(itens_classe_b, 1):
        descricao = item.get('Descrição do Item', 'Sem descrição')
        if descricao and len(str(descricao)) > 80:
            descricao = str(descricao)[:80] + '...'
        
        linhas_classe_b += f"""
        <tr class="abc-row classe-b-row">
            <td class="rank-col"><span class="rank-badge-b">#{idx}</span></td>
            <td class="codigo-col"><div class="codigo-badge">{item.get('Nome do Item', 'N/A')}</div></td>
            <td class="descricao-col">{descricao}</td>
            <td class="valor-col">{fmt_currency(item.get('Custo Total', 0))}</td>
            <td class="percent-col">{item.get('% Representatividade', 0):.2f}%</td>
            <td class="percent-col"><strong>{item.get('% Acumulado', 0):.2f}%</strong></td>
        </tr>
        """
    
    # Gerar linhas da tabela para Classe C
    linhas_classe_c = ""
    for idx, item in enumerate(itens_classe_c, 1):
        descricao = item.get('Descrição do Item', 'Sem descrição')
        if descricao and len(str(descricao)) > 80:
            descricao = str(descricao)[:80] + '...'
        
        linhas_classe_c += f"""
        <tr class="abc-row classe-c-row">
            <td class="rank-col"><span class="rank-badge-c">#{idx}</span></td>
            <td class="codigo-col"><div class="codigo-badge">{item.get('Nome do Item', 'N/A')}</div></td>
            <td class="descricao-col">{descricao}</td>
            <td class="valor-col">{fmt_currency(item.get('Custo Total', 0))}</td>
            <td class="percent-col">{item.get('% Representatividade', 0):.2f}%</td>
            <td class="percent-col"><strong>{item.get('% Acumulado', 0):.2f}%</strong></td>
        </tr>
        """
    
    total_valor = abc_classe_a.get('valor', 0) + abc_classe_b.get('valor', 0) + abc_classe_c.get('valor', 0)
    total_itens = abc_classe_a.get('qtd', 0) + abc_classe_b.get('qtd', 0) + abc_classe_c.get('qtd', 0)
    
    return f"""
    <div class="header animate-fade-in">
        <h1><i class="fas fa-chart-pie"></i> Curva ABC - Análise de Estoque</h1>
        <p class="subtitle">Classificação de itens por valor acumulado • Total: {fmt_currency(total_valor)} | {total_itens} itens</p>
    </div>
    
    <div class="filter-bar animate-fade-in">
        <select class="filter-select" id="orgSelect" onchange="filtrarOrganizacao()">
            <option value="">Todas as Organizações</option>
            {organizacoes_options}
        </select>
        <button class="export-btn" onclick="exportReport()">
            <i class="fas fa-file-excel"></i> Exportar
        </button>
    </div>
    
    <div class="kpis-grid animate-fade-in">
        <div class="stat-card class-a-card">
            <div class="stat-header">
                <span class="stat-icon">🔴</span>
                <h3>Classe A</h3>
            </div>
            <div class="stat-value">{fmt_currency(abc_classe_a.get('valor', 0))}</div>
            <div class="stat-label">{abc_classe_a.get('qtd', 0)} itens • 80% do valor acumulado</div>
        </div>
        <div class="stat-card class-b-card">
            <div class="stat-header">
                <span class="stat-icon">🟡</span>
                <h3>Classe B</h3>
            </div>
            <div class="stat-value">{fmt_currency(abc_classe_b.get('valor', 0))}</div>
            <div class="stat-label">{abc_classe_b.get('qtd', 0)} itens • 15% do valor acumulado</div>
        </div>
        <div class="stat-card class-c-card">
            <div class="stat-header">
                <span class="stat-icon">🟢</span>
                <h3>Classe C</h3>
            </div>
            <div class="stat-value">{fmt_currency(abc_classe_c.get('valor', 0))}</div>
            <div class="stat-label">{abc_classe_c.get('qtd', 0)} itens • 5% do valor acumulado</div>
        </div>
    </div>
    
    <div class="chart-card full-width animate-fade-in" style="margin: 24px;">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-table"></i> Classe A - Itens Críticos (80% do Valor)</div>
            <div style="color: #718096; font-size: 0.9rem;">{abc_classe_a.get('qtd', 0)} itens</div>
        </div>
        <div style="overflow-x: auto;">
            <table class="abc-table">
                <thead>
                    <tr>
                        <th style="width: 60px;">#</th>
                        <th style="width: 140px;">Código</th>
                        <th style="min-width: 300px;">Descrição</th>
                        <th style="width: 140px;">Valor Total</th>
                        <th style="width: 120px;">% Estoque</th>
                        <th style="width: 120px;">% Acumulado</th>
                    </tr>
                </thead>
                <tbody>
                    {linhas_classe_a}
                </tbody>
            </table>
        </div>
    </div>
    
    <div class="chart-card full-width animate-fade-in" style="margin: 24px;">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-table"></i> Classe B - Itens Intermediários (15% do Valor)</div>
            <div style="color: #718096; font-size: 0.9rem;">{abc_classe_b.get('qtd', 0)} itens</div>
        </div>
        <div style="overflow-x: auto;">
            <table class="abc-table">
                <thead>
                    <tr>
                        <th style="width: 60px;">#</th>
                        <th style="width: 140px;">Código</th>
                        <th style="min-width: 300px;">Descrição</th>
                        <th style="width: 140px;">Valor Total</th>
                        <th style="width: 120px;">% Estoque</th>
                        <th style="width: 120px;">% Acumulado</th>
                    </tr>
                </thead>
                <tbody>
                    {linhas_classe_b}
                </tbody>
            </table>
        </div>
    </div>
    
    <div class="chart-card full-width animate-fade-in" style="margin: 24px;">
        <div class="chart-header">
            <div class="chart-title"><i class="fas fa-table"></i> Classe C - Itens de Baixo Valor (5% do Valor)</div>
            <div style="color: #718096; font-size: 0.9rem;">{abc_classe_c.get('qtd', 0)} itens</div>
        </div>
        <div style="overflow-x: auto;">
            <table class="abc-table">
                <thead>
                    <tr>
                        <th style="width: 60px;">#</th>
                        <th style="width: 140px;">Código</th>
                        <th style="min-width: 300px;">Descrição</th>
                        <th style="width: 140px;">Valor Total</th>
                        <th style="width: 120px;">% Estoque</th>
                        <th style="width: 120px;">% Acumulado</th>
                    </tr>
                </thead>
                <tbody>
                    {linhas_classe_c}
                </tbody>
            </table>
        </div>
    </div>
    
    <style>
        .abc-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
        }}
        
        .abc-table thead {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
            color: white;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        
        .abc-table th {{
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
            font-size: 0.8rem;
            text-transform: uppercase;
            white-space: nowrap;
        }}
        
        .abc-table td {{
            padding: 14px 12px;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: middle;
        }}
        
        .abc-row {{
            transition: all 0.2s ease;
        }}
        
        .abc-row:hover {{
            transform: scale(1.005);
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}
        
        .classe-a-row:hover {{
            background: #fff5f5;
        }}
        
        .classe-b-row:hover {{
            background: #fffaf0;
        }}
        
        .classe-c-row:hover {{
            background: #f0fff4;
        }}
        
        .rank-col {{
            text-align: center;
        }}
        
        .rank-badge-a {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background: #f56565;
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
            box-shadow: 0 2px 6px rgba(245, 101, 101, 0.3);
        }}
        
        .rank-badge-b {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background: #ed8936;
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
            box-shadow: 0 2px 6px rgba(237, 137, 54, 0.3);
        }}
        
        .rank-badge-c {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background: #48bb78;
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
            box-shadow: 0 2px 6px rgba(72, 187, 120, 0.3);
        }}
        
        .codigo-badge {{
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            font-weight: 700;
            color: var(--primary);
            background: #f0f4ff;
            padding: 6px 10px;
            border-radius: 6px;
            display: inline-block;
        }}
        
        .descricao-col {{
            color: #2d3748;
            line-height: 1.4;
        }}
        
        .valor-col {{
            text-align: right;
            font-weight: 600;
            color: var(--dark);
        }}
        
        .percent-col {{
            text-align: right;
            color: var(--primary);
        }}
        
        .class-a-card {{ border-left: 4px solid #f56565; }}
        .class-b-card {{ border-left: 4px solid #ed8936; }}
        .class-c-card {{ border-left: 4px solid #48bb78; }}
    </style>
    
    <script>
        function filtrarOrganizacao() {{
            const org = document.getElementById('orgSelect').value;
            if (window.pyBridge) {{
                window.pyBridge.filtrarOrganizacao(org);
            }}
        }}
        
        function exportReport() {{
            if (window.pyBridge) window.pyBridge.exportReport();
        }}
    </script>
    """


# ═══════════════════════════════════════════════════════════════════════════════
# JANELA PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    """Janela principal da aplicação."""
    
    def __init__(self, loaded_data=None, user_email=None, user_password=None):
        super().__init__()
        self.data_handler = DataHandler()
        self.current_view = "dashboard"
        self.organizacao_selecionada = None
        
        # Credenciais do usuário para envio de email
        self.user_email = user_email
        self.user_password = user_password
        
        # Pipefy - estrutura para múltiplos pipes
        self.pipefy_client = None
        self.pipefy_data = {
            'compras_servicos': {'cards': [], 'name': '', 'loading': False},
            'compras_csc': {'cards': [], 'name': '', 'loading': False},
            'compras_locais': {'cards': [], 'name': '', 'loading': False},
            'envio_nfe': {'cards': [], 'name': '', 'loading': False},
            'reserva_materiais': {'cards': [], 'name': '', 'loading': False},
        }
        self.pipefy_pipes = {
            'compras_servicos': PIPE_CONTRATACAO_SERVICOS,
            'compras_csc': PIPE_COMPRAS_CSC,
            'compras_locais': PIPE_COMPRAS_LOCAIS,
            'envio_nfe': PIPE_ENVIO_NFE,
            'reserva_materiais': PIPE_RESERVA_MATERIAIS,
        }
        self.pipefy_titulos = {
            'compras_servicos': 'Contratação de Serviços',
            'compras_csc': 'Compras CSC',
            'compras_locais': 'Compras Locais',
            'envio_nfe': 'Envio de NFE',
            'reserva_materiais': 'Reserva de Materiais',
        }
        # Compatibilidade retroativa
        self.pipefy_cards = []
        self.pipefy_pipe_name = ""
        self.pipefy_loading = False
        self.current_pipefy_category = 'compras_servicos'
        
        # Aplicar dados pré-carregados se disponíveis
        if loaded_data:
            self.apply_loaded_data(loaded_data)
        
        self.setup_ui()
        self._setup_file_monitoring()  # Configurar monitoramento de arquivos
        self.load_dashboard()
    
    def apply_loaded_data(self, loaded_data):
        """Aplica os dados pré-carregados na janela principal."""
        # Aplicar cliente Pipefy
        if 'pipefy_client' in loaded_data:
            self.pipefy_client = loaded_data['pipefy_client']
        
        # Aplicar dados dos pipes
        if 'pipefy_data' in loaded_data:
            for category, data in loaded_data['pipefy_data'].items():
                if category in self.pipefy_data:
                    self.pipefy_data[category] = data
        
        # Atualizar dados da categoria atual para compatibilidade retroativa
        current_data = self.pipefy_data.get(self.current_pipefy_category, {})
        self.pipefy_cards = current_data.get('cards', [])
        self.pipefy_pipe_name = current_data.get('name', '')
    
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
        
        # Sidebar
        sidebar = QWidget()
        sidebar.setStyleSheet("""
            QWidget { background: qlineargradient(y1:0, y2:1, stop:0 #1a1a2e, stop:1 #16213e); min-width: 260px; max-width: 260px; }
            QPushButton { background: transparent; color: rgba(255,255,255,0.7); border: none; border-radius: 8px; padding: 14px 20px; text-align: left; }
            QPushButton:hover { background: rgba(255,255,255,0.1); color: white; }
            QPushButton:checked { background: qlineargradient(x1:0, x2:1, stop:0 #2d3748, stop:1 #1a202c); color: white; }
            QLabel { color: white; padding: 20px; }
        """)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        
        title = QLabel("<i class='fas fa-chart-bar'></i> Ferramenta de Gestão")
        title.setStyleSheet("QLabel { font-size: 18px; font-weight: bold; }")
        sidebar_layout.addWidget(title)
        
        subtitle = QLabel("v2.0 - Sistema Inteligente")
        subtitle.setStyleSheet("QLabel { font-size: 10px; color: rgba(255,255,255,0.5); margin: -15px 20px 10px; }")
        sidebar_layout.addWidget(subtitle)
        
        # Botões de navegação
        self.nav_buttons = {}
        nav_items = [
            ("dashboard", "Dashboard", self.load_dashboard),
            ("estoque", "Estoque", self.load_estoque),
            ("ordens", "Ordens de Compra", self.load_ordens),
            ("movimentacoes", "Movimentacoes", self.load_movimentacoes),
            ("pagamentos", "Pagamentos", self.load_pagamentos),
            ("atividades", "Atividades", self.load_atividades),
        ]
        
        for key, text, callback in nav_items:
            btn = QPushButton(text)
            btn.setObjectName(key)
            btn.setCheckable(True)
            btn.clicked.connect(callback)
            self.nav_buttons[key] = btn
            sidebar_layout.addWidget(btn)
        
        sidebar_layout.addSpacing(10)
        
        # Botões de ação
        btn_export = QPushButton("📊 Exportar PDF")
        btn_export.clicked.connect(self.export_report)
        sidebar_layout.addWidget(btn_export)
        
        sidebar_layout.addStretch()
        
        # Indicador de monitoramento
        self.status_monitoring = QLabel("🟢 Monitoramento Ativo")
        self.status_monitoring.setStyleSheet("QLabel { color: #48bb78; padding: 10px 20px; font-size: 11px; background: rgba(72, 187, 120, 0.1); border-radius: 6px; margin: 10px; }")
        self.status_monitoring.setToolTip("Sistema monitorando automaticamente as pastas de relatórios")
        sidebar_layout.addWidget(self.status_monitoring)
        
        version = QLabel(f"v{APP_VERSION}")
        version.setStyleSheet("QLabel { color: rgba(255,255,255,0.3); padding: 20px; font-size: 10px; }")
        sidebar_layout.addWidget(version)
        
        main_layout.addWidget(sidebar)
        
        # WebView
        self.web_view = QWebEngineView()
        self.web_channel = QWebChannel()
        self.bridge = WebBridge()
        self.web_channel.registerObject('pyBridge', self.bridge)
        self.web_view.page().setWebChannel(self.web_channel)
        
        self.bridge.showSupplierDetailSignal.connect(self.show_supplier_detail)
        self.bridge.exportReportSignal.connect(self.export_report)
        self.bridge.goBackSignal.connect(self.load_dashboard)
        self.bridge.filtrarOrganizacaoSignal.connect(self.filtrar_organizacao_estoque)
        self.bridge.importarEstoqueSignal.connect(self.import_data)
        self.bridge.mostrarMateriaisGrupoSignal.connect(self.mostrar_lista_materiais_grupo)
        self.bridge.filtrarOrganizacaoOrdensSignal.connect(self.filtrar_organizacao_ordens)
        self.bridge.filtrarOrganizacaoMovimentacoesSignal.connect(self.filtrar_organizacao_movimentacoes)
        self.bridge.loadAtividadeSubmenuSignal.connect(self.load_atividades)
        # Pipefy signals
        self.bridge.carregarDadosPipefySignal.connect(self.carregar_dados_pipefy)
        self.bridge.exportarPdfPipefySignal.connect(self.exportar_pdf_pipefy)
        self.bridge.atualizarDadosPipefySignal.connect(self.atualizar_dados_pipefy)
        self.bridge.enviarEmailPipefySignal.connect(self.enviar_email_pipefy)
        
        main_layout.addWidget(self.web_view, 1)
        
        # Variável para armazenar organização selecionada
        self.organizacao_selecionada = None
    
    def _setup_file_monitoring(self):
        """Configura monitoramento automático das pastas de relatórios."""
        self.file_watcher = QFileSystemWatcher()
        
        # Obter mês e ano atual
        mes, ano = self.data_handler._get_mes_ano_atual()
        base_path = self.data_handler.base_relatorios
        
        # Pastas a monitorar
        pastas_monitorar = [
            os.path.join(base_path, 'Pagamentos', ano, mes),
            os.path.join(base_path, 'Estoque periodico', ano, mes),
            os.path.join(base_path, 'Geral contabil', ano, mes),
        ]
        
        # Adicionar pastas existentes ao watcher
        for pasta in pastas_monitorar:
            if os.path.exists(pasta):
                self.file_watcher.addPath(pasta)
                print(f"📁 Monitorando: {pasta}")
        
        # Conectar sinais
        self.file_watcher.directoryChanged.connect(self._on_directory_changed)
        
        print("✓ Sistema de monitoramento automático ativado!")
    
    def _on_directory_changed(self, path: str):
        """Callback quando um arquivo é adicionado/modificado na pasta monitorada."""
        print(f"\\n🔄 Mudança detectada em: {path}")
        
        # Aguardar 1 segundo para o arquivo ser completamente copiado
        QTimer.singleShot(1000, lambda: self._recarregar_dados(path))
    
    def _recarregar_dados(self, path: str):
        """Recarrega os dados quando arquivos são adicionados."""
        print(f"🔃 Recarregando dados de: {path}")
        
        # Identificar tipo de relatório pela pasta
        if 'Pagamentos' in path:
            arquivo = self.data_handler._buscar_arquivo_na_pasta(path)
            if arquivo:
                print(f"✓ Novo arquivo de pagamentos detectado: {arquivo}")
                self.data_handler.load_payment_file(arquivo)
                self._atualizar_view_se_necessario()
        
        elif 'Estoque periodico' in path:
            arquivo = self.data_handler._buscar_arquivo_na_pasta(path)
            if arquivo:
                print(f"✓ Novo arquivo de estoque detectado: {arquivo}")
                self.data_handler.load_estoque_file(arquivo)
                self._atualizar_view_se_necessario()
        
        elif 'Geral contabil' in path:
            arquivo = self.data_handler._buscar_arquivo_na_pasta(path)
            if arquivo:
                print(f"✓ Novo arquivo de ordens detectado: {arquivo}")
                self.data_handler.load_ordens_compra_file(arquivo)
                self._atualizar_view_se_necessario()
    
    def _atualizar_view_se_necessario(self):
        """Atualiza a visualização atual após recarregar dados."""
        print(f"🔄 Atualizando visualização: {self.current_view}")
        
        # Recarregar a view atual
        if self.current_view == "dashboard":
            self.load_dashboard()
        elif self.current_view == "estoque":
            self.load_estoque()
        elif self.current_view == "ordens":
            self.load_ordens()
        elif self.current_view == "movimentacoes":
            self.load_movimentacoes()
        elif self.current_view == "pagamentos":
            self.load_pagamentos()
        
        # Mostrar notificação ao usuário
        QTimer.singleShot(500, lambda: QMessageBox.information(
            self, 
            "Dados Atualizados", 
            "Os dados foram atualizados automaticamente com os novos arquivos!",
            QMessageBox.Ok
        ))
    
    def _render_html(self, html_content: str):
        """Renderiza HTML no WebView."""
        base_html = get_base_html()
        full_html = base_html.replace(
            '<div class="dashboard-container" id="dashboard">\n        <!-- Conteúdo será injetado aqui -->\n    </div>',
            f'<div class="dashboard-container" id="dashboard">\n{html_content}\n    </div>'
        )
        self.web_view.setHtml(full_html)
    
    def _update_nav(self, active_key: str):
        """Atualiza botões de navegação."""
        for key, btn in self.nav_buttons.items():
            btn.setChecked(key == active_key)
    
    def load_dashboard(self):
        """Carrega dashboard principal."""
        self.current_view = "dashboard"
        self._update_nav("dashboard")
        
        # Passa a organização selecionada para get_kpis
        kpis = self.data_handler.get_kpis(self.organizacao_selecionada)
        ranking = self.data_handler.get_ranking_fornecedores()
        curva_abc = self.data_handler.get_curva_abc()
        evolucao = self.data_handler.get_evolucao_mensal()
        alertas = self.data_handler.get_alertas()
        evolucao_estoque = self.data_handler.get_evolucao_estoque_mensal()
        
        # Dados de estoque para o dashboard
        estoque_org = {'organizacoes': [], 'dados': []}
        if self.data_handler.df_estoque is not None and len(self.data_handler.df_estoque) > 0:
            estoque_org = self.data_handler.get_estoque_por_organizacao()
        
        html = generate_dashboard_html(kpis, ranking, curva_abc, evolucao, alertas, evolucao_estoque, estoque_org, self.organizacao_selecionada)
        self._render_html(html)
    
    def load_estoque(self):
        """Carrega página de estoque."""
        self.current_view = "estoque"
        self._update_nav("estoque")
        
        if self.data_handler.df_estoque is None or len(self.data_handler.df_estoque) == 0:
            html = """
            <div class='header'>
                <h1><i class='fas fa-box'></i> Gestão de Estoque</h1>
            </div>
            <div style='padding: 40px; text-align: center;'>
                <i class='fas fa-inbox' style='font-size: 72px; color: #ccc; margin-bottom: 20px;'></i>
                <h2>Nenhum dado de estoque carregado</h2>
                <p style='color: #666; margin: 20px 0;'>Importe o arquivo de estoque periódico para visualizar as análises.</p>
                <button onclick='importarEstoque()' style='padding: 12px 24px; background: #2d3748; color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 16px;'>
                    <i class='fas fa-upload'></i> Importar Estoque
                </button>
            </div>
            <script>
                function importarEstoque() {
                    if (window.pyBridge) window.pyBridge.importarEstoque();
                }
            </script>
            """
            self._render_html(html)
            return
        
        try:
            analise_grupos = self.data_handler.get_analise_estoque_por_grupo(self.organizacao_selecionada)
            estoque_stage = self.data_handler.get_estoque_stage()
            estoque_org = self.data_handler.get_estoque_por_organizacao()
            curva_abc = self.data_handler.get_curva_abc_estoque(self.organizacao_selecionada)
            top10 = self.data_handler.get_top10_estoque(self.organizacao_selecionada)
            subinventarios = self.data_handler.get_analise_subinventarios(self.organizacao_selecionada)
            kpis_estoque = self.data_handler.get_kpis_estoque(self.organizacao_selecionada)
            comparativo_org = self.data_handler.get_comparativo_organizacoes()
            
            html = generate_estoque_html(analise_grupos, estoque_stage, estoque_org, curva_abc, 
                                         self.organizacao_selecionada, top10, subinventarios, 
                                         kpis_estoque, comparativo_org)
            self._render_html(html)
        except Exception as e:
            print(f"Erro ao carregar estoque: {e}")
            traceback.print_exc()
            html = f"""
            <div class='header'>
                <h1><i class='fas fa-box'></i> Gestão de Estoque</h1>
            </div>
            <div style='padding: 40px; text-align: center;'>
                <h2>Erro ao carregar dados</h2>
                <p style='color: #f56565;'>{str(e)}</p>
            </div>
            """
            self._render_html(html)
    
    def filtrar_organizacao_estoque(self, organizacao: str):
        """Filtra estoque por organização."""
        self.organizacao_selecionada = organizacao if organizacao else None
        # Recarrega a página atual com o novo filtro
        if self.current_view == "estoque":
            self.load_estoque()
        elif self.current_view == "curva_abc":
            self.load_curva_abc()
        elif self.current_view == "dashboard":
            self.load_dashboard()
    
    def filtrar_organizacao_ordens(self, organizacao: str):
        """Filtra ordens de compra por organização."""
        org = organizacao if organizacao else None
        self.load_ordens(org)
    
    def mostrar_lista_materiais_grupo(self, grupo: str):
        """Mostra lista completa de materiais de um grupo."""
        try:
            materiais = self.data_handler.get_materiais_por_grupo(grupo, self.organizacao_selecionada)
            html = generate_lista_materiais_html(grupo, materiais)
            self._render_html(html)
        except Exception as e:
            print(f"Erro ao mostrar materiais do grupo {grupo}: {e}")
            traceback.print_exc()
    
    def load_ordens(self, organizacao=None):
        """Carrega página de ordens de compra."""
        # Ignorar parâmetro booleano do clicked signal
        if isinstance(organizacao, bool):
            organizacao = None
            
        self.current_view = "ordens"
        self._update_nav("ordens")
        
        try:
            # Obter análise das ordens de compra
            analise = self.data_handler.get_ordens_compra_analise(organizacao)
            
            if not analise:
                html = """
                <div class='header'>
                    <h1><i class='fas fa-shopping-cart'></i> Ordens de Compra</h1>
                </div>
                <div style='padding: 40px; text-align: center;'>
                    <i class='fas fa-inbox' style='font-size: 72px; color: #ccc; margin-bottom: 20px; display: block;'></i>
                    <h2>Nenhum dado de ordens de compra carregado</h2>
                    <p style='color: #666;'>Verifique se o arquivo está no caminho correto.</p>
                </div>
                """
            else:
                html = generate_ordens_compra_html(analise, organizacao)
            
            self._render_html(html)
            
        except Exception as e:
            print(f"Erro ao carregar ordens de compra: {e}")
            traceback.print_exc()
            html = f"""
            <div class='header'>
                <h1><i class='fas fa-shopping-cart'></i> Ordens de Compra</h1>
            </div>
            <div style='padding: 40px; text-align: center;'>
                <i class='fas fa-exclamation-triangle' style='font-size: 72px; color: #f56565; margin-bottom: 20px; display: block;'></i>
                <h2>Erro ao carregar dados</h2>
                <p style='color: #666;'>{str(e)}</p>
            </div>
            """
            self._render_html(html)
    
    def load_movimentacoes(self, organizacao=None):
        """Carrega página de movimentações."""
        # Ignorar parâmetro booleano do clicked signal
        if isinstance(organizacao, bool):
            organizacao = None
            
        self.current_view = "movimentacoes"
        self._update_nav("movimentacoes")
        
        try:
            analise = self.data_handler.get_movimentacoes_analise(organizacao)
            
            if not analise:
                html = """
                <div class='header'>
                    <h1><i class='fas fa-exchange-alt'></i> Movimentações</h1>
                </div>
                <div style='padding: 40px; text-align: center;'>
                    <i class='fas fa-inbox' style='font-size: 72px; color: #ccc; margin-bottom: 20px; display: block;'></i>
                    <h2>Nenhum dado de movimentações carregado</h2>
                    <p style='color: #666;'>Verifique se o arquivo está no caminho correto.</p>
                </div>
                """
            else:
                html = generate_movimentacoes_html(analise, organizacao)
            
            self._render_html(html)
            
        except Exception as e:
            print(f"Erro ao carregar movimentações: {e}")
            traceback.print_exc()
            html = f"""
            <div class='header'>
                <h1><i class='fas fa-exchange-alt'></i> Movimentações</h1>
            </div>
            <div style='padding: 40px; text-align: center;'>
                <i class='fas fa-exclamation-triangle' style='font-size: 72px; color: #f56565; margin-bottom: 20px; display: block;'></i>
                <h2>Erro ao carregar dados</h2>
                <p style='color: #666;'>{str(e)}</p>
            </div>
            """
            self._render_html(html)
    
    def filtrar_organizacao_movimentacoes(self, organizacao: str):
        """Filtra movimentações por organização."""
        org = organizacao if organizacao else None
        self.load_movimentacoes(org)
    
    def load_atividades(self, submenu='compras_servicos'):
        """Carrega página de Atividades com menu horizontal."""
        # Ignorar parâmetro booleano do clicked signal
        if isinstance(submenu, bool):
            submenu = 'compras_servicos'
            
        self.current_view = "atividades"
        self._update_nav("atividades")
        
        # HTML do menu horizontal - igual ao de estoque
        submenu_html = f"""
        <div class='header animate-fade-in'>
            <h1><i class="fas fa-clipboard-list"></i> Atividades por categoria</h1>
            <p class="subtitle">Gestão de processos e operações em tempo real</p>
        </div>
        
        <!-- Barra de Navegação Horizontal -->
        <div class="nav-tabs-container animate-fade-in">
            <div class="nav-tabs">
                <button class="nav-tab {'active' if submenu == 'compras_servicos' else ''}" onclick="loadAtividade('compras_servicos')" data-tab="compras_servicos">
                    <i class="fas fa-file-contract"></i> Contratação de Serviços
                </button>
                <button class="nav-tab {'active' if submenu == 'compras_csc' else ''}" onclick="loadAtividade('compras_csc')" data-tab="compras_csc">
                    <i class="fas fa-building"></i> Compras CSC
                </button>
                <button class="nav-tab {'active' if submenu == 'compras_locais' else ''}" onclick="loadAtividade('compras_locais')" data-tab="compras_locais">
                    <i class="fas fa-store"></i> Compras locais
                </button>
                <button class="nav-tab {'active' if submenu == 'envio_nfe' else ''}" onclick="loadAtividade('envio_nfe')" data-tab="envio_nfe">
                    <i class="fas fa-file-invoice"></i> Envio de NFE
                </button>
                <button class="nav-tab {'active' if submenu == 'reserva_materiais' else ''}" onclick="loadAtividade('reserva_materiais')" data-tab="reserva_materiais">
                    <i class="fas fa-boxes"></i> Reserva de materiais
                </button>
            </div>
        </div>
        
        <style>
            /* Nav Tabs - Atividades */
            .nav-tabs-container {{
                background: white;
                border-radius: 16px;
                padding: 8px;
                margin-bottom: 24px;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
                overflow-x: auto;
            }}
            
            .nav-tabs {{
                display: flex;
                gap: 8px;
                min-width: max-content;
            }}
            
            .nav-tab {{
                display: flex;
                align-items: center;
                gap: 8px;
                padding: 12px 20px;
                border: none;
                background: transparent;
                color: #718096;
                font-size: 0.95rem;
                font-weight: 600;
                border-radius: 10px;
                cursor: pointer;
                transition: all 0.3s ease;
                white-space: nowrap;
            }}
            
            .nav-tab:hover {{
                background: #f7fafc;
                color: #2d3748;
            }}
            
            .nav-tab.active {{
                background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
                color: white;
                box-shadow: 0 4px 15px rgba(45, 55, 72, 0.4);
            }}
            
            .nav-tab i {{
                font-size: 1rem;
            }}
        </style>
        
        <script>
            function loadAtividade(tipo) {{
                if (window.pyBridge) {{
                    window.pyBridge.loadAtividadeSubmenu(tipo);
                }}
            }}
        </script>
        """
        
        # Conteúdo de cada submenu - carrega dados automaticamente
        self.current_pipefy_category = submenu
        pipe_data = self.pipefy_data.get(submenu, {})
        
        if submenu == 'compras_servicos':
            if not pipe_data.get('cards') and not pipe_data.get('loading'):
                self.carregar_dados_pipefy_categoria(submenu)
            content = self.get_pipefy_dashboard_html(submenu)
        elif submenu == 'compras_csc':
            if not pipe_data.get('cards') and not pipe_data.get('loading'):
                self.carregar_dados_pipefy_categoria(submenu)
            content = self.get_pipefy_dashboard_html(submenu)
        elif submenu == 'compras_locais':
            if not pipe_data.get('cards') and not pipe_data.get('loading'):
                self.carregar_dados_pipefy_categoria(submenu)
            content = self.get_pipefy_dashboard_html(submenu)
        elif submenu == 'envio_nfe':
            if not pipe_data.get('cards') and not pipe_data.get('loading'):
                self.carregar_dados_pipefy_categoria(submenu)
            content = self.get_pipefy_dashboard_html(submenu)
        elif submenu == 'reserva_materiais':
            if not pipe_data.get('cards') and not pipe_data.get('loading'):
                self.carregar_dados_pipefy_categoria(submenu)
            content = self.get_pipefy_dashboard_html(submenu)
        else:
            if not pipe_data.get('cards') and not pipe_data.get('loading'):
                self.carregar_dados_pipefy_categoria('compras_servicos')
            content = self.get_pipefy_dashboard_html('compras_servicos')
        
        html = submenu_html + content
        self._render_html(html)
    
    def get_pipefy_dashboard_html(self, categoria='compras_servicos'):
        """Retorna HTML para dashboard Pipefy de qualquer categoria."""
        
        # Obter dados da categoria
        pipe_data = self.pipefy_data.get(categoria, {})
        cards = pipe_data.get('cards', [])
        
        # Compatibilidade retroativa - usar self.pipefy_cards se categoria for compras_servicos e pipefy_data estiver vazio
        if categoria == 'compras_servicos' and not cards and self.pipefy_cards:
            cards = self.pipefy_cards
        
        titulo_categoria = self.pipefy_titulos.get(categoria, categoria)
        
        # Se não tem dados carregados, mostrar tela de carregamento
        if not cards:
            return f"""
            <div style='padding: 16px;'>
                <div style='background: white; border-radius: 12px; padding: 40px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); text-align: center;'>
                    <div class='loading-spinner' style='
                        width: 50px; height: 50px; border: 4px solid #e2e8f0;
                        border-top: 4px solid #2d3748; border-radius: 50%;
                        animation: spin 1s linear infinite; margin: 0 auto 20px;
                    '></div>
                    <h2 style='color: #2d3748; margin-bottom: 12px; font-size: 1.4rem;'>Conectando ao Pipefy...</h2>
                    <p style='color: #666; margin-bottom: 8px;'>Carregando dados de {titulo_categoria}</p>
                    <p style='color: #a0aec0; font-size: 0.85rem;'>Isso pode levar alguns segundos</p>
                </div>
            </div>
            <style>
                @keyframes spin {{
                    0% {{ transform: rotate(0deg); }}
                    100% {{ transform: rotate(360deg); }}
                }}
            </style>
            """
        
        # Processar dados dos cards (já temos a variável cards definida acima)
        total_cards = len(cards)
        from datetime import datetime, timezone
        from collections import defaultdict
        
        # === MÉTRICAS BÁSICAS ===
        concluidos = sum(1 for c in cards if c.get('finished_at'))
        em_andamento = total_cards - concluidos
        taxa_conclusao = (concluidos / total_cards * 100) if total_cards > 0 else 0
        
        # Cards vencidos e próximos do vencimento
        hoje = datetime.now(timezone.utc)
        vencidos = 0
        vencidos_cotacao = 0  # Contador específico para vencidos na fase de cotação
        vence_7_dias = 0
        sem_responsavel = 0
        
        for card in cards:
            if not card.get('finished_at'):
                # Verificar vencimento por tempo na fase configurada (compras_locais, compras_csc, reserva_materiais)
                if verificar_vencimento_fase_cotacao(card, categoria):
                    vencidos += 1
                    vencidos_cotacao += 1  # Contar também como vencido na fase
                    if not card.get('assignees'):
                        sem_responsavel += 1
                    continue  # Já contou como vencido, não verificar due_date
                
                # Para compras_csc, compras_locais e reserva_materiais, apenas considerar vencidos na fase configurada
                # Não usar due_date para essas categorias
                if categoria in ['compras_csc', 'compras_locais', 'reserva_materiais']:
                    if not card.get('assignees'):
                        sem_responsavel += 1
                    continue  # Não verificar due_date para essas categorias
                
                # Para envio_nfe, usar campo "vencimento" dos fields
                if categoria == 'envio_nfe':
                    due = None
                    for field in card.get('fields', []):
                        field_name = (field.get('name') or '').lower()
                        if field_name == 'vencimento':
                            due = field.get('value')
                            break
                else:
                    due = card.get('due_date')
                
                if due:
                    try:
                        # Tentar diferentes formatos de data
                        if 'T' in str(due):
                            due_dt = datetime.fromisoformat(due.replace('Z', '+00:00'))
                        else:
                            # Formato DD/MM/YYYY ou YYYY-MM-DD
                            due_str = str(due)
                            if '/' in due_str:
                                from datetime import datetime as dt
                                due_dt = dt.strptime(due_str, '%d/%m/%Y').replace(tzinfo=timezone.utc)
                            else:
                                due_dt = datetime.fromisoformat(due_str + 'T00:00:00+00:00')
                        
                        if due_dt < hoje:
                            vencidos += 1
                        elif (due_dt - hoje).days <= 7:
                            vence_7_dias += 1
                    except:
                        pass
                if not card.get('assignees'):
                    sem_responsavel += 1
        
        # === ANÁLISE POR FASE ===
        fases_count = {}
        fases_tempo = {}  # Tempo APENAS para cards EM ANDAMENTO
        fases_cards_ativos = {}
        
        for card in cards:
            phase = card.get('current_phase', {})
            phase_name = phase.get('name', 'Sem Fase') if phase else 'Sem Fase'
            fases_count[phase_name] = fases_count.get(phase_name, 0) + 1
            
            # Só calcula tempo para cards NÃO concluídos
            if not card.get('finished_at'):
                fases_cards_ativos[phase_name] = fases_cards_ativos.get(phase_name, 0) + 1
                
                # Calcular tempo na fase atual (desde que entrou até agora)
                for ph in card.get('phases_history', []):
                    ph_name = ph.get('phase', {}).get('name', '')
                    first_in = ph.get('firstTimeIn')
                    if first_in and ph_name:
                        try:
                            start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                            last_out = ph.get('lastTimeOut')
                            # Se não saiu da fase, calcula até agora
                            end = datetime.fromisoformat(last_out.replace('Z', '+00:00')) if last_out else datetime.now(timezone.utc)
                            hours = (end - start).total_seconds() / 3600
                            if ph_name not in fases_tempo:
                                fases_tempo[ph_name] = []
                            fases_tempo[ph_name].append(hours)
                        except:
                            pass
        
        # Tempo médio e total por fase
        fases_analise = []
        for fase, tempos in fases_tempo.items():
            if tempos:
                media_horas = sum(tempos) / len(tempos)
                fases_analise.append({
                    'nome': fase,
                    'media_horas': media_horas,
                    'media_dias': media_horas / 24,
                    'total_passagens': len(tempos),
                    'max_horas': max(tempos),
                    'min_horas': min(tempos),
                    'ativos': fases_cards_ativos.get(fase, 0)
                })
        fases_analise.sort(key=lambda x: x['media_horas'], reverse=True)
        
        # === ANÁLISE POR RESPONSÁVEL ===
        responsaveis_analise = {}
        for card in cards:
            is_finished = bool(card.get('finished_at'))
            assignees = card.get('assignees', [])
            
            # Verificar se está vencido pela fase de Cotação
            vencido_cotacao = verificar_vencimento_fase_cotacao(card, categoria)
            
            if not assignees:
                nome = 'Não Atribuído'
                if nome not in responsaveis_analise:
                    responsaveis_analise[nome] = {'total': 0, 'concluidos': 0, 'pendentes': 0, 'vencidos': 0}
                responsaveis_analise[nome]['total'] += 1
                if is_finished:
                    responsaveis_analise[nome]['concluidos'] += 1
                else:
                    responsaveis_analise[nome]['pendentes'] += 1
                    if vencido_cotacao:
                        responsaveis_analise[nome]['vencidos'] += 1
            else:
                for assignee in assignees:
                    nome = assignee.get('name', 'Desconhecido')
                    if nome not in responsaveis_analise:
                        responsaveis_analise[nome] = {'total': 0, 'concluidos': 0, 'pendentes': 0, 'vencidos': 0}
                    responsaveis_analise[nome]['total'] += 1
                    if is_finished:
                        responsaveis_analise[nome]['concluidos'] += 1
                    else:
                        responsaveis_analise[nome]['pendentes'] += 1
                        # Verificar vencimento por fase de Cotação ou due_date
                        if vencido_cotacao:
                            responsaveis_analise[nome]['vencidos'] += 1
                        else:
                            due = card.get('due_date')
                            if due:
                                try:
                                    due_dt = datetime.fromisoformat(due.replace('Z', '+00:00'))
                                    if due_dt < hoje:
                                        responsaveis_analise[nome]['vencidos'] += 1
                                except:
                                    pass
        
        resp_lista = sorted(responsaveis_analise.items(), key=lambda x: x[1]['total'], reverse=True)
        
        # === ANÁLISE TOP SOLICITANTES DO MÊS (para Compras CSC) ===
        mes_atual = hoje.strftime('%Y-%m')
        solicitantes_mes = {}
        for card in cards:
            created = card.get('created_at', '')
            if created:
                try:
                    dt_criado = datetime.fromisoformat(created.replace('Z', '+00:00'))
                    if dt_criado.strftime('%Y-%m') == mes_atual:
                        # Buscar solicitante nos fields
                        solicitante_nome = ''
                        for field in card.get('fields', []):
                            fn = (field.get('name') or '').lower()
                            fv = field.get('value') or ''
                            if ('solicitante' in fn or 'requisitante' in fn) and isinstance(fv, str) and fv.strip():
                                solicitante_nome = fv.strip()[:40]
                                break
                        if solicitante_nome:
                            solicitantes_mes[solicitante_nome] = solicitantes_mes.get(solicitante_nome, 0) + 1
                except:
                    pass
        
        top_solicitantes = sorted(solicitantes_mes.items(), key=lambda x: x[1], reverse=True)[:5]
        
        # === ANÁLISE FORNECEDOR EXCLUSIVO ===
        fornecedor_exclusivo = 0
        fornecedor_normal = 0
        tipo_contratacao = {}
        urgencias = {'Urgente': 0, 'Normal': 0, 'Baixa': 0}
        
        for card in cards:
            for field in card.get('fields', []):
                fn = field.get('name', '').lower()
                fv = str(field.get('value', '')).lower() if field.get('value') else ''
                
                if 'exclusivo' in fn or ('fornecedor' in fn and 'tipo' in fn):
                    if 'sim' in fv or 'exclusivo' in fv or fv == 'true':
                        fornecedor_exclusivo += 1
                    elif fv:
                        fornecedor_normal += 1
                
                if 'tipo' in fn and 'contrat' in fn:
                    tipo = field.get('value', 'Outros')
                    if tipo:
                        tipo_contratacao[str(tipo)] = tipo_contratacao.get(str(tipo), 0) + 1
                
                if 'urgência' in fn or 'prioridade' in fn or 'urgencia' in fn:
                    if 'alta' in fv or 'urgente' in fv:
                        urgencias['Urgente'] += 1
                    elif 'baixa' in fv:
                        urgencias['Baixa'] += 1
                    elif fv:
                        urgencias['Normal'] += 1
        
        # === ANÁLISE TEMPORAL ===
        cards_por_mes = defaultdict(int)
        concluidos_por_mes = defaultdict(int)
        tempo_ciclo = []  # Lista de tuplas (dias, ano_criacao, ano_conclusao)
        
        for card in cards:
            created = card.get('created_at', '')
            if created:
                try:
                    dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
                    cards_por_mes[dt.strftime('%Y-%m')] += 1
                except:
                    pass
            
            finished = card.get('finished_at', '')
            if finished:
                try:
                    dt_fin = datetime.fromisoformat(finished.replace('Z', '+00:00'))
                    concluidos_por_mes[dt_fin.strftime('%Y-%m')] += 1
                    
                    if created:
                        dt_cri = datetime.fromisoformat(created.replace('Z', '+00:00'))
                        dias = (dt_fin - dt_cri).days
                        # Armazena dias, ano de criação e ano da conclusão
                        tempo_ciclo.append((dias, dt_cri.year, dt_fin.year))
                except:
                    pass
        
        # Média de atendimento do ANO atual (apenas cards CRIADOS e CONCLUÍDOS em 2026)
        ano_atual = hoje.year
        tempo_ciclo_ano = [t[0] for t in tempo_ciclo if t[1] >= ano_atual and t[2] == ano_atual]
        tempo_ciclo_medio = sum(tempo_ciclo_ano) / len(tempo_ciclo_ano) if tempo_ciclo_ano else 0
        
        todos_tempos = [t[0] for t in tempo_ciclo]
        tempo_ciclo_min = min(todos_tempos) if todos_tempos else 0
        tempo_ciclo_max = max(todos_tempos) if todos_tempos else 0
        
        # Filtrar apenas meses de 2026 em diante
        todos_meses = sorted(set(list(cards_por_mes.keys()) + list(concluidos_por_mes.keys())))
        meses_2026 = [m for m in todos_meses if m >= '2026-01']
        meses = meses_2026[-6:] if meses_2026 else []
        criados_mes = [cards_por_mes.get(m, 0) for m in meses]
        finalizados_mes = [concluidos_por_mes.get(m, 0) for m in meses]
        
        # Para envio_nfe, calcular volume total de processos por mês
        if categoria == 'envio_nfe':
            volume_mes = [cards_por_mes.get(m, 0) for m in meses]
        
        # === ANÁLISE TEMPO DE CONCLUSÃO POR MÊS (para Compras CSC e Locais) ===
        tempo_conclusao_por_mes = defaultdict(list)  # {mes: [dias_de_cada_card]}
        tempo_por_solicitante = defaultdict(list)    # {solicitante: [(dias, titulo, url, mes_conclusao)]}
        
        for card in cards:
            created = card.get('created_at', '')
            finished = card.get('finished_at', '')
            
            if created and finished:
                try:
                    dt_cri = datetime.fromisoformat(created.replace('Z', '+00:00'))
                    dt_fin = datetime.fromisoformat(finished.replace('Z', '+00:00'))
                    dias = (dt_fin - dt_cri).days
                    mes_conclusao = dt_fin.strftime('%Y-%m')
                    
                    # Só considerar meses de 2026+
                    if mes_conclusao >= '2026-01':
                        tempo_conclusao_por_mes[mes_conclusao].append(dias)
                        
                        # Buscar solicitante e requisição de compra nos fields
                        solicitante_nome = ''
                        requisicao_compra = ''
                        for field in card.get('fields', []):
                            fn = (field.get('name') or '').lower()
                            fv = field.get('value') or ''
                            if ('solicitante' in fn or 'requisitante' in fn) and isinstance(fv, str) and fv.strip():
                                solicitante_nome = sanitize_for_js(fv.strip()[:40])
                            # Buscar requisição de compra (RC)
                            if ('requisição' in fn or 'requisicao' in fn or fn == 'rc' or 'n° rc' in fn or 'nº rc' in fn or 'numero rc' in fn or 'número rc' in fn or 'req. compra' in fn or 'req compra' in fn) and isinstance(fv, str) and fv.strip():
                                requisicao_compra = sanitize_for_js(fv.strip()[:30])
                        
                        if solicitante_nome:
                            titulo_sanitizado = sanitize_for_js(card.get('title', 'Sem título')[:40])
                            tempo_por_solicitante[solicitante_nome].append({
                                'dias': dias,
                                'titulo': titulo_sanitizado,
                                'url': card.get('url', '#'),
                                'mes': mes_conclusao,
                                'created': dt_cri.strftime('%d/%m/%Y'),
                                'finished': dt_fin.strftime('%d/%m/%Y'),
                                'rc': requisicao_compra
                            })
                except:
                    pass
        
        # Calcular média por mês para o gráfico
        meses_tempo_conclusao = sorted(tempo_conclusao_por_mes.keys())[-6:]  # Últimos 6 meses
        medias_tempo_mes = [round(sum(tempo_conclusao_por_mes[m]) / len(tempo_conclusao_por_mes[m]), 1) 
                           if tempo_conclusao_por_mes[m] else 0 for m in meses_tempo_conclusao]
        qtd_concluidos_mes = [len(tempo_conclusao_por_mes[m]) for m in meses_tempo_conclusao]
        
        # Top solicitantes por tempo médio de conclusão
        solicitantes_tempo = []
        for nome, registros in tempo_por_solicitante.items():
            media_dias = sum(r['dias'] for r in registros) / len(registros) if registros else 0
            solicitantes_tempo.append({
                'nome': nome,
                'media_dias': round(media_dias, 1),
                'total': len(registros),
                'min_dias': min(r['dias'] for r in registros) if registros else 0,
                'max_dias': max(r['dias'] for r in registros) if registros else 0,
                'registros': registros
            })
        solicitantes_tempo.sort(key=lambda x: x['total'], reverse=True)
        
        # === GARGALOS (cards mais antigos em andamento) ===
        gargalos = []
        for card in cards:
            if not card.get('finished_at'):
                created = card.get('created_at', '')
                if created:
                    try:
                        start = datetime.fromisoformat(created.replace('Z', '+00:00'))
                        dias = (hoje - start).days
                        gargalos.append({
                            'titulo': sanitize_for_js(card.get('title', 'Sem título')[:40]),
                            'fase': sanitize_for_js(card.get('current_phase', {}).get('name', 'N/A') if card.get('current_phase') else 'N/A'),
                            'dias': dias,
                            'responsavel': sanitize_for_js(', '.join([a.get('name', '')[:15] for a in card.get('assignees', [])]) or 'Não atribuído'),
                            'url': card.get('url', '#')
                        })
                    except:
                        pass
        gargalos.sort(key=lambda x: x['dias'], reverse=True)
        
        # === ETIQUETAS ===
        etiquetas_count = {}
        for card in cards:
            for label in card.get('labels', []):
                nome = label.get('name', 'Sem nome')
                etiquetas_count[nome] = etiquetas_count.get(nome, 0) + 1
        
        # === ANÁLISE MENSAL POR FASE (para gráficos de meta) ===
        tempo_mensal_por_fase = defaultdict(lambda: defaultdict(list))  # {mes: {fase: [tempos_em_dias]}}
        
        for card in cards:
            for ph in card.get('phases_history', []):
                ph_name = ph.get('phase', {}).get('name', '')
                first_in = ph.get('firstTimeIn')
                last_out = ph.get('lastTimeOut')
                
                if first_in and ph_name:
                    try:
                        start = datetime.fromisoformat(first_in.replace('Z', '+00:00'))
                        # Se não saiu da fase ainda, usar a data atual como fim
                        end = datetime.fromisoformat(last_out.replace('Z', '+00:00')) if last_out else datetime.now(timezone.utc)
                        
                        # Calcular tempo em dias
                        dias = (end - start).total_seconds() / (3600 * 24)
                        
                        # Agrupar por mês (usar mês de entrada na fase)
                        mes = start.strftime('%Y-%m')
                        
                        # Só considerar meses de 2026 em diante
                        if mes >= '2026-01':
                            tempo_mensal_por_fase[mes][ph_name].append(dias)
                    except:
                        pass
        
        # Calcular médias mensais por fase
        metas_data = {}  # {fase: {meses: [...], medias: [...]}}
        todos_meses_metas = sorted(set(tempo_mensal_por_fase.keys()))[-12:]  # Últimos 12 meses
        
        for mes in todos_meses_metas:
            for fase, tempos in tempo_mensal_por_fase[mes].items():
                if fase not in metas_data:
                    metas_data[fase] = {'meses': [], 'medias': []}
                
                # Calcular média do mês para esta fase
                media_mes = sum(tempos) / len(tempos) if tempos else 0
                
                # Adicionar mês se ainda não existe
                if mes not in metas_data[fase]['meses']:
                    metas_data[fase]['meses'].append(mes)
                    metas_data[fase]['medias'].append(round(media_mes, 1))
        
        # Preencher meses faltantes com 0 para cada fase
        for fase in metas_data:
            for mes in todos_meses_metas:
                if mes not in metas_data[fase]['meses']:
                    metas_data[fase]['meses'].append(mes)
                    metas_data[fase]['medias'].append(0)
            
            # Reordenar por mês
            combined = list(zip(metas_data[fase]['meses'], metas_data[fase]['medias']))
            combined.sort()
            metas_data[fase]['meses'] = [m for m, _ in combined]
            metas_data[fase]['medias'] = [v for _, v in combined]
        
        metas_data_json = json.dumps(metas_data)
        
        # === DADOS PARA GRÁFICOS ===
        fases_labels = list(fases_count.keys())[:8]
        fases_values = [fases_count[f] for f in fases_labels]
        
        # Para envio_nfe, mostrar apenas fases específicas
        if categoria == 'envio_nfe':
            fases_nfe = ['pendente de lançamento', 'falha', 'financeiro']
            fases_filtradas = [f for f in fases_analise if f['nome'].lower() in fases_nfe]
            # Ordenar na ordem desejada
            ordem_fases = {nome: i for i, nome in enumerate(fases_nfe)}
            fases_filtradas.sort(key=lambda x: ordem_fases.get(x['nome'].lower(), 99))
            tempo_labels = [f['nome'][:18] for f in fases_filtradas]
            tempo_values = [round(f['media_dias'], 1) for f in fases_filtradas]
        else:
            tempo_labels = [f['nome'][:18] for f in fases_analise[:6]]
            tempo_values = [round(f['media_dias'], 1) for f in fases_analise[:6]]  # Em DIAS
        
        resp_labels = [r[0][:12] for r in resp_lista[:8]]
        resp_concl = [r[1]['concluidos'] for r in resp_lista[:8]]
        resp_pend = [r[1]['pendentes'] for r in resp_lista[:8]]
        
        etiq_labels = list(etiquetas_count.keys())[:6]
        etiq_values = [etiquetas_count[e] for e in etiq_labels]
        
        # === GERAR HTML ===
        # Tabela de gargalos (sem coluna de responsável para caber melhor)
        gargalos_html = ""
        for i, g in enumerate(gargalos[:6], 1):
            cor = '#f56565' if g['dias'] > 30 else '#ed8936' if g['dias'] > 15 else '#48bb78'
            gargalos_html += f"<tr><td>{i}</td><td><a href='{g['url']}' target='_blank' style='color:#2d3748;text-decoration:none;'>{g['titulo'][:25]}</a></td><td>{g['fase'][:12]}</td><td style='text-align:center;'><span style='background:{cor};color:white;padding:2px 6px;border-radius:10px;font-size:0.7rem;'>{g['dias']}d</span></td></tr>"
        
        # Cards de responsáveis
        resp_cards_html = ""
        for nome, dados in resp_lista[:6]:
            taxa = (dados['concluidos'] / dados['total'] * 100) if dados['total'] > 0 else 0
            resp_cards_html += f"""
            <div style='background:#f8fafc;padding:10px 12px;border-radius:8px;margin-bottom:8px;'>
                <div style='display:flex;justify-content:space-between;align-items:center;'>
                    <span style='font-weight:600;color:#2d3748;font-size:0.85rem;'>{nome[:20]}</span>
                    <span style='font-size:0.75rem;color:#666;'>{dados['total']} total</span>
                </div>
                <div style='display:flex;gap:10px;margin-top:6px;font-size:0.75rem;'>
                    <span style='color:#48bb78;'><i class="fas fa-check"></i> {dados['concluidos']} concluidos</span>
                    <span style='color:#ed8936;'><i class="fas fa-clock"></i> {dados['pendentes']} pendentes</span>
                </div>
                <div style='background:#e2e8f0;height:4px;border-radius:2px;margin-top:6px;'>
                    <div style='background:#48bb78;height:100%;width:{taxa}%;border-radius:2px;'></div>
                </div>
            </div>"""
        
        # Cards de top solicitantes do mês (para Compras CSC)
        nome_mes_atual = hoje.strftime('%B/%Y').title()
        solicitantes_cards_html = ""
        total_solicit_mes = sum(s[1] for s in top_solicitantes)
        for i, (nome, qtd) in enumerate(top_solicitantes, 1):
            pct = (qtd / total_solicit_mes * 100) if total_solicit_mes > 0 else 0
            cores = ['#4299e1', '#48bb78', '#ed8936', '#9f7aea', '#f56565']
            cor = cores[(i-1) % len(cores)]
            # Escapar nome para uso em JavaScript
            nome_js = nome.replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"')
            nome_display = nome[:25].replace("'", "&#39;").replace('"', '&quot;')
            solicitantes_cards_html += f"""
            <div style='background:#f8fafc;padding:10px 12px;border-radius:8px;margin-bottom:8px;cursor:pointer;' onclick="showModal('solicitante', '{nome_js}')">
                <div style='display:flex;justify-content:space-between;align-items:center;'>
                    <span style='font-weight:600;color:#2d3748;font-size:0.85rem;'><span style='color:{cor};font-weight:700;'>{i}.</span> {nome_display}</span>
                    <span style='font-size:0.85rem;font-weight:700;color:{cor};'>{qtd}</span>
                </div>
                <div style='background:#e2e8f0;height:4px;border-radius:2px;margin-top:6px;'>
                    <div style='background:{cor};height:100%;width:{pct}%;border-radius:2px;'></div>
                </div>
            </div>"""
        if not solicitantes_cards_html:
            solicitantes_cards_html = "<p style='color:#999;font-size:0.8rem;text-align:center;'>Nenhum solicitante no mês</p>"
        
        # Fases críticas
        fases_criticas_html = ""
        for f in fases_analise[:5]:
            dias = f['media_dias']
            cor = '#f56565' if dias > 5 else '#ed8936' if dias > 2 else '#48bb78'
            fases_criticas_html += f"""
            <div style='display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #eee;'>
                <span style='font-size:0.85rem;color:#2d3748;'>{f['nome'][:22]}</span>
                <div style='text-align:right;'>
                    <span style='background:{cor};color:white;padding:2px 8px;border-radius:10px;font-size:0.75rem;font-weight:600;'>{dias:.1f} dias</span>
                    <span style='font-size:0.7rem;color:#999;margin-left:6px;'>{f['ativos']} ativos</span>
                </div>
            </div>"""
        
        # Preparar dados do modal (fora da f-string para evitar problemas com {{}})
        cards_modal_data = []
        for c in cards:
            fase = c.get('current_phase', {}).get('name', 'N/A') if c.get('current_phase') else 'N/A'
            
            # Buscar campos de solicitação, solicitante e serviço nos fields
            solicitacao = ''
            solicitante = ''
            servico = ''
            tem_etiqueta = len(c.get('labels', [])) > 0
            etiquetas = ', '.join([l.get('name', '') for l in c.get('labels', [])])
            responsavel = ', '.join([a.get('name', '') for a in c.get('assignees', [])])
            
            for field in c.get('fields', []):
                fn = (field.get('name') or '').lower()
                fv = field.get('value') or ''
                
                if 'solicitação' in fn or 'solicitacao' in fn or 'descrição' in fn or 'descricao' in fn:
                    if isinstance(fv, str):
                        solicitacao = sanitize_for_js(fv[:80])
                elif 'solicitante' in fn or 'requisitante' in fn or 'nome' in fn:
                    if isinstance(fv, str) and not solicitante:
                        solicitante = sanitize_for_js(fv[:40])
                elif 'serviço' in fn or 'servico' in fn or 'tipo de serviço' in fn or 'tipo de servico' in fn:
                    if isinstance(fv, str) and not servico:
                        servico = sanitize_for_js(fv[:60])
            
            # Verificar se está vencido pela regra de tempo na fase de Cotação
            vencido_por_cotacao = verificar_vencimento_fase_cotacao(c, categoria)
            
            cards_modal_data.append({
                'titulo': sanitize_for_js(c.get('title', 'Sem título')[:50]),
                'fase': sanitize_for_js(fase),
                'url': c.get('url', '#'),
                'finished': bool(c.get('finished_at')),
                'due': c.get('due_date') or '',
                'created': c.get('created_at') or '',
                'finished_at': c.get('finished_at') or '',
                'solicitacao': solicitacao,
                'solicitante': solicitante,
                'servico': servico,
                'etiquetas': sanitize_for_js(etiquetas),
                'temEtiqueta': tem_etiqueta,
                'responsavel': sanitize_for_js(responsavel),
                'vencidoCotacao': vencido_por_cotacao
            })
        cards_json = json.dumps(cards_modal_data)
        solicitantes_tempo_json = json.dumps(solicitantes_tempo)
        
        return f"""
        <style>
            .dash {{padding:12px;}}
            .toolbar {{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;align-items:center;}}
            .tb-btn {{padding:8px 16px;border:none;border-radius:6px;font-weight:600;font-size:0.8rem;cursor:pointer;display:flex;align-items:center;gap:6px;}}
            .tb-primary {{background:linear-gradient(135deg,#2d3748,#1a202c);color:white;}}
            .tb-success {{background:linear-gradient(135deg,#48bb78,#38a169);color:white;}}
            .kpi-row {{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:16px;}}
            .kpi {{background:white;padding:14px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);text-align:center;cursor:pointer;transition:transform 0.2s,box-shadow 0.2s;}}
            .kpi:hover {{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,0.12);}}
            .kpi-val {{font-size:1.6rem;font-weight:800;color:#2d3748;}}
            .kpi-lbl {{font-size:0.7rem;color:#718096;margin-top:2px;}}
            .grid-2 {{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:12px;}}
            .grid-3 {{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px;}}
            .grid-4 {{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:12px;}}
            .card {{background:white;padding:14px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);overflow:hidden;}}
            .card-title {{font-size:0.85rem;font-weight:700;color:#2d3748;margin-bottom:10px;display:flex;align-items:center;gap:6px;}}
            .chart-container {{position:relative;height:140px !important;max-height:140px !important;width:100%;overflow:hidden;}}
            .chart-container canvas {{max-height:140px !important;}}
            .mini-table {{width:100%;font-size:0.75rem;}}
            .mini-table th {{background:#f7fafc;padding:6px;text-align:left;font-weight:600;color:#4a5568;}}
            .mini-table td {{padding:6px;border-bottom:1px solid #f0f0f0;}}
            .alert-card {{padding:10px 12px;border-radius:8px;margin-bottom:8px;display:flex;align-items:center;gap:10px;}}
            .alert-danger {{background:#fff5f5;border-left:3px solid #f56565;}}
            .alert-warning {{background:#fffaf0;border-left:3px solid #ed8936;}}
            .alert-info {{background:#ebf8ff;border-left:3px solid #4299e1;}}
            .modal-overlay {{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:1000;display:none;align-items:center;justify-content:center;}}
            .modal {{background:white;border-radius:12px;max-width:800px;width:90%;max-height:80vh;overflow:auto;box-shadow:0 20px 60px rgba(0,0,0,0.3);}}
            .modal-large {{max-width:1200px;}}
            .modal-header {{padding:16px 20px;border-bottom:1px solid #eee;display:flex;justify-content:space-between;align-items:center;}}
            .modal-title {{font-size:1.1rem;font-weight:700;color:#2d3748;}}
            .modal-close {{background:none;border:none;font-size:1.5rem;cursor:pointer;color:#666;}}
            .modal-body {{padding:20px;max-height:60vh;overflow-y:auto;}}
            .tb-info {{background:linear-gradient(135deg,#667eea,#764ba2);color:white;}}
            .meta-chart-container {{position:relative;height:300px;margin-bottom:24px;}}
            @media(max-width:1400px){{.kpi-row{{grid-template-columns:repeat(3,1fr);}} .grid-3{{grid-template-columns:repeat(2,1fr);}}}}
        </style>
        
        <div class='dash'>
            <!-- Toolbar -->
            <div class='toolbar'>
                <button class='tb-btn tb-primary' onclick="atualizarDados()"><i class='fas fa-sync-alt'></i> Atualizar</button>
                <button class='tb-btn tb-primary' onclick="exportarPdf()"><i class='fas fa-file-pdf'></i> Exportar PDF</button>
                <button class='tb-btn tb-primary' onclick="showMetasModal()"><i class='fas fa-bullseye'></i>Metas</button>
                <button class='tb-btn tb-primary' id='btnToggleView' onclick="toggleView()"><i class='fas fa-users'></i>Responsável</button>
                <button class='tb-btn' style='background:linear-gradient(135deg,#dc2626,#991b1b);color:white;' onclick="enviarEmail()"><i class='fas fa-envelope'></i> Enviar Email</button>
                <span style='flex:1;'></span>
                <span style='color:#666;font-size:0.8rem;'><i class='fas fa-database'></i> {total_cards} registros | Atualizado: {datetime.now().strftime('%H:%M')}</span>
            </div>
            
            <!-- KPIs Principais -->
            <div class='kpi-row'>
                <div class='kpi' style='border-top:3px solid #4299e1;' onclick="showModal('total')"><div class='kpi-val'>{total_cards}</div><div class='kpi-lbl'>TOTAL</div></div>
                <div class='kpi' style='border-top:3px solid #48bb78;' onclick="showModal('concluidos')"><div class='kpi-val'>{concluidos}</div><div class='kpi-lbl'>CONCLUÍDOS</div></div>
                <div class='kpi' style='border-top:3px solid #ed8936;' onclick="showModal('andamento')"><div class='kpi-val'>{em_andamento}</div><div class='kpi-lbl'>EM ANDAMENTO</div></div>
                <div class='kpi' style='border-top:3px solid #9f7aea;'><div class='kpi-val'>{taxa_conclusao:.0f}%</div><div class='kpi-lbl'>TX CONCLUSÃO</div></div>
                <div class='kpi' style='border-top:3px solid #f56565;' onclick="showModal('vencidos')">
                    <div class='kpi-val'>{vencidos}</div>
                    <div class='kpi-lbl'>VENCIDOS</div>
                    {f"<div style='font-size:0.65rem;color:#f56565;margin-top:2px;font-weight:600;'>{vencidos_cotacao} na Cotação</div>" if categoria in ['compras_csc', 'compras_locais'] and vencidos_cotacao > 0 else ""}
                </div>
                <div class='kpi' style='border-top:3px solid #38b2ac;'><div class='kpi-val'>{tempo_ciclo_medio:.0f}d</div><div class='kpi-lbl'>MÉDIA ATENDIMENTO {ano_atual}</div></div>
            </div>
            
            <!-- Alertas -->
            <div class='grid-3' style='margin-bottom:12px;'>
                <div class='alert-card alert-danger'>
                    <i class='fas fa-exclamation-circle' style='color:#f56565;font-size:1.2rem;'></i>
                    <div>
                        <strong style='color:#c53030;'>{vencidos}</strong> <span style='color:#666;font-size:0.8rem;'>cards vencidos</span>
                        {f"<div style='font-size:0.75rem;color:#f56565;margin-top:2px;'><i class='fas fa-hourglass-end'></i> {vencidos_cotacao} ultrapassaram meta na Cotação</div>" if categoria in ['compras_csc', 'compras_locais'] and vencidos_cotacao > 0 else ""}
                    </div>
                </div>
                <div class='alert-card alert-warning'>
                    <i class='fas fa-clock' style='color:#ed8936;font-size:1.2rem;'></i>
                    <div><strong style='color:#c05621;'>{vence_7_dias}</strong> <span style='color:#666;font-size:0.8rem;'>vencem em 7 dias</span></div>
                </div>
                <div class='alert-card alert-info'>
                    <i class='fas fa-user-slash' style='color:#4299e1;font-size:1.2rem;'></i>
                    <div><strong style='color:#2b6cb0;'>{sem_responsavel}</strong> <span style='color:#666;font-size:0.8rem;'>sem responsável</span></div>
                </div>
            </div>
            
            <!-- Gráficos Linha 1: 2 gráficos -->
            <div class='grid-2'>
                <div class='card'>
                    <div class='card-title'><i class='fas fa-chart-pie' style='color:#4299e1;'></i> Por Fase</div>
                    <div class='chart-container'><canvas id='chartFases'></canvas></div>
                </div>
                <div class='card'>
                    <div class='card-title'><i class='fas fa-hourglass-half' style='color:#f56565;'></i> Tempo por Fase (dias)</div>
                    <div class='chart-container'><canvas id='chartTempo'></canvas></div>
                </div>
            </div>
            
            <!-- Gráficos Linha 2: mais 2 -->
            <div class='grid-2'>
                <div class='card'>
                    <div class='card-title'><i class='fas fa-chart-line' style='color:#9f7aea;'></i> {"Volume Mensal" if categoria == "envio_nfe" else "Evolução Mensal"}</div>
                    <div class='chart-container'><canvas id='chartEvolucao'></canvas></div>
                </div>
                <div class='card'>
                    <div class='card-title'><i class='fas fa-fire' style='color:#f56565;'></i> Fases Críticas (dias)</div>
                    {fases_criticas_html if fases_criticas_html else "<p style='color:#999;font-size:0.8rem;'>Sem dados</p>"}
                </div>
            </div>
            
            {"<!-- Card Tempo por Solicitante (só para Compras CSC e Locais) -->" if categoria in ['compras_csc', 'compras_locais'] else ""}
            {"<div class='grid-1'>" if categoria in ['compras_csc', 'compras_locais'] else ""}
            {"<div class='card'>" if categoria in ['compras_csc', 'compras_locais'] else ""}
            {"<div class='card-title' style='display:flex;justify-content:space-between;align-items:center;'><span><i class='fas fa-users' style='color:#667eea;'></i> Tempo por Solicitante</span><button onclick=\"showModal('tempo_solicitantes')\" style='background:#667eea;color:white;border:none;padding:4px 10px;border-radius:5px;font-size:0.7rem;cursor:pointer;'>Ver Todos</button></div>" if categoria in ['compras_csc', 'compras_locais'] else ""}
            {"<div id='solicitantesTempoContainer'></div>" if categoria in ['compras_csc', 'compras_locais'] else ""}
            {"</div>" if categoria in ['compras_csc', 'compras_locais'] else ""}
            {"</div>" if categoria in ['compras_csc', 'compras_locais'] else ""}
            
            <!-- Linha 3: Performance/Solicitantes e Gargalos -->
            <div class='grid-2'>
                <!-- Performance Responsáveis OU Top Solicitantes (Compras CSC) -->
                <div class='card'>
                    {"<div class='card-title'><i class='fas fa-user-friends' style='color:#4299e1;'></i> Top 5 Solicitantes (" + nome_mes_atual + ")</div>" + solicitantes_cards_html if categoria == 'compras_csc' else "<div class='card-title'><i class='fas fa-user-check' style='color:#48bb78;'></i> Performance</div>" + (resp_cards_html if resp_cards_html else "<p style='color:#999;font-size:0.8rem;'>Sem dados</p>")}
                </div>
                
                <!-- Gargalos -->
                <div class='card'>
                    <div class='card-title'><i class='fas fa-exclamation-triangle' style='color:#ed8936;'></i> Principais Gargalos</div>
                    <table class='mini-table'>
                        <thead><tr><th>#</th><th>Título</th><th>Fase</th><th>Tempo</th></tr></thead>
                        <tbody>{gargalos_html if gargalos_html else "<tr><td colspan='4' style='text-align:center;color:#999;'>Sem gargalos</td></tr>"}</tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- Visualização Por Responsável (inicialmente oculta) -->
        <div class='dash' id='dashResponsaveis' style='display:none;'></div>
        
        <!-- Modal para listas -->
        <div id='modalOverlay' class='modal-overlay' onclick="closeModal(event)">
            <div class='modal' onclick="event.stopPropagation()">
                <div class='modal-header'>
                    <span class='modal-title' id='modalTitle'>Lista</span>
                    <button class='modal-close' onclick="closeModal()">&times;</button>
                </div>
                <div class='modal-body' id='modalBody'></div>
            </div>
        </div>
        
        <script>
            function atualizarDados() {{ if (window.pyBridge) window.pyBridge.atualizarDadosPipefy(); }}
            function exportarPdf() {{ if (window.pyBridge) window.pyBridge.exportarPdfPipefy(); }}
            function enviarEmail() {{ if (window.pyBridge) window.pyBridge.enviarEmailPipefy(); }}
            
            // Estado da visualização (geral ou por responsável)
            let viewMode = 'geral';
            
            // Função para alternar entre visualizações
            function toggleView() {{
                const dashGeral = document.querySelector('.dash:not(#dashResponsaveis)');
                const dashResp = document.getElementById('dashResponsaveis');
                const btn = document.getElementById('btnToggleView');
                
                if (viewMode === 'geral') {{
                    // Mudar para visualização por responsável
                    viewMode = 'responsavel';
                    dashGeral.style.display = 'none';
                    dashResp.style.display = 'block';
                    btn.innerHTML = '<i class="fas fa-chart-bar"></i> Visão Geral';
                    
                    // Construir dashboard de responsáveis
                    buildResponsaveisDashboard();
                }} else {{
                    // Voltar para visualização geral
                    viewMode = 'geral';
                    dashGeral.style.display = 'block';
                    dashResp.style.display = 'none';
                    btn.innerHTML = '<i class="fas fa-users"></i> Responsável';
                }}
            }}
            
            // Construir dashboard por responsável
            function buildResponsaveisDashboard() {{
                const dashResp = document.getElementById('dashResponsaveis');
                
                // Agrupar cards por responsável
                const responsaveisMap = {{}};
                
                cardsData.forEach(card => {{
                    if (!card.responsavel || card.responsavel.trim() === '') {{
                        if (!responsaveisMap['Não Atribuído']) {{
                            responsaveisMap['Não Atribuído'] = [];
                        }}
                        responsaveisMap['Não Atribuído'].push(card);
                    }} else {{
                        // Pode ter múltiplos responsáveis separados por vírgula
                        const nomes = card.responsavel.split(',').map(n => n.trim());
                        nomes.forEach(nome => {{
                            if (!responsaveisMap[nome]) {{
                                responsaveisMap[nome] = [];
                            }}
                            responsaveisMap[nome].push(card);
                        }});
                    }}
                }});
                
                // Converter para array e ordenar por quantidade de cards
                const responsaveisArray = Object.entries(responsaveisMap)
                    .map(([nome, cards]) => ({{ nome, cards }}))
                    .sort((a, b) => b.cards.length - a.cards.length);
                
                // Gerar HTML
                let html = `
                    <div style='margin-bottom:16px;'>
                        <h2 style='color:#2d3748;font-size:1.4rem;margin-bottom:4px;'>
                            <i class='fas fa-users'></i> Análise Individual por Responsável
                        </h2>
                        <p style='color:#718096;font-size:0.9rem;'>Métricas detalhadas de cada responsável</p>
                    </div>
                    
                    <!-- Grid de responsáveis -->
                    <div style='display:grid;grid-template-columns:repeat(auto-fill,minmax(350px,1fr));gap:16px;'>
                `;
                
                responsaveisArray.forEach(resp => {{
                    const nome = resp.nome;
                    const cards = resp.cards;
                    const total = cards.length;
                    const concluidos = cards.filter(c => c.finished).length;
                    const pendentes = total - concluidos;
                    const taxa = total > 0 ? ((concluidos / total) * 100).toFixed(0) : 0;
                    
                    // Contar vencidos
                    let vencidos = 0;
                    cards.forEach(c => {{
                        if (!c.finished) {{
                            if (c.vencidoCotacao) {{
                                vencidos++;
                            }} else if (c.due) {{
                                const dueDate = new Date(c.due);
                                if (dueDate < hoje) {{
                                    vencidos++;
                                }}
                            }}
                        }}
                    }});
                    
                    // Distribuição por fase
                    const fases = {{}};
                    cards.forEach(c => {{
                        if (!c.finished) {{
                            fases[c.fase] = (fases[c.fase] || 0) + 1;
                        }}
                    }});
                    const fasesArray = Object.entries(fases).sort((a, b) => b[1] - a[1]).slice(0, 3);
                    
                    // Cor do card baseado na performance
                    let corBorda = '#48bb78'; // Verde
                    if (vencidos > 0) corBorda = '#f56565'; // Vermelho
                    else if (pendentes > concluidos) corBorda = '#ed8936'; // Laranja
                    
                    html += `
                        <div class='card' style='border-top:3px solid ${{corBorda}};cursor:pointer;transition:transform 0.2s,box-shadow 0.2s;'
                             onmouseover='this.style.transform="translateY(-4px)";this.style.boxShadow="0 8px 20px rgba(0,0,0,0.15)";'
                             onmouseout='this.style.transform="translateY(0)";this.style.boxShadow="0 2px 8px rgba(0,0,0,0.06)";'
                             onclick='showModal("responsavel", "${{nome.replace(/'/g, "\\\\'")}}");'>
                            
                            <!-- Cabeçalho com nome -->
                            <div style='display:flex;align-items:center;gap:10px;margin-bottom:12px;'>
                                <div style='width:40px;height:40px;border-radius:50%;background:linear-gradient(135deg,${{corBorda}},${{corBorda}}dd);display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:1.1rem;'>
                                    ${{nome.charAt(0).toUpperCase()}}
                                </div>
                                <div style='flex:1;'>
                                    <div style='font-weight:700;color:#2d3748;font-size:0.95rem;'>${{nome.length > 25 ? nome.substring(0, 25) + '...' : nome}}</div>
                                    <div style='font-size:0.75rem;color:#718096;'>${{total}} cards no total</div>
                                </div>
                            </div>
                            
                            <!-- KPIs mini -->
                            <div style='display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px;'>
                                <div style='background:#f7fafc;padding:8px;border-radius:6px;text-align:center;'>
                                    <div style='font-size:1.3rem;font-weight:800;color:#48bb78;'>${{concluidos}}</div>
                                    <div style='font-size:0.65rem;color:#718096;'>Concluídos</div>
                                </div>
                                <div style='background:#f7fafc;padding:8px;border-radius:6px;text-align:center;'>
                                    <div style='font-size:1.3rem;font-weight:800;color:#ed8936;'>${{pendentes}}</div>
                                    <div style='font-size:0.65rem;color:#718096;'>Pendentes</div>
                                </div>
                                <div style='background:#f7fafc;padding:8px;border-radius:6px;text-align:center;'>
                                    <div style='font-size:1.3rem;font-weight:800;color:#f56565;'>${{vencidos}}</div>
                                    <div style='font-size:0.65rem;color:#718096;'>Vencidos</div>
                                </div>
                            </div>
                            
                            <!-- Taxa de conclusão -->
                            <div style='margin-bottom:10px;'>
                                <div style='display:flex;justify-content:space-between;margin-bottom:4px;'>
                                    <span style='font-size:0.75rem;color:#718096;font-weight:600;'>Taxa de Conclusão</span>
                                    <span style='font-size:0.75rem;color:#2d3748;font-weight:700;'>${{taxa}}%</span>
                                </div>
                                <div style='background:#e2e8f0;height:6px;border-radius:3px;overflow:hidden;'>
                                    <div style='background:${{corBorda}};height:100%;width:${{taxa}}%;transition:width 0.3s;'></div>
                                </div>
                            </div>
                            
                            <!-- Top 3 fases com cards -->
                            ${{fasesArray.length > 0 ? `
                                <div style='border-top:1px solid #e2e8f0;padding-top:10px;'>
                                    <div style='font-size:0.7rem;color:#718096;font-weight:600;margin-bottom:6px;'>CARDS POR FASE (EM ANDAMENTO)</div>
                                    ${{fasesArray.map(([fase, qtd]) => `
                                        <div style='display:flex;justify-content:space-between;align-items:center;padding:4px 0;font-size:0.75rem;'>
                                            <span style='color:#4a5568;'>${{fase.length > 20 ? fase.substring(0, 20) + '...' : fase}}</span>
                                            <span style='background:#e2e8f0;color:#2d3748;padding:2px 8px;border-radius:10px;font-weight:600;'>${{qtd}}</span>
                                        </div>
                                    `).join('')}}
                                </div>
                            ` : '<p style="color:#a0aec0;font-size:0.75rem;text-align:center;margin-top:8px;">Nenhum card em andamento</p>'}}
                        </div>
                    `;
                }});
                
                html += '</div>';
                dashResp.innerHTML = html;
            }}
            
            // Dados para modal
            const cardsData = {cards_json};
            const hoje = new Date();
            const fasesLabels = {json.dumps(fases_labels)};

            
            function showModal(tipo, filtroValor) {{
                const modal = document.getElementById('modalOverlay');
                const title = document.getElementById('modalTitle');
                const body = document.getElementById('modalBody');
                let items = [];
                
                // Tratamento especial para modais de tempo de conclusão
                if (tipo === 'tempo_solicitantes') {{
                    title.textContent = 'Tempo de Conclusão por Solicitante';
                    let html = '<table class="mini-table"><thead><tr><th>#</th><th>Solicitante</th><th>Média (dias)</th><th>Qtd</th><th>Mín</th><th>Máx</th><th>Ação</th></tr></thead><tbody>';
                    
                    if (typeof solicitantesTempoData !== 'undefined') {{
                        solicitantesTempoData.forEach((s, i) => {{
                            const cor = s.media_dias > 10 ? '#f56565' : s.media_dias > 5 ? '#ed8936' : '#48bb78';
                            const nome_js = s.nome.replace(/\\\\/g, '\\\\\\\\').replace(/'/g, "\\\\'");
                            html += `<tr>
                                <td>${{i+1}}</td>
                                <td style="font-weight:600;">${{s.nome.substring(0,30)}}</td>
                                <td style="text-align:center;"><span style="background:${{cor}};color:white;padding:2px 8px;border-radius:10px;font-size:0.75rem;">${{s.media_dias}}d</span></td>
                                <td style="text-align:center;">${{s.total}}</td>
                                <td style="text-align:center;">${{s.min_dias}}d</td>
                                <td style="text-align:center;">${{s.max_dias}}d</td>
                                <td><button onclick="showModal('detalhe_solicitante', '${{nome_js}}')" style="background:#667eea;color:white;border:none;padding:4px 8px;border-radius:4px;font-size:0.7rem;cursor:pointer;">Detalhes</button></td>
                            </tr>`;
                        }});
                    }}
                    html += '</tbody></table>';
                    body.innerHTML = html;
                    modal.style.display = 'flex';
                    return;
                }}
                
                if (tipo === 'detalhe_solicitante') {{
                    title.textContent = 'Detalhes: ' + filtroValor;
                    let solicitanteData = null;
                    if (typeof solicitantesTempoData !== 'undefined') {{
                        solicitanteData = solicitantesTempoData.find(s => s.nome === filtroValor);
                    }}
                    
                    if (solicitanteData && solicitanteData.registros) {{
                        let html = `<div style="margin-bottom:16px;background:#f7fafc;padding:12px;border-radius:8px;">
                            <div style="display:flex;gap:20px;flex-wrap:wrap;">
                                <div><strong>Média:</strong> <span style="color:#667eea;font-weight:700;">${{solicitanteData.media_dias}} dias</span></div>
                                <div><strong>Total Concluídos:</strong> ${{solicitanteData.total}}</div>
                                <div><strong>Mínimo:</strong> ${{solicitanteData.min_dias}} dias</div>
                                <div><strong>Máximo:</strong> ${{solicitanteData.max_dias}} dias</div>
                            </div>
                        </div>`;
                        html += '<table class="mini-table"><thead><tr><th>#</th><th>Título</th><th>RC</th><th>Solicitado</th><th>Concluído</th><th>Tempo</th></tr></thead><tbody>';
                        solicitanteData.registros.forEach((r, i) => {{
                            const cor = r.dias > 10 ? '#f56565' : r.dias > 5 ? '#ed8936' : '#48bb78';
                            html += `<tr>
                                <td>${{i+1}}</td>
                                <td><a href="${{r.url}}" target="_blank" style="color:#2d3748;text-decoration:none;">${{r.titulo}}</a></td>
                                <td style="font-size:0.75rem;color:#667eea;font-weight:600;">${{r.rc || '-'}}</td>
                                <td style="font-size:0.75rem;">${{r.created}}</td>
                                <td style="font-size:0.75rem;">${{r.finished}}</td>
                                <td style="text-align:center;"><span style="background:${{cor}};color:white;padding:2px 8px;border-radius:10px;font-size:0.75rem;">${{r.dias}}d</span></td>
                            </tr>`;
                        }});
                        html += '</tbody></table>';
                        body.innerHTML = html;
                    }} else {{
                        body.innerHTML = '<p style="text-align:center;color:#999;padding:20px;">Nenhum registro encontrado.</p>';
                    }}
                    modal.style.display = 'flex';
                    return;
                }}
                
                if (tipo === 'total') {{
                    title.textContent = 'Todos os Cards ({total_cards})';
                    items = cardsData;
                }} else if (tipo === 'concluidos') {{
                    title.textContent = 'Cards Concluídos ({concluidos})';
                    items = cardsData.filter(c => c.finished);
                }} else if (tipo === 'andamento') {{
                    title.textContent = 'Cards em Andamento ({em_andamento})';
                    items = cardsData.filter(c => !c.finished);
                }} else if (tipo === 'vencidos') {{
                    title.textContent = 'Cards Vencidos ({vencidos})';
                    items = cardsData.filter(c => {{
                        if (c.finished) return false;
                        // Para compras_csc, compras_locais e reserva_materiais, apenas considerar vencidos na fase configurada
                        // compras_csc: Cotação > 11 dias, compras_locais: Cotação > 2 dias, reserva_materiais: Pendente > 2 dias
                        if (categoria_atual === 'compras_csc' || categoria_atual === 'compras_locais' || categoria_atual === 'reserva_materiais') {{
                            return c.vencidoCotacao === true;
                        }}
                        // Para outras categorias, verificar vencimento por tempo na fase
                        if (c.vencidoCotacao) return true;
                        // Verificar vencimento por due_date
                        if (c.due && new Date(c.due) < hoje) return true;
                        return false;
                    }});
                }} else if (tipo === 'fase') {{
                    title.textContent = 'Cards na Fase: ' + filtroValor;
                    items = cardsData.filter(c => c.fase === filtroValor);
                }} else if (tipo === 'etiqueta') {{
                    title.textContent = 'Cards com Etiqueta: ' + filtroValor;
                    items = cardsData.filter(c => c.etiquetas && c.etiquetas.includes(filtroValor));
                }} else if (tipo === 'sem_etiqueta') {{
                    title.textContent = 'Cards SEM Etiqueta';
                    items = cardsData.filter(c => !c.temEtiqueta);
                }} else if (tipo === 'responsavel') {{
                    title.textContent = 'Cards do Responsável: ' + filtroValor;
                    items = cardsData.filter(c => c.responsavel && c.responsavel.includes(filtroValor));
                }} else if (tipo === 'solicitante') {{
                    title.textContent = 'Cards do Solicitante: ' + filtroValor;
                    items = cardsData.filter(c => c.solicitante && c.solicitante.includes(filtroValor));
                }}
                
                // Salvar itens filtrados globalmente para a função de filtro
                window.currentModalItems = items;
                
                let html = `<div style="margin-bottom:12px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;">
                    <div style="flex:1;position:relative;">
                        <i class="fas fa-search" style="position:absolute;left:10px;top:50%;transform:translateY(-50%);color:#a0aec0;"></i>
                        <input type="text" id="modalFilterInput" placeholder="Filtrar por título, serviço, fase ou responsável..." 
                            style="width:100%;padding:10px 10px 10px 35px;border:1px solid #e2e8f0;border-radius:8px;font-size:0.85rem;outline:none;"
                            onkeyup="filterModalTable()">
                    </div>
                    <select id="modalFilterFase" onchange="filterModalTable()" style="padding:10px;border:1px solid #e2e8f0;border-radius:8px;font-size:0.85rem;outline:none;min-width:120px;">
                        <option value="">Todas as Fases</option>
                    </select>
                    <span id="modalItemCount" style="font-size:0.8rem;color:#718096;"></span>
                </div>`;
                
                html += '<div id="modalTableContainer">';
                html += buildModalTable(items);
                html += '</div>';
                
                body.innerHTML = html;
                
                // Preencher dropdown de fases
                const fases = [...new Set(items.map(i => i.fase))].sort();
                const selectFase = document.getElementById('modalFilterFase');
                fases.forEach(f => {{
                    const opt = document.createElement('option');
                    opt.value = f;
                    opt.textContent = f;
                    selectFase.appendChild(opt);
                }});
                
                document.getElementById('modalItemCount').textContent = items.length + ' registros';
                modal.style.display = 'flex';
            }}
            
            function buildModalTable(items) {{
                if (items.length === 0) {{
                    return '<p style="text-align:center;color:#999;padding:20px;">Nenhum item encontrado.</p>';
                }}
                let html = '<table class="mini-table"><thead><tr><th>#</th><th>Título</th><th>Serviço</th><th>Solicitante</th><th>Fase</th><th>Responsável</th></tr></thead><tbody>';
                items.forEach((item, i) => {{
                    html += `<tr>
                        <td>${{i+1}}</td>
                        <td><a href="${{item.url}}" target="_blank" style="color:#2d3748;text-decoration:none;">${{item.titulo}}</a></td>
                        <td style="font-size:0.75rem;color:#4a5568;">${{item.servico || '-'}}</td>
                        <td style="font-size:0.75rem;color:#666;">${{item.solicitante || '-'}}</td>
                        <td>${{item.fase}}</td>
                        <td style="font-size:0.75rem;">${{item.responsavel || '-'}}</td>
                    </tr>`;
                }});
                html += '</tbody></table>';
                return html;
            }}
            
            function filterModalTable() {{
                const searchText = document.getElementById('modalFilterInput').value.toLowerCase();
                const faseFilter = document.getElementById('modalFilterFase').value;
                
                let filtered = window.currentModalItems.filter(item => {{
                    // Filtro por texto
                    const matchText = !searchText || 
                        (item.titulo && item.titulo.toLowerCase().includes(searchText)) ||
                        (item.servico && item.servico.toLowerCase().includes(searchText)) ||
                        (item.solicitante && item.solicitante.toLowerCase().includes(searchText)) ||
                        (item.fase && item.fase.toLowerCase().includes(searchText)) ||
                        (item.responsavel && item.responsavel.toLowerCase().includes(searchText));
                    
                    // Filtro por fase
                    const matchFase = !faseFilter || item.fase === faseFilter;
                    
                    return matchText && matchFase;
                }});
                
                document.getElementById('modalTableContainer').innerHTML = buildModalTable(filtered);
                document.getElementById('modalItemCount').textContent = filtered.length + ' de ' + window.currentModalItems.length + ' registros';
            }}
            
            function closeModal(event) {{
                if (!event || event.target.id === 'modalOverlay') {{
                    const modal = document.getElementById('modalOverlay');
                    const modalDiv = modal.querySelector('.modal');
                    modal.style.display = 'none';
                    // Remover classe modal-large se existir
                    if (modalDiv) {{
                        modalDiv.classList.remove('modal-large');
                    }}
                }}
            }}
            
            // Gráfico Fases (Doughnut compacto)
            const chartFases = new Chart(document.getElementById('chartFases'), {{
                type: 'doughnut',
                data: {{
                    labels: {json.dumps(fases_labels)},
                    datasets: [{{ data: {json.dumps(fases_values)}, backgroundColor: ['#4299e1','#48bb78','#ed8936','#f56565','#9f7aea','#38b2ac','#667eea','#fc8181'], borderWidth: 0 }}]
                }},
                options: {{ 
                    responsive: true, maintainAspectRatio: false, animation: false, 
                    plugins: {{ legend: {{ display: false }} }}, cutout: '60%',
                    onClick: (e, elements) => {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            showModal('fase', fasesLabels[idx]);
                        }}
                    }}
                }}
            }});
            
            // Gráfico Tempo (Bar horizontal)
            const chartTempo = new Chart(document.getElementById('chartTempo'), {{
                type: 'bar',
                data: {{
                    labels: {json.dumps(tempo_labels)},
                    datasets: [{{ data: {json.dumps(tempo_values)}, backgroundColor: 'rgba(245,101,101,0.8)', borderRadius: 4 }}]
                }},
                options: {{ 
                    indexAxis: 'y', responsive: true, maintainAspectRatio: false, animation: false, 
                    plugins: {{ legend: {{ display: false }} }}, 
                    scales: {{ x: {{ display: false }}, y: {{ ticks: {{ font: {{ size: 9 }} }} }} }},
                    onClick: (e, elements) => {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            const faseNome = {json.dumps(tempo_labels)}[idx];
                            showModal('fase', faseNome);
                        }}
                    }}
                }}
            }});
            

            
            // Gráfico Evolução (Line)
            {"" if categoria == 'envio_nfe' else ""}
            new Chart(document.getElementById('chartEvolucao'), {{
                type: '{"bar" if categoria == "envio_nfe" else "line"}',
                data: {{
                    labels: {json.dumps(meses)},
                    datasets: [
                        {"{ label: 'Volume', data: " + json.dumps(volume_mes if categoria == 'envio_nfe' else criados_mes) + ", backgroundColor: 'rgba(66,153,225,0.8)', borderColor: '#4299e1', borderRadius: 4 }" if categoria == 'envio_nfe' else "{ label: 'Criados', data: " + json.dumps(criados_mes) + ", borderColor: '#4299e1', backgroundColor: 'rgba(66,153,225,0.1)', fill: true, tension: 0.4 }, { label: 'Concluídos', data: " + json.dumps(finalizados_mes) + ", borderColor: '#48bb78', backgroundColor: 'rgba(72,187,120,0.1)', fill: true, tension: 0.4 }"}
                    ]
                }},
                options: {{ responsive: true, maintainAspectRatio: false, animation: false, plugins: {{ legend: {{ display: false }} }}, scales: {{ x: {{ ticks: {{ font: {{ size: 8 }} }} }}, y: {{ display: false }} }} }}
            }});
            
            // Dados de tempo por solicitante (só para Compras CSC e Locais)
            {"" if categoria not in ['compras_csc', 'compras_locais'] else f'''
            const solicitantesTempoData = {solicitantes_tempo_json};
            
            // Popular card de tempo por solicitante
            const container = document.getElementById('solicitantesTempoContainer');
            if (container && solicitantesTempoData.length > 0) {{
                let html = '';
                const top5 = solicitantesTempoData.slice(0, 5);
                top5.forEach((s, i) => {{
                    const cor = s.media_dias > 10 ? '#f56565' : s.media_dias > 5 ? '#ed8936' : '#48bb78';
                    const nome_js = s.nome.replace(/\\\\/g, '\\\\\\\\\\\\\\\\').replace(/'/g, "\\\\'");
                    html += `
                    <div style='background:#f8fafc;padding:10px 12px;border-radius:8px;margin-bottom:8px;cursor:pointer;' onclick="showModal('detalhe_solicitante', '${{nome_js}}')">
                        <div style='display:flex;justify-content:space-between;align-items:center;'>
                            <span style='font-weight:600;color:#2d3748;font-size:0.85rem;'>${{s.nome.substring(0,25)}}</span>
                            <span style='font-size:0.85rem;font-weight:700;color:${{cor}};'>${{s.media_dias}}d</span>
                        </div>
                        <div style='display:flex;gap:10px;margin-top:4px;font-size:0.7rem;color:#718096;'>
                            <span>${{s.total}} concluídos</span>
                            <span>Min: ${{s.min_dias}}d</span>
                            <span>Max: ${{s.max_dias}}d</span>
                        </div>
                    </div>`;
                }});
                container.innerHTML = html;
            }} else if (container) {{
                container.innerHTML = '<p style="color:#999;font-size:0.8rem;text-align:center;">Nenhum dado de conclusão</p>';
            }}
            '''}
            
            // === FUNÇÃO PARA MODAL DE METAS ===
            const categoria_atual = '{categoria}';
            const metasData = {metas_data_json};
            
            // Configuração de metas por categoria e fase (em dias)
            const metasConfig = {{
                'compras_servicos': {{
                    'Cotação': 5,
                    'Emissão ordem de compra': 1
                }},
                'compras_csc': {{
                    'Follow up': 1,
                    'Cotação': 11
                }},
                'compras_locais': {{
                    'Cotação': 2,
                    'Follow up': 1
                }},
                'envio_nfe': {{
                    'Pendente de lançamento': 3,
                    'Falha': 6
                }},
                'reserva_materiais': {{
                    'Pendente': 2,
                    'Entregue': 2
                }}
            }};
            
            function showMetasModal() {{
                const modal = document.getElementById('modalOverlay');
                const modalDiv = modal.querySelector('.modal');
                const title = document.getElementById('modalTitle');
                const body = document.getElementById('modalBody');
                
                // Adicionar classe modal-large
                modalDiv.classList.add('modal-large');
                
                title.textContent = 'Metas de Tempo por Fase';
                
                // Obter configurações de meta para a categoria atual
                const metas = metasConfig[categoria_atual] || {{}};
                
                if (Object.keys(metas).length === 0) {{
                    body.innerHTML = '<p style="text-align:center;color:#999;padding:40px;">Nenhuma meta definida para esta categoria.</p>';
                    modal.style.display = 'flex';
                    return;
                }}
                
                // Obter anos disponíveis nos dados
                let todosAnos = new Set();
                for (let fase in metasData) {{
                    if (metasData[fase].meses) {{
                        metasData[fase].meses.forEach(mes => {{
                            const ano = mes.split('-')[0];
                            todosAnos.add(ano);
                        }});
                    }}
                }}
                const anosDisponiveis = Array.from(todosAnos).sort().reverse();
                const anoAtual = anosDisponiveis[0] || '2026';
                
                // Adicionar seletor de ano
                let html = `
                <div style="margin-bottom:20px;display:flex;align-items:center;gap:12px;">
                    <label style="font-weight:600;color:#2d3748;">Ano:</label>
                    <select id="anoSelector" style="padding:8px 12px;border:1px solid #e2e8f0;border-radius:6px;font-size:0.9rem;outline:none;cursor:pointer;" onchange="atualizarGraficosMetas()">
                        ${{anosDisponiveis.map(ano => `<option value="${{ano}}" ${{ano === anoAtual ? 'selected' : ''}}>${{ano}}</option>`).join('')}}
                    </select>
                </div>`;
                
                // Gerar gráficos para cada meta
                html += '<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:20px;">';
                
                Object.keys(metas).forEach((faseName, index) => {{
                    const meta = metas[faseName];
                    const chartId = 'metaChart' + index;
                    
                    html += `
                    <div class="card">
                        <div class="card-title"><i class="fas fa-chart-bar" style="color:#667eea;"></i> ${{faseName}}</div>
                        <div style="background:#f7fafc;padding:8px 12px;border-radius:6px;margin-bottom:12px;display:flex;justify-content:space-between;align-items:center;">
                            <span style="font-size:0.85rem;color:#666;">Meta:</span>
                            <span style="font-weight:700;color:#48bb78;font-size:1rem;">${{meta}} dias</span>
                        </div>
                        <div class="meta-chart-container">
                            <canvas id="${{chartId}}"></canvas>
                        </div>
                    </div>`;
                }});
                
                html += '</div>';
                body.innerHTML = html;
                modal.style.display = 'flex';
                html += '</div>';
                body.innerHTML = html;
                modal.style.display = 'flex';
                
                // Armazenar referências dos gráficos globalmente
                window.metasCharts = [];
                
                // Criar gráficos iniciais
                atualizarGraficosMetas();
            }}
            
            function atualizarGraficosMetas() {{
                const anoSelecionado = document.getElementById('anoSelector')?.value || '2026';
                const metas = metasConfig[categoria_atual] || {{}};
                
                // Destruir gráficos anteriores se existirem
                if (window.metasCharts) {{
                    window.metasCharts.forEach(chart => {{
                        if (chart) chart.destroy();
                    }});
                    window.metasCharts = [];
                }}
                
                // Todos os meses do ano
                const todosMeses = [
                    'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
                    'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'
                ];
                
                Object.keys(metas).forEach((faseName, index) => {{
                    const meta = metas[faseName];
                    
                    // Buscar dados da fase (case-insensitive)
                    let faseData = null;
                    for (let key in metasData) {{
                        if (key.toLowerCase() === faseName.toLowerCase()) {{
                            faseData = metasData[key];
                            break;
                        }}
                    }}
                    
                    const chartId = 'metaChart' + index;
                    const canvas = document.getElementById(chartId);
                    
                    if (!canvas) return;
                    
                    // Preparar dados para o ano selecionado
                    let valores = new Array(12).fill(0);
                    
                    if (faseData && faseData.meses && faseData.medias) {{
                        faseData.meses.forEach((mes, i) => {{
                            const [ano, mesNum] = mes.split('-');
                            if (ano === anoSelecionado) {{
                                const mesIndex = parseInt(mesNum) - 1;
                                if (mesIndex >= 0 && mesIndex < 12) {{
                                    valores[mesIndex] = faseData.medias[i];
                                }}
                            }}
                        }});
                    }}
                    
                    // Criar array de meta para todos os meses
                    const metaArray = new Array(12).fill(meta);
                    
                    const chart = new Chart(canvas, {{
                        type: 'bar',
                        data: {{
                            labels: todosMeses,
                            datasets: [
                                {{
                                    label: 'Média (dias)',
                                    data: valores,
                                    backgroundColor: valores.map(v => v > meta ? 'rgba(245,101,101,0.8)' : 'rgba(72,187,120,0.8)'),
                                    borderColor: valores.map(v => v > meta ? '#f56565' : '#48bb78'),
                                    borderWidth: 2,
                                    borderRadius: 4,
                                    barThickness: 30
                                }},
                                {{
                                    label: 'Meta',
                                    data: metaArray,
                                    type: 'line',
                                    borderColor: '#48bb78',
                                    backgroundColor: '#48bb78',
                                    borderWidth: 3,
                                    pointRadius: 0,
                                    pointHoverRadius: 0,
                                    fill: false,
                                    tension: 0,
                                    borderDash: [],
                                    order: 0
                                }}
                            ]
                        }},
                        options: {{
                            responsive: true,
                            maintainAspectRatio: false,
                            interaction: {{
                                mode: 'index',
                                intersect: false
                            }},
                            plugins: {{
                                legend: {{
                                    display: true,
                                    position: 'top',
                                    labels: {{
                                        font: {{ size: 11 }},
                                        padding: 8,
                                        usePointStyle: true
                                    }}
                                }},
                                tooltip: {{
                                    callbacks: {{
                                        label: function(context) {{
                                            if (context.dataset.label === 'Meta') {{
                                                return 'Meta: ' + context.parsed.y + ' dias';
                                            }}
                                            const valor = context.parsed.y;
                                            if (valor === 0) {{
                                                return 'Sem dados';
                                            }}
                                            const diff = valor - meta;
                                            const status = diff > 0 ? 
                                                '+' + diff.toFixed(1) + ' dias acima da meta' : 
                                                Math.abs(diff).toFixed(1) + ' dias abaixo da meta';
                                            return 'Média: ' + valor.toFixed(1) + ' dias (' + status + ')';
                                        }}
                                    }}
                                }}
                            }},
                            scales: {{
                                x: {{
                                    ticks: {{
                                        font: {{ size: 10 }}
                                    }},
                                    grid: {{
                                        display: false
                                    }}
                                }},
                                y: {{
                                    beginAtZero: true,
                                    ticks: {{
                                        font: {{ size: 10 }},
                                        callback: function(value) {{
                                            return value + 'd';
                                        }}
                                    }},
                                    grid: {{
                                        color: '#f0f0f0'
                                    }}
                                }}
                            }}
                        }}
                    }});
                    
                    window.metasCharts.push(chart);
                }});
            }}

        </script>
        """
    
    def load_curva_abc(self):
        """Carrega página da Curva ABC."""
        self.current_view = "curva_abc"
        self._update_nav("curva_abc")
        
        if self.data_handler.df_estoque is None or len(self.data_handler.df_estoque) == 0:
            html = """
            <div class='header'>
                <h1><i class='fas fa-chart-bar'></i> Curva ABC</h1>
            </div>
            <div style='padding: 40px; text-align: center;'>
                <i class='fas fa-inbox' style='font-size: 72px; color: #ccc; margin-bottom: 20px;'></i>
                <h2>Nenhum dado de estoque carregado</h2>
                <p style='color: #666; margin: 20px 0;'>Importe o arquivo de estoque periódico para visualizar a Curva ABC.</p>
            </div>
            """
            self._render_html(html)
            return
        
        try:
            curva_abc = self.data_handler.get_curva_abc_estoque(self.organizacao_selecionada)
            estoque_org = self.data_handler.get_estoque_por_organizacao()
            html = generate_curva_abc_page_html(curva_abc, estoque_org, self.organizacao_selecionada)
            self._render_html(html)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar Curva ABC:\n{str(e)}")
    
    def load_pagamentos(self):
        """Carrega página de relatório de pagamentos."""
        self.current_view = "pagamentos"
        self._update_nav("pagamentos")
        
        if self.data_handler.df_pagamentos is None or len(self.data_handler.df_pagamentos) == 0:
            html = """
            <div class='header'>
                <h1><i class="fas fa-money-bill-wave"></i> Relatório de Pagamentos</h1>
            </div>
            <div style='padding: 40px; text-align: center;'>
                <i class='fas fa-inbox' style='font-size: 72px; color: #ccc; margin-bottom: 20px;'></i>
                <h2>Nenhum dado de pagamento carregado</h2>
                <p style='color: #666; margin: 20px 0;'>Importe o arquivo de pagamentos para visualizar as análises.</p>
            </div>
            """
            self._render_html(html)
            return
        
        try:
            # Criar analisador de pagamentos
            analyzer = PaymentAnalyzer(self.data_handler.df_pagamentos)
            
            # Coletar todos os dados necessários
            payment_data = {
                'top_fornecedores': analyzer.get_fornecedores_mais_faturados(10),
                'formas_pagamento': analyzer.get_resumo_formas_pagamento(),
                'curva_abc': analyzer.get_curva_abc(),
                'impostos': analyzer.get_total_impostos(),
                'condicao_menor_28': analyzer.get_fornecedores_condicao_menor_28_dias(),
                'valor_total_pago': self.data_handler.df_pagamentos['Valor Pago'].sum(),
                'qtd_total_nffs': len(self.data_handler.df_pagamentos)
            }
            
            # Gerar HTML
            html = generate_payment_report_html(payment_data)
            self._render_html(html)
            
        except Exception as e:
            print(f"Erro ao carregar relatório de pagamentos: {e}")
            traceback.print_exc()
            html = f"""
            <div class='header'>
                <h1><i class="fas fa-money-bill-wave"></i> Relatório de Pagamentos</h1>
            </div>
            <div style='padding: 40px; text-align: center;'>
                <h2>Erro ao carregar dados</h2>
                <p style='color: #f56565;'>{str(e)}</p>
                <pre style='text-align: left; background: #f7fafc; padding: 16px; border-radius: 8px; margin-top: 20px;'>{traceback.format_exc()}</pre>
            </div>
            """
            self._render_html(html)
    
    def show_supplier_detail(self, supplier_name: str):
        """Mostra detalhes do fornecedor."""
        ficha = self.data_handler.get_ficha_fornecedor(supplier_name)
        total_pago = ficha.get('total_pago', 0)
        total_pago_fmt = f"R$ {total_pago:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        html = f"""<div class='header'><h1>{ficha.get('nome', '')}</h1><p>CNPJ: {ficha.get('cnpj', '')}</p></div>
                   <div style='padding: 24px;'><p><strong>Total Pago:</strong> {total_pago_fmt}</p>
                   <p><strong>Qtd Notas:</strong> {ficha.get('qtd_notas', 0)}</p></div>
                   <button onclick='goBack()' style='margin: 24px; padding: 10px 20px; background: #2d3748; color: white; border: none; border-radius: 8px; cursor: pointer;'>Voltar</button>
                   <script>function goBack() {{ if (window.pyBridge) window.pyBridge.goBack(); }}</script>"""
        self._render_html(html)
    
    def import_data(self):
        """Importa dados de arquivo."""
        # Diálogo para escolher tipo de importação
        from PyQt5.QtWidgets import QInputDialog
        
        tipos = ["Estoque Periódico", "Pagamentos/Fornecedores"]
        tipo, ok = QInputDialog.getItem(
            self, "Importar Dados", "Selecione o tipo de dados:", tipos, 0, False
        )
        
        if not ok:
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo", "", "Excel (*.xlsx *.xls);;CSV (*.csv)"
        )
        
        if file_path:
            try:
                if tipo == "Estoque Periódico":
                    if self.data_handler.load_estoque_file(file_path):
                        QMessageBox.information(self, "Sucesso", "Dados de estoque importados com sucesso!")
                        self.load_estoque()
                    else:
                        QMessageBox.warning(self, "Erro", "Erro ao importar o arquivo de estoque.")
                else:
                    if self.data_handler.load_file(file_path):
                        QMessageBox.information(self, "Sucesso", "Dados de pagamentos importados com sucesso!")
                        self.load_dashboard()
                    else:
                        QMessageBox.warning(self, "Erro", "Erro ao importar o arquivo de pagamentos.")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao processar arquivo:\n{str(e)}")
    
    def export_report(self):
        """Exporta relatório executivo profissional em PDF."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Salvar Relatório Executivo PDF", 
            f"relatorio_executivo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", 
            "PDF (*.pdf)"
        )
        
        if file_path:
            # Mostrar diálogo de progresso
            progress = QProgressDialog("Gerando relatório executivo...", None, 0, 0, self)
            progress.setWindowTitle("Exportando PDF")
            progress.setWindowModality(Qt.WindowModal)
            progress.show()
            QApplication.processEvents()
            
            try:
                if self.data_handler.exportar_pdf_executivo(file_path, self.organizacao_selecionada):
                    progress.close()
                    QMessageBox.information(
                        self, 
                        "Sucesso", 
                        f"📊 Relatório Executivo salvo com sucesso!\n\n{file_path}"
                    )
                else:
                    progress.close()
                    QMessageBox.warning(self, "Erro", "Erro ao exportar o relatório PDF.")
            except Exception as e:
                progress.close()
                QMessageBox.critical(self, "Erro", f"Erro ao gerar PDF:\n{str(e)}")

    # ═══════════════════════════════════════════════════════════════════════════
    # MÉTODOS PIPEFY - MÚLTIPLAS CATEGORIAS
    # ═══════════════════════════════════════════════════════════════════════════
    
    def carregar_dados_pipefy_categoria(self, categoria):
        """Inicia o carregamento dos dados do Pipefy para uma categoria específica."""
        pipe_data = self.pipefy_data.get(categoria, {})
        
        if pipe_data.get('loading'):
            return
        
        self.pipefy_data[categoria]['loading'] = True
        self.current_pipefy_category = categoria
        titulo = self.pipefy_titulos.get(categoria, categoria)
        
        # Mostrar loading - usando CSS inline para evitar problemas de parsing
        spin_style = "@keyframes spin { to { transform: rotate(360deg); } }"
        loading_html = "<div style='padding: 24px;'>" + \
            "<div style='background: white; border-radius: 16px; padding: 48px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); text-align: center;'>" + \
            "<div style='display: inline-block; width: 60px; height: 60px; border: 4px solid #e2e8f0; border-top-color: #2d3748; border-radius: 50%; animation: spin 1s linear infinite;'></div>" + \
            "<h2 style='color: #2d3748; margin: 24px 0 16px;'>Carregando dados do Pipefy...</h2>" + \
            "<p id='loading-status' style='color: #666; font-size: 1.1rem;'>" + titulo + "</p>" + \
            "</div></div><style>" + spin_style + "</style>"
        self._render_html(loading_html)
        QApplication.processEvents()
        
        try:
            # Autenticar
            if not self.pipefy_client:
                self.pipefy_client = PipefyClient(
                    PIPEFY_CLIENT_ID,
                    PIPEFY_CLIENT_SECRET,
                    PIPEFY_TOKEN_URL
                )
                self.pipefy_client.authenticate()
            
            # Obter ID do pipe para a categoria
            pipe_id = self.pipefy_pipes.get(categoria, PIPE_CONTRATACAO_SERVICOS)
            
            # Iniciar thread de carregamento
            self.pipefy_thread = PipefyLoadThread(self.pipefy_client, pipe_id)
            self.pipefy_thread.finished.connect(lambda data: self.on_pipefy_categoria_loaded(data, categoria))
            self.pipefy_thread.error.connect(lambda err: self.on_pipefy_categoria_error(err, categoria))
            self.pipefy_thread.progress.connect(self.on_pipefy_progress)
            self.pipefy_thread.start()
            
        except Exception as e:
            self.pipefy_data[categoria]['loading'] = False
            QMessageBox.critical(self, "Erro de Autenticação", f"Erro ao conectar ao Pipefy:\n{str(e)}")
            self.load_atividades(categoria)
    
    def on_pipefy_categoria_loaded(self, data, categoria):
        """Callback quando dados do Pipefy são carregados para uma categoria."""
        self.pipefy_data[categoria]['loading'] = False
        self.pipefy_data[categoria]['cards'] = data.get('cards', [])
        self.pipefy_data[categoria]['name'] = data.get('name', self.pipefy_titulos.get(categoria, ''))
        
        # REGRA ESPECÍFICA: Atribuição automática de responsáveis por categoria e fase
        if categoria in ['compras_csc', 'compras_locais', 'reserva_materiais']:
            andre_assignee = {
                'id': 'andre_castro_souza',
                'name': 'André Castro de Souza',
                'email': 'acsouza@essencis.com.br'
            }
            
            eviane_assignee = {
                'id': 'eviane',
                'name': 'Eviane',
                'email': 'eviane@essencis.com.br'
            }
            
            for card in self.pipefy_data[categoria]['cards']:
                # Reserva de Materiais: TODOS os cards são da Eviane
                if categoria == 'reserva_materiais':
                    card['assignees'] = [eviane_assignee]
                    continue
                
                current_phase = card.get('current_phase', {})
                if current_phase:
                    phase_name = current_phase.get('name', '').lower()
                    
                    # Compras CSC: Cotação e Pedido de Compra Parcial sempre André
                    # Outros cards sem responsável também recebem André
                    if categoria == 'compras_csc':
                        if 'cotação' in phase_name or 'cotacao' in phase_name or 'pedido de compra parcial' in phase_name:
                            card['assignees'] = [andre_assignee]
                        elif not card.get('assignees'):
                            # Se não tem responsável, atribui André como padrão
                            card['assignees'] = [andre_assignee]
                    
                    # Compras Locais: Aprovação Pedido e Cotação sempre André
                    # Outros cards sem responsável também recebem André
                    elif categoria == 'compras_locais':
                        if 'aprovação pedido' in phase_name or 'aprovacao pedido' in phase_name or 'cotação' in phase_name or 'cotacao' in phase_name:
                            card['assignees'] = [andre_assignee]
                        elif not card.get('assignees'):
                            # Se não tem responsável, atribui André como padrão
                            card['assignees'] = [andre_assignee]
        
        # Compatibilidade retroativa para compras_servicos
        if categoria == 'compras_servicos':
            self.pipefy_cards = self.pipefy_data[categoria]['cards']
            self.pipefy_pipe_name = self.pipefy_data[categoria]['name']
            self.pipefy_loading = False
        
        # Recarregar a aba com os dados
        self.load_atividades(categoria)
        
        titulo = self.pipefy_titulos.get(categoria, categoria)
        cards_count = len(self.pipefy_data[categoria]['cards'])
        QMessageBox.information(
            self, 
            "Dados Carregados", 
            f"{cards_count} cards carregados de:\n{titulo}"
        )
    
    def on_pipefy_categoria_error(self, error_msg, categoria):
        """Callback de erro do Pipefy para uma categoria."""
        self.pipefy_data[categoria]['loading'] = False
        if categoria == 'compras_servicos':
            self.pipefy_loading = False
        QMessageBox.critical(self, "Erro Pipefy", f"Erro ao carregar dados:\n{error_msg}")
        self.load_atividades(categoria)
    
    def carregar_dados_pipefy(self):
        """Inicia o carregamento dos dados do Pipefy (compatibilidade retroativa)."""
        self.carregar_dados_pipefy_categoria('compras_servicos')
    
    def on_pipefy_loaded(self, data):
        """Callback quando dados do Pipefy são carregados (compatibilidade retroativa)."""
        self.on_pipefy_categoria_loaded(data, 'compras_servicos')
    
    def on_pipefy_error(self, error_msg):
        """Callback de erro do Pipefy (compatibilidade retroativa)."""
        self.on_pipefy_categoria_error(error_msg, 'compras_servicos')
    
    def on_pipefy_progress(self, msg):
        """Callback de progresso do Pipefy."""
        # Atualizar status na interface se possível
        pass
    
    def atualizar_dados_pipefy(self):
        """Recarrega os dados do Pipefy para a categoria atual."""
        categoria = self.current_pipefy_category
        self.pipefy_data[categoria]['cards'] = []  # Limpar cache
        if categoria == 'compras_servicos':
            self.pipefy_cards = []
        self.carregar_dados_pipefy_categoria(categoria)
    
    def enviar_email_pipefy(self):
        """Envia email com relatório PDF da categoria atual."""
        categoria = self.current_pipefy_category
        cards = self.pipefy_data.get(categoria, {}).get('cards', [])
        titulo = self.pipefy_titulos.get(categoria, categoria)
        
        # Verificar se há destinatários para esta categoria
        destinatarios = EMAIL_RECIPIENTS.get(categoria, [])
        if not destinatarios:
            QMessageBox.warning(
                self,
                "Aviso",
                f"Não há destinatários configurados para a categoria {titulo}."
            )
            return
        
        if not cards:
            QMessageBox.warning(
                self,
                "Aviso",
                "Não há dados para enviar. Carregue os dados primeiro."
            )
            return
        
        if not self.user_email or not self.user_password:
            QMessageBox.warning(
                self,
                "Erro",
                "Credenciais de email não disponíveis. Faça login novamente."
            )
            return
        
        # Confirmar envio
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Question)
        msg_box.setWindowTitle("Confirmar Envio")
        msg_box.setText(f"Enviar relatório de {titulo}?")
        msg_box.setInformativeText(
            f"Destinatários: {', '.join(destinatarios)}\n"
            f"Total de registros: {len(cards)}"
        )
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.Yes)
        
        if msg_box.exec_() != QMessageBox.Yes:
            return
        
        # Criar PDF temporário
        progress = QProgressDialog("Preparando email...", None, 0, 3, self)
        progress.setWindowTitle("Enviando Email")
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        
        try:
            progress.setValue(1)
            progress.setLabelText("Gerando relatório PDF completo...")
            QApplication.processEvents()
            
            # Criar arquivo PDF temporário com o mesmo nome do relatório
            nome_arquivo = categoria.replace('_', ' ').title().replace(' ', '_')
            temp_pdf = os.path.join(
                tempfile.gettempdir(),
                f"relatorio_{nome_arquivo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            )
            
            # Gerar o PDF usando a mesma lógica do exportar_pdf_pipefy
            # Salvar em arquivo temporário ao invés de perguntar ao usuário
            self._gerar_pdf_pipefy(temp_pdf, categoria, cards, titulo, progress)
            
            progress.setValue(2)
            progress.setLabelText("Enviando email...")
            QApplication.processEvents()
            
            # Preparar corpo do email
            corpo_html = f"""
            <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #2d3748;">Relatório Automático - {titulo}</h2>
                <p>Prezados,</p>
                <p>Segue em anexo o relatório completo de <strong>{titulo}</strong> gerado automaticamente pelo sistema.</p>
                
                <p><strong>Resumo:</strong></p>
                <ul>
                    <li>Total de registros: {len(cards)}</li>
                    <li>Data/Hora: {datetime.now().strftime('%d/%m/%Y às %H:%M')}</li>
                </ul>
                
                <p style="color: #666; font-size: 12px; margin-top: 30px;">
                    <em>Este é um email automático gerado pelo Painel Inteligente de Gestão.</em><br>
                    <em>Não responda a este email.</em>
                </p>
            </body>
            </html>
            """
            
            # Enviar email
            sucesso = enviar_email_relatorio(
                remetente_email=self.user_email,
                remetente_senha=self.user_password,
                destinatarios=destinatarios,
                assunto=f"Relatório Automático - {titulo} - {datetime.now().strftime('%d/%m/%Y')}",
                corpo=corpo_html,
                anexo_pdf=temp_pdf
            )
            
            progress.setValue(3)
            
            if sucesso:
                QMessageBox.information(
                    self,
                    "Sucesso",
                    f"✅ Email enviado com sucesso!\n\nDestinatários: {', '.join(destinatarios)}\n\nO relatório completo foi anexado ao email."
                )
            else:
                QMessageBox.critical(
                    self,
                    "Erro",
                    "Falha ao enviar o email. Verifique suas credenciais e conexão."
                )
            
            # Remover arquivo temporário
            try:
                if os.path.exists(temp_pdf):
                    os.remove(temp_pdf)
            except:
                pass
                
        except Exception as e:
            QMessageBox.critical(
                self,
                "Erro",
                f"Erro ao enviar email:\n{str(e)}"
            )
        finally:
            progress.close()
    
    def exportar_pdf_pipefy(self):
        """Exporta os dados do Pipefy para PDF executivo."""
        categoria = self.current_pipefy_category
        cards = self.pipefy_data.get(categoria, {}).get('cards', [])
        titulo = self.pipefy_titulos.get(categoria, categoria)
        
        if not cards:
            QMessageBox.warning(self, "Aviso", "Não há dados para exportar. Carregue os dados primeiro.")
            return
        
        # Nome do arquivo baseado na categoria
        nome_arquivo = categoria.replace('_', ' ').title().replace(' ', '_')
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Salvar Relatório Executivo PDF", 
            f"relatorio_{nome_arquivo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", 
            "PDF (*.pdf)"
        )
        
        if not file_path:
            return
        
        # Mostrar progresso
        progress = QProgressDialog("Gerando relatório executivo Pipefy...", None, 0, 0, self)
        progress.setWindowTitle("Exportando PDF")
        progress.setWindowModality(Qt.WindowModal)
        progress.show()
        QApplication.processEvents()
        
        try:
            # Chamar método auxiliar para gerar o PDF
            self._gerar_pdf_pipefy(file_path, categoria, cards, titulo, progress)
            
            progress.close()
            QMessageBox.information(
                self, 
                "Exportação Concluída", 
                f"📊 Relatório PDF exportado com sucesso!\n\n{file_path}\n\nTotal: {len(cards)} registros"
            )
            
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Erro de Exportação", f"Erro ao exportar PDF:\n{str(e)}")
    
    def _gerar_pdf_pipefy(self, file_path, categoria, cards, titulo, progress_dialog=None):
        """Gera o PDF do Pipefy (método auxiliar usado por exportar e enviar email)."""
        from datetime import timezone
        
        # Configurar documento PDF
        doc = SimpleDocTemplate(
            file_path,
            pagesize=landscape(A4),
            rightMargin=1.5*cm,
            leftMargin=1.5*cm,
            topMargin=1.5*cm,
            bottomMargin=1.5*cm
        )
        
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'TitlePipefy', parent=styles['Heading1'],
            fontSize=22, textColor=colors.HexColor('#2d3748'),
            spaceAfter=20, alignment=TA_CENTER, fontName='Helvetica-Bold'
        )
        
        section_style = ParagraphStyle(
            'SectionPipefy', parent=styles['Heading2'],
            fontSize=14, textColor=colors.HexColor('#2d3748'),
            spaceBefore=20, spaceAfter=12, fontName='Helvetica-Bold'
        )
        
        normal_style = ParagraphStyle(
            'NormalPipefy', parent=styles['Normal'],
            fontSize=9, textColor=colors.HexColor('#4a5568'), spaceAfter=6
        )
        
        kpi_value_style = ParagraphStyle(
            'KPIPipefy', parent=styles['Normal'],
            fontSize=16, textColor=colors.HexColor('#2d3748'),
            fontName='Helvetica-Bold', alignment=TA_CENTER
        )
        
        kpi_label_style = ParagraphStyle(
            'KPILabelPipefy', parent=styles['Normal'],
            fontSize=8, textColor=colors.HexColor('#718096'), alignment=TA_CENTER
        )
        
        elements = []
        
        # ═══ ANÁLISE DOS DADOS (primeiro para ter os valores) ═══
        hoje = datetime.now(timezone.utc)
        ano_atual = hoje.year
        
        total_cards = len(cards)
        concluidos = sum(1 for c in cards if c.get('finished_at'))
        em_andamento = total_cards - concluidos
        taxa_conclusao = (concluidos / total_cards * 100) if total_cards > 0 else 0
        
        # Análise de datas para período
        datas_criacao = []
        datas_finalizacao = []
        cards_vencidos = []
        cards_vence_7_dias = []
        cards_pendentes = []
        tempos_ciclo = []  # Para calcular tempo médio de atendimento
        
        vencidos = 0
        vence_7_dias = 0
        for card in cards:
                # Coletar datas de criação
                created = card.get('created_at')
                dt_created = None
                if created:
                    try:
                        if 'T' in str(created):
                            dt_created = datetime.fromisoformat(created.replace('Z', '+00:00'))
                        else:
                            dt_created = datetime.strptime(str(created)[:10], '%Y-%m-%d').replace(tzinfo=timezone.utc)
                        datas_criacao.append(dt_created)
                    except:
                        pass
                
                # Coletar datas de finalização e calcular tempo de ciclo
                finished = card.get('finished_at')
                if finished:
                    try:
                        if 'T' in str(finished):
                            dt_finished = datetime.fromisoformat(finished.replace('Z', '+00:00'))
                        else:
                            dt_finished = datetime.strptime(str(finished)[:10], '%Y-%m-%d').replace(tzinfo=timezone.utc)
                        datas_finalizacao.append(dt_finished)
                        
                        # Calcular tempo de ciclo se temos data de criação e finalização no ano atual
                        if dt_created and dt_finished.year == ano_atual:
                            tempo_dias = (dt_finished - dt_created).days
                            if tempo_dias >= 0:
                                tempos_ciclo.append(tempo_dias)
                    except:
                        pass
                    continue  # Card concluído, não verificar vencimento
                
                # Card pendente
                cards_pendentes.append(card)
                
                # Verificar vencimento por tempo na fase de Cotação (compras_locais e compras_csc)
                if verificar_vencimento_fase_cotacao(card, categoria):
                    vencidos += 1
                    cards_vencidos.append(card)
                    continue  # Já contou como vencido, não verificar due_date
                
                # Para compras_locais e compras_servicos: APENAS fase Cotação é vencimento
                # Para outras categorias: verificar due_date
                if categoria not in ['compras_locais', 'compras_servicos']:
                    due = card.get('due_date')
                    if due:
                        try:
                            if 'T' in str(due):
                                due_dt = datetime.fromisoformat(due.replace('Z', '+00:00'))
                            else:
                                due_dt = datetime.strptime(str(due)[:10], '%Y-%m-%d').replace(tzinfo=timezone.utc)
                            if due_dt < hoje:
                                vencidos += 1
                                cards_vencidos.append(card)
                            elif (due_dt - hoje).days <= 7:
                                vence_7_dias += 1
                                cards_vence_7_dias.append(card)
                        except:
                            pass
        
        # Calcular tempo médio de atendimento
        tempo_medio = sum(tempos_ciclo) / len(tempos_ciclo) if tempos_ciclo else 0
        
        # Fases
        fases_count = {}
        for card in cards:
            phase = card.get('current_phase', {})
            phase_name = phase.get('name', 'Sem Fase') if phase else 'Sem Fase'
            fases_count[phase_name] = fases_count.get(phase_name, 0) + 1
        
        # Período dos dados
            periodo_texto = ""
            if datas_criacao:
                data_mais_antiga = min(datas_criacao).strftime("%d/%m/%Y")
                data_mais_recente = max(datas_criacao).strftime("%d/%m/%Y")
                periodo_texto = f"{data_mais_antiga} a {data_mais_recente}"
            else:
                periodo_texto = "Período não identificado"
            
            data_geracao = datetime.now().strftime("%d/%m/%Y às %H:%M")
            
            # ═══ CAPA COM RESUMO EXECUTIVO ═══
            elements.append(Spacer(1, 1.5*cm))
            elements.append(Paragraph(f"📋 RELATÓRIO", title_style))
            elements.append(Paragraph(titulo, ParagraphStyle(
                'Subtitle', parent=styles['Normal'],
                fontSize=14, textColor=colors.HexColor('#718096'),
                alignment=TA_CENTER, spaceAfter=20
            )))
            elements.append(Paragraph(f"Gerado em: {data_geracao}", ParagraphStyle(
                'DataGeracao', parent=styles['Normal'],
                fontSize=10, textColor=colors.HexColor('#a0aec0'),
                alignment=TA_CENTER, spaceAfter=30
            )))
            
            # ═══ VISÃO GERAL DO PERÍODO (na capa) ═══
            elements.append(Paragraph("📊 Visão Geral do Período", ParagraphStyle(
                'VisaoGeralTitle', parent=styles['Heading2'],
                fontSize=14, textColor=colors.HexColor('#2d3748'),
                fontName='Helvetica-Bold', spaceAfter=12
            )))
            
            # Texto descritivo
            resumo_texto = f"""Este relatório apresenta a análise consolidada de <b>{total_cards} processos</b> no sistema. 
            Atualmente, <b>{concluidos} processos</b> foram finalizados (<b>{taxa_conclusao:.1f}%</b> de conclusão) 
            e <b>{em_andamento}</b> permanecem em andamento."""
            
            elements.append(Paragraph(resumo_texto, ParagraphStyle(
                'ResumoTexto', parent=styles['Normal'],
                fontSize=11, textColor=colors.HexColor('#4a5568'),
                spaceAfter=15, leading=16
            )))
            
            # Alertas se houver vencidos ou próximos de vencer
            if vencidos > 0 or vence_7_dias > 0:
                alerta_partes = []
                if vencidos > 0:
                    alerta_partes.append(f"Existem <b>{vencidos} processos vencidos</b> que requerem ação imediata.")
                if vence_7_dias > 0:
                    alerta_partes.append(f"Há também <b>{vence_7_dias} processos</b> que vencerão nos próximos 7 dias.")
                if tempos_ciclo:
                    alerta_partes.append(f"O tempo médio de atendimento em {ano_atual} é de <b>{tempo_medio:.0f} dias</b>.")
                
                alerta_texto = "⚠️ <b>Atenção:</b> " + " ".join(alerta_partes)
                
                # Criar tabela para o alerta com fundo
                alerta_table = Table([[Paragraph(alerta_texto, ParagraphStyle(
                    'AlertaTexto', parent=styles['Normal'],
                    fontSize=10, textColor=colors.HexColor('#744210'),
                    leading=14
                ))]], colWidths=[24*cm])
                alerta_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#fffaf0')),
                    ('BOX', (0, 0), (-1, -1), 2, colors.HexColor('#ed8936')),
                    ('LEFTPADDING', (0, 0), (-1, -1), 15),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 15),
                    ('TOPPADDING', (0, 0), (-1, -1), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ]))
                elements.append(alerta_table)
            elif tempos_ciclo:
                # Se não há alertas mas tem tempo médio
                info_texto = f"✅ Não há processos vencidos. O tempo médio de atendimento em {ano_atual} é de <b>{tempo_medio:.0f} dias</b>."
                info_table = Table([[Paragraph(info_texto, ParagraphStyle(
                    'InfoTexto', parent=styles['Normal'],
                    fontSize=10, textColor=colors.HexColor('#276749'),
                    leading=14
                ))]], colWidths=[24*cm])
                info_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f0fff4')),
                    ('BOX', (0, 0), (-1, -1), 2, colors.HexColor('#48bb78')),
                    ('LEFTPADDING', (0, 0), (-1, -1), 15),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 15),
                    ('TOPPADDING', (0, 0), (-1, -1), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ]))
                elements.append(info_table)
            
            elements.append(Spacer(1, 1*cm))
            
            # ═══ KPIs NA CAPA ═══
            kpi_data = [
                [
                    Paragraph(str(total_cards), kpi_value_style),
                    Paragraph(str(concluidos), kpi_value_style),
                    Paragraph(str(em_andamento), kpi_value_style),
                    Paragraph(f"{taxa_conclusao:.0f}%", kpi_value_style),
                    Paragraph(str(vencidos), kpi_value_style),
                ],
                [
                    Paragraph('TOTAL', kpi_label_style),
                    Paragraph('CONCLUÍDOS', kpi_label_style),
                    Paragraph('EM ANDAMENTO', kpi_label_style),
                    Paragraph('TX CONCLUSÃO', kpi_label_style),
                    Paragraph('VENCIDOS', kpi_label_style),
                ]
            ]
            
            kpi_table = Table(kpi_data, colWidths=[4.5*cm]*5)
            kpi_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f7fafc')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
                ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('TOPPADDING', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                # Destacar vencidos em vermelho se houver
                ('TEXTCOLOR', (4, 0), (4, 0), colors.HexColor('#c53030') if vencidos > 0 else colors.HexColor('#2d3748')),
            ]))
            elements.append(kpi_table)
            
            elements.append(Spacer(1, 0.5*cm))
            
            # Info do período
            elements.append(Paragraph(f"📅 Período: {periodo_texto}", ParagraphStyle(
                'PeriodoInfo', parent=styles['Normal'],
                fontSize=9, textColor=colors.HexColor('#718096'),
                alignment=TA_CENTER, spaceAfter=10
            )))
            
            elements.append(PageBreak())
            
            # ═══ GRÁFICO DE FASES ═══
            elements.append(Paragraph("📈 DISTRIBUIÇÃO POR FASE", section_style))
            
            if fases_count:
                fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
                
                fases_sorted = sorted(fases_count.items(), key=lambda x: x[1], reverse=True)
                fases_names = [f[0][:25] for f in fases_sorted]
                fases_vals = [f[1] for f in fases_sorted]
                
                cores_fases = plt.cm.Blues(np.linspace(0.3, 0.9, len(fases_names)))
                
                bars = ax1.barh(fases_names[::-1], fases_vals[::-1], color=cores_fases[::-1])
                ax1.set_xlabel('Quantidade de Cards')
                ax1.set_title('Cards por Fase', fontweight='bold')
                ax1.spines['top'].set_visible(False)
                ax1.spines['right'].set_visible(False)
                
                for bar, val in zip(bars, fases_vals[::-1]):
                    ax1.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2, 
                            str(val), va='center', fontweight='bold')
                
                # Pizza de status
                status_vals = [concluidos, em_andamento, vencidos]
                status_labels = ['Concluídos', 'Em Andamento', 'Vencidos']
                status_cores = ['#48bb78', '#ed8936', '#f56565']
                
                ax2.pie([v for v in status_vals if v > 0], 
                       labels=[l for l, v in zip(status_labels, status_vals) if v > 0],
                       colors=[c for c, v in zip(status_cores, status_vals) if v > 0],
                       autopct='%1.1f%%', startangle=90)
                ax2.set_title('Status dos Cards', fontweight='bold')
                
                plt.tight_layout()
                
                img_buffer = io.BytesIO()
                plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight',
                           facecolor='white', edgecolor='none')
                img_buffer.seek(0)
                plt.close(fig)
                
                img = Image(img_buffer, width=24*cm, height=9*cm)
                elements.append(img)
            
            elements.append(PageBreak())
            
            # ═══ SEÇÃO DE CARDS VENCIDOS (se houver) ═══
            if cards_vencidos:
                elements.append(Paragraph("🚨 CARDS VENCIDOS - ATENÇÃO IMEDIATA", section_style))
                elements.append(Paragraph(
                    f"Existem {len(cards_vencidos)} cards com prazo vencido que requerem ação imediata.",
                    ParagraphStyle('AlertText', parent=styles['Normal'],
                        fontSize=10, textColor=colors.HexColor('#c53030'), spaceAfter=10)
                ))
                
                vencidos_header = ['#', 'Título', 'Fase', 'Responsável', 'Vencimento', 'Dias Atraso']
                vencidos_data = [vencidos_header]
                
                for idx, card in enumerate(cards_vencidos[:20], 1):
                    titulo_card = card.get('title', '')[:35]
                    if len(card.get('title', '')) > 35:
                        titulo_card += '...'
                    
                    phase = card.get('current_phase', {})
                    fase_nome = phase.get('name', 'N/A')[:18] if phase else 'N/A'
                    
                    assignees = card.get('assignees', [])
                    resp = assignees[0].get('name', '')[:18] if assignees else 'Sem responsável'
                    
                    due = card.get('due_date', '')
                    venc = due[:10] if due else 'N/A'
                    
                    # Calcular dias de atraso
                    dias_atraso = ''
                    if due:
                        try:
                            if 'T' in str(due):
                                due_dt = datetime.fromisoformat(due.replace('Z', '+00:00'))
                            else:
                                due_dt = datetime.strptime(str(due)[:10], '%Y-%m-%d').replace(tzinfo=timezone.utc)
                            dias_atraso = str((hoje - due_dt).days) + ' dias'
                        except:
                            dias_atraso = '-'
                    
                    vencidos_data.append([str(idx), titulo_card, fase_nome, resp, venc, dias_atraso])
                
                if len(cards_vencidos) > 20:
                    vencidos_data.append(['...', f'+ {len(cards_vencidos) - 20} vencidos', '', '', '', ''])
                
                vencidos_table = Table(vencidos_data, colWidths=[1*cm, 8*cm, 4.5*cm, 4.5*cm, 3*cm, 3*cm])
                vencidos_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#c53030')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#c53030')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#feb2b2')),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#fff5f5')),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ]))
                elements.append(vencidos_table)
                elements.append(Spacer(1, 1*cm))
            
            # ═══ CARDS QUE VENCEM EM 7 DIAS ═══
            if cards_vence_7_dias:
                elements.append(Paragraph("⚠️ CARDS COM VENCIMENTO PRÓXIMO (7 dias)", section_style))
                
                prox_header = ['#', 'Título', 'Fase', 'Responsável', 'Vencimento']
                prox_data = [prox_header]
                
                for idx, card in enumerate(cards_vence_7_dias[:15], 1):
                    titulo_card = card.get('title', '')[:38]
                    if len(card.get('title', '')) > 38:
                        titulo_card += '...'
                    
                    phase = card.get('current_phase', {})
                    fase_nome = phase.get('name', 'N/A')[:18] if phase else 'N/A'
                    
                    assignees = card.get('assignees', [])
                    resp = assignees[0].get('name', '')[:18] if assignees else 'Sem responsável'
                    
                    due = card.get('due_date', '')
                    venc = due[:10] if due else 'N/A'
                    
                    prox_data.append([str(idx), titulo_card, fase_nome, resp, venc])
                
                prox_table = Table(prox_data, colWidths=[1*cm, 9*cm, 5*cm, 5*cm, 3*cm])
                prox_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#c05621')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#c05621')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#feebc8')),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#fffaf0')),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ]))
                elements.append(prox_table)
                elements.append(Spacer(1, 1*cm))
            
            # ═══ TABELA DE CARDS PENDENTES (exceto vencidos e próximos) ═══
            # Filtrar cards pendentes que não estão nas listas de vencidos ou próximos
            cards_vencidos_ids = {c.get('id') for c in cards_vencidos}
            cards_prox_ids = {c.get('id') for c in cards_vence_7_dias}
            cards_outros_pendentes = [c for c in cards_pendentes 
                                       if c.get('id') not in cards_vencidos_ids 
                                       and c.get('id') not in cards_prox_ids]
            
            if cards_outros_pendentes:
                elements.append(Paragraph("📋 DEMAIS CARDS PENDENTES", section_style))
                
                table_header = ['#', 'Título', 'Fase', 'Responsável', 'Vencimento']
                table_data = [table_header]
                
                for idx, card in enumerate(cards_outros_pendentes[:40], 1):
                    titulo_card = card.get('title', '')[:40]
                    if len(card.get('title', '')) > 40:
                        titulo_card += '...'
                    
                    phase = card.get('current_phase', {})
                    fase_nome = phase.get('name', 'N/A')[:20] if phase else 'N/A'
                    
                    assignees = card.get('assignees', [])
                    resp = assignees[0].get('name', '')[:20] if assignees else 'N/A'
                    
                    due = card.get('due_date', '')
                    venc = due[:10] if due else 'N/A'
                    
                    table_data.append([str(idx), titulo_card, fase_nome, resp, venc])
                
                if len(cards_outros_pendentes) > 40:
                    table_data.append(['...', f'+ {len(cards_outros_pendentes) - 40} cards adicionais', '', '', ''])
                
                cards_table = Table(table_data, colWidths=[1*cm, 10*cm, 5*cm, 5*cm, 3*cm])
                cards_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2d3748')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e2e8f0')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f7fafc')]),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ]))
                elements.append(cards_table)
            elif not cards_vencidos and not cards_vence_7_dias and not cards_pendentes:
                elements.append(Paragraph("📋 CARDS PENDENTES", section_style))
                elements.append(Paragraph("✅ Não há cards pendentes no momento. Todos os cards foram concluídos!", 
                    ParagraphStyle('SuccessText', parent=styles['Normal'],
                        fontSize=11, textColor=colors.HexColor('#276749'), spaceAfter=10)))
            
            # Rodapé
            elements.append(Spacer(1, 2*cm))
            footer_style = ParagraphStyle('Footer', parent=styles['Normal'],
                fontSize=8, textColor=colors.HexColor('#a0aec0'), alignment=TA_CENTER)
            elements.append(Paragraph("─" * 80, footer_style))
            elements.append(Paragraph(f"Relatório - {titulo} | {data_geracao}", footer_style))
            
            doc.build(elements)

# ═══════════════════════════════════════════════════════════════════════════════
# PONTO DE ENTRADA
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    """Função principal da aplicação."""
    app = QApplication(sys.argv)
    
    # Mostrar login
    login = LoginScreen()
    if login.exec() != QDialog.DialogCode.Accepted:
        # Usuário cancelou o login
        sys.exit(0)
    
    # Obter credenciais do login
    user_email = login.user_email
    user_password = login.user_password
    
    # Mostrar splash screen
    splash = SplashScreen()
    splash.show()
    
    # Processar eventos enquanto a splash screen está visível
    while splash.isVisible():
        app.processEvents()
        QThread.msleep(50)
    
    # Criar janela principal com dados carregados e credenciais
    window = MainWindow(
        loaded_data=splash.loaded_data,
        user_email=user_email,
        user_password=user_password
    )
    window.showMaximized()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

