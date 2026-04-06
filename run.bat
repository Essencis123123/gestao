@echo off
echo.
echo ╔═══════════════════════════════════════════════════════╗
echo ║  Painel Inteligente de Gestão v2.0                   ║
echo ║  Sistema de Análise Empresarial                      ║
echo ╚═══════════════════════════════════════════════════════╝
echo.

REM Verificar se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python não está instalado ou não está no PATH
    echo.
    echo Acesse: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✓ Python encontrado
echo.

REM Criar ambiente virtual se não existir
if not exist "venv" (
    echo 📦 Criando ambiente virtual...
    python -m venv venv
    echo ✓ Ambiente virtual criado
)

echo.
echo 📥 Ativando ambiente virtual...
call venv\Scripts\activate.bat

echo ✓ Ambiente ativo
echo.

REM Instalar dependências
echo 📚 Instalando dependências...
pip install -q -r requirements.txt
if errorlevel 1 (
    echo ❌ Erro ao instalar dependências
    pause
    exit /b 1
)
echo ✓ Dependências instaladas
echo.

REM Iniciar aplicação
echo 🚀 Iniciando aplicação...
echo.
python main.py

pause
