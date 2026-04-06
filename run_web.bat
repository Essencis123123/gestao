@echo off
chcp 65001 > nul
cd /d "%~dp0"

REM Ferramenta de Gestão - Versão Web
REM Script para rodar o servidor Flask

echo.
echo ==========================================
echo    Ferramenta de Gestão - Web v2.0
echo ==========================================
echo.
echo. Iniciando servidor na porta 5000...
echo %DATE% %TIME%
echo.
echo. Site: http://localhost:5000
echo. Pressione CTRL+C para parar
echo.
echo ==========================================
echo.

REM Atualizar PATH
for /f "tokens=*" %%A in ('python -c "import sys; print(sys.version)"') do set PYTHON_VERSION=%%A

python app.py

pause
