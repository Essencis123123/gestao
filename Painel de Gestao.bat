@echo off
:: ============================================================================
:: Launcher - Painel de Gestao
:: Copia a pasta do servidor para cache local e executa de la.
:: Resultado: abertura rapida mesmo rodando do servidor.
:: ============================================================================
setlocal enabledelayedexpansion

:: --- Configuracao -----------------------------------------------------------
:: Pasta no servidor onde fica a build --onedir
set "SERVIDOR=S:\temp\Painel de Gestao"

:: Cache local (LOCALAPPDATA e persistente entre reboots)
set "CACHE=%LOCALAPPDATA%\PainelGestao\app"

:: Nome do executavel dentro da pasta
set "EXE_NAME=Painel de Gestao.exe"

:: Arquivo de controle de versao (atualizar ao fazer novo build)
set "VERSION_FILE=.build_version"
:: ----------------------------------------------------------------------------

:: Verificar se o servidor esta acessivel
if not exist "%SERVIDOR%\%EXE_NAME%" (
    echo [ERRO] Nao foi possivel acessar o servidor: %SERVIDOR%
    echo Verifique sua conexao de rede.
    pause
    exit /b 1
)

:: Criar pasta de cache se nao existe
if not exist "%CACHE%" mkdir "%CACHE%"

:: Verificar se precisa atualizar (compara versao)
set "NEEDS_UPDATE=0"

:: Checar se o cache existe
if not exist "%CACHE%\%EXE_NAME%" set "NEEDS_UPDATE=1"

:: Checar arquivo de versao
if exist "%SERVIDOR%\%VERSION_FILE%" (
    if not exist "%CACHE%\%VERSION_FILE%" (
        set "NEEDS_UPDATE=1"
    ) else (
        fc /b "%SERVIDOR%\%VERSION_FILE%" "%CACHE%\%VERSION_FILE%" >nul 2>&1
        if errorlevel 1 set "NEEDS_UPDATE=1"
    )
) else (
    :: Sem arquivo de versao, comparar data do exe
    if exist "%CACHE%\%EXE_NAME%" (
        for %%A in ("%SERVIDOR%\%EXE_NAME%") do set "SRV_DATE=%%~tA"
        for %%A in ("%CACHE%\%EXE_NAME%") do set "LOC_DATE=%%~tA"
        if not "!SRV_DATE!"=="!LOC_DATE!" set "NEEDS_UPDATE=1"
    )
)

:: Sincronizar se necessario
if "%NEEDS_UPDATE%"=="1" (
    echo Atualizando Painel de Gestao...
    robocopy "%SERVIDOR%" "%CACHE%" /mir /nfl /ndl /njh /njs /nc /ns /np >nul 2>&1
    if errorlevel 8 (
        echo [ERRO] Falha ao copiar arquivos do servidor.
        pause
        exit /b 1
    )
    echo Atualizado com sucesso.
)

:: Executar do cache local
start "" "%CACHE%\%EXE_NAME%"
exit /b 0
