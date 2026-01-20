@echo off
chcp 65001 >nul
echo ========================================
echo  ATUALIZADOR DE EXECUTÁVEL E DEPENDÊNCIAS
echo  ADC Pro
echo ========================================
echo.

cd /d "%~dp0\.."

echo [1/3] Verificando Ambiente Virtual (.venv)...
if not exist ".venv" (
    echo [!] Criando ambiente virtual...
    python -m venv .venv
)
echo ✓ Ambiente virtual detectado!
echo.

echo [2/3] Atualizando dependências (scripts/requirements.txt)...
call .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r scripts\requirements.txt --upgrade
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar dependências.
    pause
    exit /b 1
)
echo ✓ Todas as dependências foram instaladas com sucesso!
echo.

echo [3/3] Criando Executável (PyInstaller)...
echo Aguarde, isso pode levar alguns minutos...
:: Usando 'python -m PyInstaller' para ignorar bloqueio de Device Guard no .exe
python -m PyInstaller --noconfirm --onedir --windowed --add-data "src/config.json;." "src/main.py"
if %errorlevel% neq 0 (
    echo ERRO: Falha ao gerar o executável.
    pause
    exit /b 1
)

echo ========================================
echo  ✓ PROCESSO CONCLUÍDO!
echo  O executáveis atualizados estão em dist/main
echo ========================================
echo.
pause
