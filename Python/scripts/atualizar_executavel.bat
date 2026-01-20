@echo off
chcp 65001 >nul
echo ========================================
echo  INSTALADOR DE DEPENDÊNCIAS
echo  ADC Pro
echo ========================================
echo.

echo [1/2] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python não encontrado no sistema.
    echo Por favor, instale o Python 3.8 ou superior.
    pause
    exit /b 1
)
echo ✓ Python encontrado!
echo.

echo [2/2] Instalando dependências (requirements.txt)...
python -m pip install --quiet --upgrade --disable-pip-version-check -r requirements.txt
if %errorlevel% neq 0 (
    echo ERRO: Falha ao instalar dependências.
    pause
    exit /b 1
)
echo ✓ Todas as dependências foram instaladas com sucesso!
echo.

echo ========================================
echo  ✓ AMBIENTE PRONTO PARA USO!
echo  Você já pode rodar o ADC via:
echo  python src/main.py
echo ========================================
echo.
pause
