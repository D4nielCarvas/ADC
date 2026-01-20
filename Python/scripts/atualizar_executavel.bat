@echo off
chcp 65001 >nul
echo ========================================
echo  ATUALIZADOR DE EXECUTAVEL
echo  ADC
echo ========================================
echo.

cd ..

echo [1/5] Limpando arquivos antigos...
if exist build rd /s /q build
if exist dist rd /s /q dist
if exist __pycache__ rd /s /q __pycache__
if exist *.spec del /q *.spec
echo ✓ Limpeza concluída!
echo.

echo [2/5] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python não encontrado
    pause
    exit /b 1
)
echo ✓ Python encontrado!
echo.

echo [3/5] Verificando dependências...
python -m pip install --quiet --disable-pip-version-check pyinstaller pandas openpyxl xlrd matplotlib
echo ✓ Dependências verificadas!
echo.

echo [4/5] Recompilando executável...
echo Isso pode levar alguns minutos. Aguarde...
echo.
python -m PyInstaller --onefile --windowed --clean --noconfirm --name="ADC" --add-data "src/config.json;src" "src/main.py"

if %errorlevel% neq 0 (
    echo.
    echo ERRO: Falha ao criar executável
    pause
    exit /b 1
)
echo.

echo [5/5] Finalizando...
if exist build rd /s /q build
if exist ADC.spec del /q ADC.spec
echo.

if exist "dist\ADC.exe" (
    echo ========================================
    echo  ✓ EXECUTÁVEL ADC ATUALIZADO COM SUCESSO!
    echo ========================================
    echo.
    echo Local: %cd%\dist\ADC.exe
    echo.
) else (
    echo ERRO: Executável não foi criado
    pause
    exit /b 1
)

pause
