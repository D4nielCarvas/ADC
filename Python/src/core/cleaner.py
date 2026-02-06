# -*- coding: utf-8 -*-
"""
ADC Core Logic Module

This module contains the core business logic for the Advanced Data Cleaner (ADC) application.
It handles Excel file loading, data cleaning, filtering, and summary generation for e-commerce
product spreadsheets.

Classes:
    ADCLogic: Main class containing all data processing logic
"""
import pandas as pd
import os
import json
from datetime import datetime

class ADCLogic:
    """
    Core business logic for ADC (Advanced Data Cleaner).
    
    This class handles all data processing operations including:
    - Loading and parsing Excel files (.xlsx and .xls)
    - Managing cleaning presets
    - Applying filters and transformations
    - Generating statistical summaries
    
    Attributes:
        presets (list): List of cleaning preset configurations
        cache_excel (dict): Cache for loaded Excel files
    """
    
    def __init__(self):
        """Initialize ADCLogic with presets loaded from configuration file."""
        self.presets = self.carregar_presets()
        self.cache_excel = {}
    
    @staticmethod
    def limpar_valor(x):
        """
        Clean and convert Brazilian currency/numeric strings to float.
        
        Handles formats like:
        - "R$ 1.200,50" -> 1200.50
        - "1.200,50" -> 1200.50
        - "1200.50" -> 1200.50
        
        Args:
            x: Value to clean (str, int, or float)
            
        Returns:
            float: Cleaned numeric value, or 0.0 if conversion fails
        """
        if isinstance(x, str):
            # Remove currency symbol, thousand separators (.), and replace decimal comma with dot
            x = x.replace('R$', '').replace('.', '').replace(',', '.').strip()
        try:
            return float(x)
        except (ValueError, TypeError):
            return 0.0
    
    @staticmethod
    def clean_numeric(val):
        """
        Clean numeric values from DataFrame cells.
        
        Similar to limpar_valor but handles pandas NA values explicitly.
        
        Args:
            val: Value to clean (can be NA, str, int, or float)
            
        Returns:
            float: Cleaned numeric value, or 0.0 if conversion fails
        """
        if pd.isna(val):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        
        # Handle string values
        s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
        try:
            return float(s)
        except (ValueError, TypeError):
            return 0.0

    def carregar_presets(self):
        """
        Load cleaning presets from config/settings.json file.
        
        Tries multiple paths to find the configuration file:
        1. Development path (relative to source file)
        2. Executable distribution path
        3. Fallback path
        
        Returns:
            list: List of preset dictionaries, empty list if file not found
        """
        # Estrategia de busca de arquivo robusta
        caminhos_possiveis = [
            # 1. Desenvolvimento: ../../config/settings.json (relativo a src/core/cleaner.py)
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "settings.json"),
            # 2. Executável (Dist): pasta config ao lado do exe
            os.path.join(os.getcwd(), "config", "settings.json"),
            # 3. Fallback
            "config/settings.json"
        ]
        
        for caminho in caminhos_possiveis:
            if os.path.exists(caminho):
                try:
                    with open(caminho, 'r', encoding='utf-8') as f:
                        return json.load(f).get("presets", [])
                except Exception as e:
                    print(f"Erro ao ler config {caminho}: {e}")
                    continue
        return []

    def salvar_presets(self, presets):
        """
        Save presets list to settings.json file.
        
        Args:
            presets (list): List of preset dictionaries to save
            
        Returns:
            tuple: (success: bool, message: str)
        """
        self.presets = presets
        
        # Tenta salvar no mesmo local que carregou ou no padrão dev
        caminho = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "settings.json")
        if not os.path.exists(os.path.dirname(caminho)):
             # Se não existir (ex: rodando do exe), tenta no cwd
             caminho = os.path.join(os.getcwd(), "config", "settings.json")
        
        try:
            os.makedirs(os.path.dirname(caminho), exist_ok=True)
            with open(caminho, 'w', encoding='utf-8') as f:
                json.dump({"presets": self.presets}, f, indent=4, ensure_ascii=False)
            return True, "Presets salvos com sucesso."
        except Exception as e:
            return False, str(e)

    def carregar_planilha(self, caminho, aba=None, log_callback=None):
        """
        Load Excel spreadsheet with robust error handling.
        
        Features:
        - Automatic engine selection (.xls vs .xlsx)
        - Fallback to alternate engine on failure
        - Empty sheet name handling (loads first sheet)
        
        Args:
            caminho (str): Path to Excel file
            aba (str, optional): Sheet name. None returns ExcelFile object, "" loads first sheet
            log_callback (callable, optional): Callback function for logging
            
        Returns:
            pd.DataFrame or pd.ExcelFile: Loaded data or ExcelFile object
            
        Raises:
            Exception: If file cannot be loaded with any engine
        """
        try:
            if log_callback and aba: log_callback(f"Tentando carregar aba: {aba}")
            
            # Determine engine usage
            engine = 'xlrd' if caminho.lower().endswith('.xls') else 'openpyxl'

            if aba is None:
                # Retorna apenas o objeto para listar abas
                # Try multiple engines for better compatibility
                for eng in [engine, 'openpyxl', 'xlrd', None]:
                    try:
                        return pd.ExcelFile(caminho, engine=eng)
                    except:
                        continue
                raise Exception("Nao foi possivel abrir o arquivo com nenhum engine disponivel")
            
            # FIX: Tratamento para aba vazia (string vazia)
            if aba == "":
                if log_callback: log_callback("[WARNING] Nenhuma aba especificada. Carregando a primeira disponivel.")
                # Try to get sheet names with fallback engines
                xl_file = None
                for eng in [engine, 'openpyxl', 'xlrd', None]:
                    try:
                        xl_file = pd.ExcelFile(caminho, engine=eng)
                        aba = xl_file.sheet_names[0]
                        if log_callback: log_callback(f"[OK] Primeira aba identificada: {aba} (engine: {eng})")
                        break
                    except:
                        continue
                
                if xl_file is None or not aba:
                    raise Exception("Nao foi possivel abrir o arquivo para listar abas")

            # Try primary engine first
            try:
                df = pd.read_excel(caminho, sheet_name=aba, engine=engine)
                if log_callback: log_callback(f"[OK] Planilha carregada com engine {engine}")
                return df
            except Exception as e1:
                if log_callback: log_callback(f"[WARNING] Engine {engine} falhou: {str(e1)[:100]}")
                
                # Try alternate engine
                engine_alt = 'openpyxl' if engine == 'xlrd' else 'xlrd'
                try:
                    df = pd.read_excel(caminho, sheet_name=aba, engine=engine_alt)
                    if log_callback: log_callback(f"[OK] Planilha carregada com engine alternativo {engine_alt}")
                    return df
                except Exception as e2:
                    if log_callback: log_callback(f"[WARNING] Engine {engine_alt} tambem falhou: {str(e2)[:100]}")
                    
                    # Last resort: try without specifying engine (let pandas decide)
                    try:
                        df = pd.read_excel(caminho, sheet_name=aba)
                        if log_callback: log_callback("[OK] Planilha carregada com engine automatico")
                        return df
                    except Exception as e3:
                        # All methods failed, raise detailed error
                        error_msg = f"Falha ao carregar planilha com todos os metodos:\n"
                        error_msg += f"  - {engine}: {str(e1)[:80]}\n"
                        error_msg += f"  - {engine_alt}: {str(e2)[:80]}\n"
                        error_msg += f"  - Auto: {str(e3)[:80]}"
                        raise Exception(error_msg)
            
            return df
        except Exception as e:
            raise Exception(f"Erro ao carregar planilha: {e}")

    def listar_abas(self, caminho, log_callback=None):
        """
        List all sheet names in an Excel file.
        
        Args:
            caminho (str): Path to Excel file
            log_callback (callable, optional): Callback function for logging
            
        Returns:
            list: List of sheet names
            
        Raises:
            Exception: If file cannot be read
        """
        try:
            xl = self.carregar_planilha(caminho) # Retorna ExcelFile neste caso
            if isinstance(xl, pd.ExcelFile):
                 return xl.sheet_names
            # Fallback
            engine = 'xlrd' if caminho.lower().endswith('.xls') else 'openpyxl'
            xl = pd.ExcelFile(caminho, engine=engine)
            return xl.sheet_names
        except Exception as e:
             if log_callback: log_callback(f"Erro ao listar abas: {e}")
             raise e

    def validar_arquivo_entrada(self, caminho, log_callback=None):
        """
        Validate that input file exists and is a valid file.
        
        Args:
            caminho (str): Path to file to validate
            log_callback (callable, optional): Callback function for logging
            
        Returns:
            bool: True if file is valid
            
        Raises:
            FileNotFoundError: If file does not exist
            ValueError: If path is not a file
        """
        if not os.path.exists(caminho):
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
        if not os.path.isfile(caminho):
            raise ValueError(f"O caminho especificado não é um arquivo: {caminho}")
        if log_callback: log_callback(f"[OK] Arquivo de entrada encontrado: {os.path.basename(caminho)}")
        return True

    def validar_indices_colunas(self, df, indices, log_callback=None):
        """
        Validate that column indices are within DataFrame bounds.
        
        Args:
            df (pd.DataFrame): DataFrame to validate against
            indices (list): List of column indices to validate
            log_callback (callable, optional): Callback function for logging
            
        Returns:
            bool: True if all indices are valid
            
        Raises:
            ValueError: If any index is out of bounds
        """
        total_colunas = len(df.columns)
        indices_invalidos = [i for i in indices if i >= total_colunas or i < 0]
        
        if indices_invalidos:
            msg = f"Índices inválidos: {indices_invalidos}. A planilha tem {total_colunas} colunas (0-{total_colunas-1})"
            if log_callback: log_callback(f"[ERROR] {msg}")
            raise ValueError(msg)
        
        if log_callback: log_callback(f"[OK] Indices validados: {indices}")
        return True

    def processar_limpeza(self, caminho_entrada, aba, indices_deletar, opcoes_filtros=None, log_callback=None):
        """
        Main data cleaning pipeline.
        
        Process flow:
        1. Validate input file
        2. Load spreadsheet
        3. Validate column indices
        4. Delete specified columns
        5. Apply additional filters (if provided)
        
        Args:
            caminho_entrada (str): Path to input Excel file
            aba (str): Sheet name to process
            indices_deletar (list): List of column indices to delete (0-based)
            opcoes_filtros (dict, optional): Filter options with keys:
                - remover_duplicadas (bool): Remove duplicate rows
                - remover_vazias (bool): Remove empty rows
                - filtro_valor (dict): {'ativo': bool, 'minimo': float, 'coluna': str}
                - filtro_texto (dict): {'ativo': bool, 'texto': str}
            log_callback (callable, optional): Callback function for logging
            
        Returns:
            pd.DataFrame: Cleaned DataFrame
        """
        # 1. Validar Arquivo
        self.validar_arquivo_entrada(caminho_entrada, log_callback)

        # 2. Carregar
        df = self.carregar_planilha(caminho_entrada, aba, log_callback)
        
        # 3. Validar Índices
        self.validar_indices_colunas(df, indices_deletar, log_callback)

        # 4. Deletar Colunas
        if log_callback: log_callback("Removendo colunas selecionadas...")
        validador_indices = [i for i in indices_deletar if i < len(df.columns)]
        colunas_deletar = [df.columns[i] for i in validador_indices]
        if log_callback: log_callback(f"[OK] Deletando colunas: {colunas_deletar}")
        
        df_limpo = df.drop(df.columns[validador_indices], axis=1)
        
        # 5. Filtros Extras
        if opcoes_filtros:
             return self.aplicar_filtros_adicionais(df_limpo, opcoes_filtros, log_callback)
        
        # Default behavior if no options passed (backward compatibility)
        if log_callback: log_callback("Removendo duplicatas e vazios (padrão)...")
        df_limpo = df_limpo.drop_duplicates()
        df_limpo = df_limpo.dropna(how='all') 
        
        return df_limpo

    def aplicar_filtros_adicionais(self, df, opcoes, log_callback=None):
        linhas_iniciais = len(df)
        
        # Filtro 1: Remover linhas duplicadas
        if opcoes.get('remover_duplicadas'):
            df = df.drop_duplicates()
            if log_callback: log_callback(f"  🔄 Duplicatas removidas: {linhas_iniciais - len(df)}")
        
        # Filtro 2: Remover linhas vazias/incompletas
        if opcoes.get('remover_vazias'):
            linhas_antes = len(df)
            df = df.dropna(how='all')
            # Limpeza agressiva apenas se muitos nas
            # df = df.dropna(thresh=len(df.columns) // 2)
            if log_callback: log_callback(f"  🗑️ Vazias removidas: {linhas_antes - len(df)}")

        # Filtro 3: Filtrar por valor mínimo
        filtro_val = opcoes.get('filtro_valor', {})
        if filtro_val.get('ativo'):
            df = self.filtro_por_valor_minimo(df, filtro_val.get('minimo', 0), filtro_val.get('coluna'), log_callback)

        # Filtro 4: Filtrar por texto
        filtro_txt = opcoes.get('filtro_texto', {})
        if filtro_txt.get('ativo'):
            df = self.filtro_por_texto(df, filtro_txt.get('texto', ''), log_callback)
            
        return df

    def filtro_por_valor_minimo(self, df, valor_min, coluna_filtro, log_callback=None):
        """
        Filter DataFrame rows by minimum value in specified column.
        
        Args:
            df (pd.DataFrame): DataFrame to filter
            valor_min (float): Minimum value threshold
            coluna_filtro (str): Column name to apply filter on
            log_callback (callable, optional): Callback function for logging
            
        Returns:
            pd.DataFrame: Filtered DataFrame
        """
        try:
            if not coluna_filtro or coluna_filtro not in df.columns:
                if log_callback: log_callback(f"  [WARNING] Coluna '{coluna_filtro}' nao disponivel para filtro de valor")
                return df
            
            # Use the static method instead of inline function
            df_temp = df.copy()
            df_temp[coluna_filtro] = df_temp[coluna_filtro].apply(self.limpar_valor)
            
            df_filtrado = df[df_temp[coluna_filtro] >= valor_min]
            removidas = len(df) - len(df_filtrado)
            
            if log_callback and removidas > 0:
                log_callback(f"  [INFO] Removidas {removidas} linhas com {coluna_filtro} < {valor_min}")
            
            return df_filtrado
        except Exception as e:
            if log_callback: log_callback(f"  [WARNING] Erro no filtro de valor: {e}")
            return df

    def filtro_por_texto(self, df, texto, log_callback=None):
        texto = str(texto).strip()
        if not texto: return df
        
        colunas_texto = df.select_dtypes(include=['object']).columns
        if len(colunas_texto) == 0: return df

        mascara = pd.Series(False, index=df.index)
        for col in colunas_texto:
            mascara |= df[col].astype(str).str.contains(texto, case=False, na=False)
        
        df_filtrado = df[mascara]
        removidas = len(df) - len(df_filtrado)
        
        if log_callback: log_callback(f"  🔍 Filtro texto '{texto}': {removidas} removidas")
        return df_filtrado

    def salvar_planilha(self, df, caminho_saida):
        if not caminho_saida.endswith('.xlsx'):
            caminho_saida += '.xlsx'
        df.to_excel(caminho_saida, index=False)
        return caminho_saida

    def gerar_resumo(self, caminho_entrada, aba, log_callback=None):
        """
        Generate statistical summary from Excel spreadsheet.
        
        Calculates:
        - Total unique orders
        - Total items quantity
        - Total sales value
        
        Args:
            caminho_entrada (str): Path to Excel file
            aba (str): Sheet name to process
            log_callback (callable, optional): Callback function for logging
            
        Returns:
            dict: Summary with keys 'total_itens', 'total_pedidos', 'valor_total', 'df'
                  or 'erro' key if processing fails
        """
        # Carregar
        df = self.carregar_planilha(caminho_entrada, aba, log_callback)
        
        resultado = {
            "total_itens": 0,
            "total_pedidos": 0,
            "valor_total": 0.0,
            "df": None # Para dashboard
        }
        
        try:
            # --- Configuracao de Colunas ---
            # Conforme especificacao do usuario:
            # Coluna B (indice 1) -> ID do Pedido (contar unicos)
            # Coluna Z (indice 25) -> Quantidade de Itens (somar)
            # Coluna AA (indice 26) -> Preco Unitario
            # Formula: Valor Total = SOMA(Z * AA) para cada linha
            
            COL_PEDIDOS_IDX = 1   # Coluna B
            COL_QTD_IDX = 25      # Coluna Z
            COL_PRECO_IDX = 26    # Coluna AA
            
            # Validacao de limites
            max_idx = max(COL_PEDIDOS_IDX, COL_QTD_IDX, COL_PRECO_IDX)
            if df.shape[1] <= max_idx:
                 msg = f"[WARNING] A planilha tem apenas {df.shape[1]} colunas, mas o resumo exige ate a coluna indice {max_idx} (AA)."
                 if log_callback: log_callback(msg)
                 # Retorna zerado mas com aviso, nao crasha
                 resultado["erro"] = msg
                 return resultado
            
            # 1. PEDIDOS UNICOS - Coluna B (contar apenas numeros diferentes)
            resultado["total_pedidos"] = df.iloc[:, COL_PEDIDOS_IDX].nunique()
            
            # 2. TOTAL DE ITENS - Soma da Coluna Z
            df['qty_clean'] = df.iloc[:, COL_QTD_IDX].apply(self.clean_numeric)
            resultado["total_itens"] = int(df['qty_clean'].sum())
            
            # 3. VALOR TOTAL - Formula: SOMA(Z * AA)
            # Limpar coluna AA (preco unitario)
            df['price_clean'] = df.iloc[:, COL_PRECO_IDX].apply(self.clean_numeric)
            
            # Multiplicar quantidade (Z) * preco unitario (AA) para cada linha
            df['valor_linha'] = df['qty_clean'] * df['price_clean']
            
            # Somar todos os valores
            resultado["valor_total"] = df['valor_linha'].sum()
            
            resultado["df"] = df
            
            if log_callback:
                log_callback(f"[OK] Resumo calculado: {resultado['total_pedidos']} pedidos, {resultado['total_itens']} itens, R$ {resultado['valor_total']:.2f}")
            
            return resultado
            
        except Exception as e:
            raise Exception(f"Erro ao calcular resumo: {e}")
