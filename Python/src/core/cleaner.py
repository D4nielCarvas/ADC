import pandas as pd
import os
import json
from datetime import datetime

class ADCLogic:
    def __init__(self):
        self.presets = self.carregar_presets()
        self.cache_excel = {}

    def carregar_presets(self):
        """Carregar presets do arquivo config/settings.json"""
        # Estratégia de busca de arquivo robusta
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
        """Salva a lista atual de presets no settings.json"""
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
        """Carrega a planilha. Se aba for None, retorna objeto ExcelFile ou primeira aba."""
        try:
            if log_callback and aba: log_callback(f"Tentando carregar aba: {aba}")
            
            if aba is None:
                # Retorna apenas o objeto para listar abas ou carregar default depois
                return pd.ExcelFile(caminho, engine='openpyxl' if not caminho.lower().endswith('.xls') else 'xlrd')

            try:
                if caminho.lower().endswith('.xls'):
                    df = pd.read_excel(caminho, sheet_name=aba, engine='xlrd')
                else:
                    df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
            except Exception:
                # Retry logic
                engine_alt = 'openpyxl' if caminho.lower().endswith('.xls') else 'xlrd'
                if log_callback: log_callback(f"⚠️ Motor primário falhou. Tentando {engine_alt}...")
                df = pd.read_excel(caminho, sheet_name=aba, engine=engine_alt)
            
            return df
        except Exception as e:
            raise Exception(f"Erro ao carregar planilha: {e}")

    def listar_abas(self, caminho, log_callback=None):
        try:
            xl = self.carregar_planilha(caminho) # Retorna ExcelFile neste caso se implementado ou reusa logica
            # Na implementação original do pandas read_excel pode ler tudo, mas ExcelFile é melhor pra abas
            if isinstance(xl, pd.ExcelFile):
                 return xl.sheet_names
            # Se a func carregar_planilha retornar df direto qdo aba é None, ajustamos:
            # Aqui vamos instanciar direto para garantir
            engine = 'xlrd' if caminho.lower().endswith('.xls') else 'openpyxl'
            xl = pd.ExcelFile(caminho, engine=engine)
            return xl.sheet_names
        except Exception as e:
             if log_callback: log_callback(f"Erro ao listar abas: {e}")
             raise e

    def validar_arquivo_entrada(self, caminho, log_callback=None):
        if not os.path.exists(caminho):
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
        if not os.path.isfile(caminho):
            raise ValueError(f"O caminho especificado não é um arquivo: {caminho}")
        if log_callback: log_callback(f"✓ Arquivo de entrada encontrado: {os.path.basename(caminho)}")
        return True

    def validar_indices_colunas(self, df, indices, log_callback=None):
        total_colunas = len(df.columns)
        indices_invalidos = [i for i in indices if i >= total_colunas or i < 0]
        
        if indices_invalidos:
            msg = f"Índices inválidos: {indices_invalidos}. A planilha tem {total_colunas} colunas (0-{total_colunas-1})"
            if log_callback: log_callback(f"❌ {msg}")
            raise ValueError(msg)
        
        if log_callback: log_callback(f"✓ Índices validados: {indices}")
        return True

    def processar_limpeza(self, caminho_entrada, aba, indices_deletar, opcoes_filtros=None, log_callback=None):
        """
        opcoes_filtros: dict com chaves:
          - remover_duplicadas (bool)
          - remover_vazias (bool)
          - filtro_valor (dict): { 'ativo': bool, 'minimo': float, 'coluna': str }
          - filtro_texto (dict): { 'ativo': bool, 'texto': str }
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
        if log_callback: log_callback(f"✓ Deletando colunas: {colunas_deletar}")
        
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
        try:
            if not coluna_filtro or coluna_filtro not in df.columns:
                if log_callback: log_callback(f"  ⚠️ Coluna '{coluna_filtro}' não disponível para filtro de valor")
                return df
            
            def limpar_valor(x):
                if isinstance(x, str):
                    x = x.replace('R$', '').replace('.', '').replace(',', '.').strip()
                try: return float(x)
                except: return 0.0

            df_temp = df.copy()
            df_temp[coluna_filtro] = df_temp[coluna_filtro].apply(limpar_valor)
            
            df_filtrado = df[df_temp[coluna_filtro] >= valor_min]
            removidas = len(df) - len(df_filtrado)
            
            if log_callback and removidas > 0:
                log_callback(f"  📊 Removidas {removidas} linhas com {coluna_filtro} < {valor_min}")
            
            return df_filtrado
        except Exception as e:
            if log_callback: log_callback(f"  ⚠️ Erro no filtro de valor: {e}")
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
        # Carregar
        df = self.carregar_planilha(caminho_entrada, aba, log_callback)
        
        resultado = {
            "total_itens": 0,
            "total_pedidos": 0,
            "valor_total": 0.0,
            "df": None # Para dashboard
        }
        
        try:
            # Lógica robusta de resumo (Baseada no main.py)
            col_pedidos = 1  # B (index 1)
            col_quantidade = 25 # Z (index 25)
            col_preco = 26 # AA (index 26)
            
            # Validação Mínima
            # Ajuste: se a planilha for pequena, não quebra, mas avisa
            if df.shape[1] <= col_preco:
                 if log_callback: log_callback("⚠️ Planilha menor que o esperado para resumo completo (Z/AA).")
            
            def clean_numeric(val):
                if pd.isna(val): return 0.0
                if isinstance(val, (int, float)): return float(val)
                s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
                try: return float(s)
                except: return 0.0

            # Pedidos Únicos
            if df.shape[1] > col_pedidos:
                resultado["total_pedidos"] = df.iloc[:, col_pedidos].nunique()
            
            # Totais
            if df.shape[1] > col_quantidade:
                df['qty_clean'] = df.iloc[:, col_quantidade].apply(clean_numeric)
                resultado["total_itens"] = int(df['qty_clean'].sum())
            
            if df.shape[1] > col_preco:
                df['price_clean'] = df.iloc[:, col_preco].apply(clean_numeric)
                # Ensure qty_clean exists
                if 'qty_clean' not in df.columns: df['qty_clean'] = 0
                df['total_row'] = df['qty_clean'] * df['price_clean']
                resultado["valor_total"] = df['total_row'].sum()
            
            resultado["df"] = df
            return resultado
            
        except Exception as e:
            raise Exception(f"Erro ao calcular resumo: {e}")
