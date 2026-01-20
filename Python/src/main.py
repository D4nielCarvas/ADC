"""
Script para limpeza de planilha de itens mais vendidos por SKU
Versão com Interface Gráfica (GUI) usando tkinter
"""

import pandas as pd
import os
import json
from datetime import datetime
import shutil
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading


class LimpadorPlanilhaGUI:
    """
    MELHORIA GUI: Interface gráfica para facilitar o uso do script
    O usuário pode selecionar arquivos visualmente e acompanhar o progresso
    """
    
    def __init__(self, root):
        self.root = root
        self.root.title("ADC v2.0")
        self.root.geometry("800x850")
        self.root.configure(bg="#1e1e2e")  # Cor de fundo escura (Catppuccin Mocha)
        self.root.resizable(False, False)
        
        # Variáveis
        self.caminho_entrada = tk.StringVar()
        self.caminho_saida = tk.StringVar()
        self.processando = False
        
        # Novas Variáveis Multi-Relatório
        self.presets = self.carregar_presets()
        self.nome_preset = tk.StringVar()
        self.aba_selecionada = tk.StringVar()
        self.coluna_valor_selecionada = tk.StringVar()
        self.lista_abas = []
        self.lista_colunas = []
        
        # Variáveis para filtros adicionais (Padrão: Desmarcados)
        self.remover_duplicadas = tk.BooleanVar(value=False)
        self.remover_vazias = tk.BooleanVar(value=False)
        self.filtrar_por_valor = tk.BooleanVar(value=False)
        self.valor_minimo = tk.StringVar(value="")
        self.filtrar_por_texto = tk.BooleanVar(value=False)
        self.texto_filtro = tk.StringVar(value="")
        
        self.criar_interface()
        
    def carregar_presets(self):
        """Carregar presets do arquivo config.json"""
        caminho_config = os.path.join(os.path.dirname(__file__), "config.json")
        if os.path.exists(caminho_config):
            try:
                with open(caminho_config, 'r', encoding='utf-8') as f:
                    return json.load(f).get("presets", [])
            except Exception:
                return []
        return []
        
    def configurar_estilos(self):
        """Configurar o tema moderno do aplicativo"""
        style = ttk.Style()
        
        # Cores do Tema (Modern Dark)
        bg_dark = "#1e1e2e"
        bg_card = "#313244"
        accent = "#89b4fa"      # Azul pastel
        success = "#a6e3a1"     # Verde pastel
        warning = "#fab387"     # Laranja pastel
        text_color = "#cdd6f4"
        entry_bg = "#45475a"

        style.theme_use('clam')
        
        # Estilos Gerais
        style.configure("TFrame", background=bg_dark)
        style.configure("Card.TFrame", background=bg_card, relief="flat")
        
        style.configure("TLabel", 
            background=bg_dark, 
            foreground=text_color, 
            font=("Segoe UI", 10)
        )
        
        style.configure("Title.TLabel", 
            background=bg_dark, 
            foreground=accent, 
            font=("Segoe UI", 20, "bold")
        )
        
        style.configure("Header.TLabel", 
            background=bg_card, 
            foreground=accent, 
            font=("Segoe UI", 11, "bold")
        )
        
        style.configure("SubCard.TLabel", 
            background=bg_card, 
            foreground=text_color, 
            font=("Segoe UI", 10)
        )

        # Estilo de Botões Modernos
        style.configure("Accent.TButton", 
            padding=10, 
            font=("Segoe UI", 11, "bold"),
            background=accent,
            foreground="#11111b"
        )
        style.map("Accent.TButton",
            background=[('active', '#b4befe'), ('pressed', '#74c7ec')]
        )

        style.configure("Secondary.TButton", 
            padding=5, 
            font=("Segoe UI", 9),
            background=entry_bg,
            foreground=text_color
        )
        style.map("Secondary.TButton",
            background=[('active', '#585b70'), ('pressed', '#313244')]
        )

        # Estilo de Entries e Comboboxes
        style.configure("TEntry", 
            fieldbackground=entry_bg, 
            foreground=text_color,
            lightcolor=accent,
            bordercolor=bg_card
        )
        
        style.configure("TCombobox", 
            fieldbackground=entry_bg, 
            background=bg_card,
            foreground=text_color,
            arrowcolor=accent
        )

        style.configure("TLabelframe", 
            background=bg_card, 
            foreground=accent, 
            font=("Segoe UI", 10, "bold"),
            bordercolor=bg_dark
        )
        style.configure("TLabelframe.Label", background=bg_card, foreground=accent)
        
        style.configure("TCheckbutton", 
            background=bg_card, 
            foreground=text_color,
            font=("Segoe UI", 10)
        )
        style.map("TCheckbutton",
            background=[('active', bg_card)],
            foreground=[('active', accent)]
        )

    def criar_interface(self):
        """Criar interface com design premium"""
        self.configurar_estilos()
        
        # Frame principal com padding
        container = ttk.Frame(self.root, padding="30")
        container.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        
        # Header / Logo
        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 25))
        
        logo_label = ttk.Label(
            header_frame, 
            text="✨ ADC",
            style="Title.TLabel"
        )
        logo_label.pack(side=tk.LEFT)
        
        slogan_label = ttk.Label(
            header_frame, 
            text="CLEANER PRO",
            font=("Segoe UI", 8, "bold"),
            foreground="#585b70"
        )
        slogan_label.pack(side=tk.LEFT, padx=10, pady=(10, 0))

        # Botão Tela Cheia
        self.is_fullscreen = False
        btn_full = ttk.Button(
            header_frame, 
            text="🔲 Tela Cheia (F11)", 
            command=self.toggle_fullscreen,
            style="Secondary.TButton"
        )
        btn_full.pack(side=tk.RIGHT)
        self.root.bind("<F11>", lambda e: self.toggle_fullscreen())

        # --- SEÇÃO DE ARQUIVOS (CARD) ---
        file_card = ttk.Frame(container, style="Card.TFrame", padding=15)
        file_card.pack(fill=tk.X, pady=5)

        ttk.Label(file_card, text="📂 ARQUIVOS", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))

        # Entrada
        input_frame = ttk.Frame(file_card, style="Card.TFrame")
        input_frame.pack(fill=tk.X, pady=5)
        ttk.Label(input_frame, text="Entrada:", style="SubCard.TLabel").pack(side=tk.LEFT)
        self.entrada_entry = ttk.Entry(input_frame, textvariable=self.caminho_entrada, font=("Segoe UI", 9))
        self.entrada_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(input_frame, text="📁", width=3, command=self.selecionar_arquivo_entrada, style="Secondary.TButton").pack(side=tk.LEFT)

        # Saída
        output_frame = ttk.Frame(file_card, style="Card.TFrame")
        output_frame.pack(fill=tk.X, pady=5)
        ttk.Label(output_frame, text="Saída:  ", style="SubCard.TLabel").pack(side=tk.LEFT)
        self.saida_entry = ttk.Entry(output_frame, textvariable=self.caminho_saida, font=("Segoe UI", 9))
        self.saida_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(output_frame, text="💾", width=3, command=self.selecionar_arquivo_saida, style="Secondary.TButton").pack(side=tk.LEFT)

        # --- SEÇÃO DE CONFIGURAÇÃO (CARD) ---
        config_card = ttk.Frame(container, style="Card.TFrame", padding=15)
        config_card.pack(fill=tk.X, pady=15)

        ttk.Label(config_card, text="⚙️ CONFIGURAÇÕES", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))

        # Preset e Colunas
        row1 = ttk.Frame(config_card, style="Card.TFrame")
        row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(row1, text="Preset:", style="SubCard.TLabel").pack(side=tk.LEFT)
        self.preset_combo = ttk.Combobox(row1, textvariable=self.nome_preset, values=[p["nome"] for p in self.presets], state="readonly", width=20)
        self.preset_combo.pack(side=tk.LEFT, padx=10)
        self.preset_combo.bind("<<ComboboxSelected>>", self.aplicar_preset)

        ttk.Label(row1, text="Deletar Colunas:", style="SubCard.TLabel").pack(side=tk.LEFT, padx=(15, 0))
        self.indices_entry = ttk.Entry(row1, width=15)
        self.indices_entry.pack(side=tk.LEFT, padx=10)
        ttk.Label(row1, text="(A=1, B=2...)", foreground="#6c7086", font=("Segoe UI", 8), style="SubCard.TLabel").pack(side=tk.LEFT)

        # Aba e Coluna de Valor
        row2 = ttk.Frame(config_card, style="Card.TFrame")
        row2.pack(fill=tk.X, pady=10)
        
        ttk.Label(row2, text="Aba:", style="SubCard.TLabel").pack(side=tk.LEFT)
        self.aba_combo = ttk.Combobox(row2, textvariable=self.aba_selecionada, state="readonly", width=15)
        self.aba_combo.pack(side=tk.LEFT, padx=10)
        self.aba_combo.bind("<<ComboboxSelected>>", self.atualizar_colunas_aba)

        ttk.Label(row2, text="Coluna Valor:", style="SubCard.TLabel").pack(side=tk.LEFT, padx=(20, 0))
        self.coluna_combo = ttk.Combobox(row2, textvariable=self.coluna_valor_selecionada, state="readonly", width=20)
        self.coluna_combo.pack(side=tk.LEFT, padx=10)

        # --- SEÇÃO DE FILTROS (CARD) ---
        filters_card = ttk.LabelFrame(container, text=" ✨ REFINAMENTO ", padding=10, style="TLabelframe")
        filters_card.pack(fill=tk.X, pady=5)

        # Grid de filtros
        f_grid = ttk.Frame(filters_card, style="Card.TFrame")
        f_grid.pack(fill=tk.X)

        self.remover_duplicadas.set(True)
        ttk.Checkbutton(f_grid, text="🔄 Duplicadas", variable=self.remover_duplicadas).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        
        self.remover_vazias.set(True)
        ttk.Checkbutton(f_grid, text="🗑️ Vazias", variable=self.remover_vazias).grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        ttk.Checkbutton(f_grid, text="📊 Valor Mín:", variable=self.filtrar_por_valor).grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        ttk.Entry(f_grid, textvariable=self.valor_minimo, width=8).grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Checkbutton(f_grid, text="🔍 Texto:", variable=self.filtrar_por_texto).grid(row=1, column=2, sticky=tk.W, padx=(20, 0), pady=5)
        ttk.Entry(f_grid, textvariable=self.texto_filtro, width=15).grid(row=1, column=3, sticky=tk.W, pady=5)

        # Botão Processar
        self.btn_processar = ttk.Button(
            container, 
            text="🚀 INICIAR PROCESSAMENTO MÁGICO", 
            command=self.iniciar_processamento,
            style="Accent.TButton"
        )
        self.btn_processar.pack(fill=tk.X, pady=25)

        # Log
        log_frame = ttk.Frame(container)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(log_frame, text="REAL-TIME LOG", font=("Segoe UI", 8, "bold"), foreground="#585b70").pack(anchor=tk.W)
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            width=75, 
            height=10,
            font=("Consolas", 10),
            bg="#11111b",
            fg="#a6e3a1",  # Texto do log verde pastel
            borderwidth=0,
            padx=10,
            pady=10
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 10))

        # Status Bar
        self.status_label = tk.Label(
            self.root, 
            text="✨ Sistema Pronto",
            bg="#313244",
            fg="#cdd6f4",
            font=("Segoe UI", 9),
            anchor=tk.W,
            padx=10,
            pady=3
        )
        self.status_label.grid(row=1, column=0, sticky="ew")

        if self.presets:
            self.preset_combo.current(0)
            # Aplicar preset após a interface estar totalmente carregada
            self.root.after(100, self.aplicar_preset)

    def toggle_fullscreen(self):
        """Alternar entre modo tela cheia e janela"""
        self.is_fullscreen = not self.is_fullscreen
        self.root.attributes("-fullscreen", self.is_fullscreen)
        if not self.is_fullscreen:
            self.root.geometry("800x850")
        
    def aplicar_preset(self, event=None):
        """Aplicar as configurações do preset selecionado"""
        nome = self.nome_preset.get()
        
        # Primeiro limpa tudo para garantir
        self.indices_entry.delete(0, tk.END)
        self.valor_minimo.set("")
        self.texto_filtro.set("")
        self.remover_duplicadas.set(False)
        self.remover_vazias.set(False)
        self.filtrar_por_valor.set(False)
        self.filtrar_por_texto.set(False)

        if nome == "Personalizado":
            return

        for p in self.presets:
            if p["nome"] == nome:
                self.indices_entry.insert(0, p.get("colunas_deletar", ""))
                val_padrao = p.get("filtro_valor_padrao", "")
                if val_padrao and val_padrao != "0":
                    self.valor_minimo.set(val_padrao)
                    self.filtrar_por_valor.set(True)
                break

    def atualizar_colunas_aba(self, event=None):
        """Atualizar a lista de colunas baseada na aba selecionada"""
        arquivo = self.caminho_entrada.get()
        aba = self.aba_selecionada.get()
        
        if arquivo and aba:
            try:
                # Carrega apenas o cabeçalho para ser rápido
                df_temp = pd.read_excel(arquivo, sheet_name=aba, nrows=0)
                self.lista_colunas = list(df_temp.columns)
                self.coluna_combo['values'] = self.lista_colunas
                
                # Tenta pré-selecionar a última coluna numérica
                df_exemplo = pd.read_excel(arquivo, sheet_name=aba, nrows=5)
                cols_num = df_exemplo.select_dtypes(include=['number']).columns
                if len(cols_num) > 0:
                    self.coluna_valor_selecionada.set(cols_num[-1])
                elif self.lista_colunas:
                    self.coluna_valor_selecionada.set(self.lista_colunas[-1])
                
            except Exception as e:
                self.log(f"⚠ Erro ao ler colunas da aba: {e}")

    def selecionar_arquivo_entrada(self):
        """Abrir diálogo para selecionar arquivo e carregar abas"""
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha de entrada",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~/Downloads")
        )
        
        if arquivo:
            self.caminho_entrada.set(arquivo)
            
            # Carregar abas dinamicamente
            try:
                excel_file = pd.ExcelFile(arquivo)
                self.lista_abas = excel_file.sheet_names
                self.aba_combo['values'] = self.lista_abas
                if self.lista_abas:
                    self.aba_combo.current(0)
                    self.atualizar_colunas_aba()
                
                self.log(f"✓ Arquivo carregado. {len(self.lista_abas)} aba(s) encontrada(s).")
            except Exception as e:
                self.log(f"❌ Erro ao ler abas do arquivo: {e}")
                messagebox.showerror("Erro", f"Não foi possível ler as abas do arquivo:\n{e}")
    
    def selecionar_arquivo_saida(self):
        """Abrir diálogo para selecionar arquivo de saída"""
        arquivo = filedialog.asksaveasfilename(
            title="Salvar planilha limpa como",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.caminho_entrada.get()) if self.caminho_entrada.get() else os.path.expanduser("~/Downloads")
        )
        
        if arquivo:
            self.caminho_saida.set(arquivo)
            self.log(f"✓ Destino definido: {os.path.basename(arquivo)}")
    
    def log(self, mensagem):
        """Adicionar mensagem ao log visual de forma segura para threads"""
        def atualizar():
            self.log_text.insert(tk.END, mensagem + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, atualizar)
    
    def limpar_log(self):
        """Limpar área de log"""
        self.log_text.delete(1.0, tk.END)
    
    def iniciar_processamento(self):
        """Iniciar processamento em thread separada para não travar a GUI"""
        if self.processando:
            return
        
        # Validações básicas
        if not self.caminho_entrada.get():
            messagebox.showerror("Erro", "Selecione o arquivo de entrada!")
            return
        
        if not self.caminho_saida.get():
            messagebox.showerror("Erro", "Defina o arquivo de saída!")
            return
        
        # Executar em thread separada
        self.processando = True
        self.btn_processar.config(state="disabled", text="⏳ Processando...")
        self.status_label.config(text="⏳ Processando planilha...")
        
        thread = threading.Thread(target=self.processar_planilha)
        thread.daemon = True
        thread.start()
    
    def processar_planilha(self):
        """Processar a planilha (roda em thread separada)"""
        try:
            self.limpar_log()
            self.log("=" * 60)
            self.log("INICIANDO PROCESSAMENTO")
            self.log("=" * 60)
            self.log("")
            
            inicio = datetime.now()
            
            # Obter configurações
            caminho_entrada = self.caminho_entrada.get()
            caminho_saida = self.caminho_saida.get()
            
            # Processar índices (Convertendo de 1-based para 0-based)
            try:
                indices_texto = self.indices_entry.get()
                index_delet = [int(i.strip()) - 1 for i in indices_texto.split(",") if i.strip()]
                self.log(f"✓ Índices de colunas (Interno): {index_delet}")
            except ValueError:
                raise ValueError("Formato de índices inválido! Use números (Ex: 1, 2, 3).")
            
            # Passo 1: Validar arquivo
            self.log("\n1️⃣ Validando arquivo de entrada...")
            self.validar_arquivo_entrada(caminho_entrada)
            
            # Passo 2: Criar backup
            self.log("\n2️⃣ Criando backup do arquivo original...")
            self.criar_backup(caminho_entrada)
            
            # Passo 3: Carregar planilha
            self.log("\n3️⃣ Carregando planilha...")
            df = self.carregar_planilha(caminho_entrada)
            
            # Passo 4: Validar índices
            self.log("\n4️⃣ Validando índices das colunas...")
            self.validar_indices_colunas(df, index_delet)
            
            # Passo 5: Deletar colunas
            self.log("\n5️⃣ Deletando colunas desnecessárias...")
            df = self.deletar_colunas(df, index_delet)
            
            # Passo 6: Aplicar Filtros Adicionais
            self.log("\n6️⃣ Aplicando filtros adicionais...")
            df_limpo = self.aplicar_filtros_adicionais(df)
            
            # Passo 7: Salvar
            self.log("\n7️⃣ Salvando planilha limpa...")
            self.salvar_planilha(df_limpo, caminho_saida)
            
            # Finalizar
            tempo_execucao = (datetime.now() - inicio).total_seconds()
            
            self.log("")
            self.log("=" * 60)
            self.log("✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
            self.log(f"   Tempo de execução: {tempo_execucao:.2f} segundos")
            self.log("=" * 60)
            
            self.status_label.config(text="✅ Processamento concluído com sucesso!")
            
            messagebox.showinfo(
                "Sucesso!", 
                f"Planilha processada com sucesso!\n\nArquivo salvo em:\n{os.path.basename(caminho_saida)}"
            )
            
        except Exception as e:
            self.log("")
            self.log(f"❌ ERRO: {str(e)}")
            self.log("")
            self.status_label.config(text="❌ Erro no processamento")
            messagebox.showerror("Erro", f"Erro ao processar planilha:\n\n{str(e)}")
        
        finally:
            self.processando = False
            self.btn_processar.config(state="normal", text="▶️ PROCESSAR PLANILHA")
    
    # Funções de processamento (mesmas do código anterior, adaptadas para GUI)
    
    def criar_backup(self, caminho_arquivo):
        """Criar backup do arquivo original"""
        try:
            if os.path.exists(caminho_arquivo):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_base, extensao = os.path.splitext(caminho_arquivo)
                caminho_backup = f"{nome_base}_backup_{timestamp}{extensao}"
                
                shutil.copy2(caminho_arquivo, caminho_backup)
                self.log(f"✓ Backup criado: {os.path.basename(caminho_backup)}")
                return caminho_backup
        except Exception as e:
            self.log(f"⚠ Aviso: Não foi possível criar backup: {e}")
        return None
    
    def validar_arquivo_entrada(self, caminho):
        """Validar se o arquivo de entrada existe"""
        if not os.path.exists(caminho):
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
        
        if not os.path.isfile(caminho):
            raise ValueError(f"O caminho especificado não é um arquivo: {caminho}")
        
        self.log(f"✓ Arquivo de entrada encontrado: {os.path.basename(caminho)}")
        return True
    
    def carregar_planilha(self, caminho):
        """Carregar planilha do Excel com aba específica"""
        try:
            aba = self.aba_selecionada.get()
            self.log(f"   Aba selecionada: {aba}")
            df = pd.read_excel(caminho, sheet_name=aba)
            self.log(f"✓ Planilha carregada: {len(df)} linhas, {len(df.columns)} colunas")
            return df
        except Exception as e:
            raise Exception(f"Erro ao carregar planilha: {e}")
    
    def validar_indices_colunas(self, df, indices):
        """Validar se os índices das colunas são válidos"""
        total_colunas = len(df.columns)
        indices_invalidos = [i for i in indices if i >= total_colunas or i < 0]
        
        if indices_invalidos:
            raise ValueError(
                f"Índices de colunas inválidos: {indices_invalidos}. "
                f"A planilha tem apenas {total_colunas} colunas (índices 0-{total_colunas-1})"
            )
        
        self.log(f"✓ Índices de colunas validados: {indices}")
        return True
    
    def deletar_colunas(self, df, indices):
        """Deletar colunas especificadas (pelo índice atual no DataFrame)"""
        try:
            validador_indices = [i for i in indices if i < len(df.columns)]
            colunas_deletar = [df.columns[i] for i in validador_indices]
            self.log(f"✓ Deletando colunas: {colunas_deletar}")
            
            df_limpo = df.drop(df.columns[validador_indices], axis=1)
            
            self.log(f"✓ Colunas removidas com sucesso")
            self.log(f"  Colunas restantes: {len(df_limpo.columns)}")
            
            return df_limpo
        except Exception as e:
            raise Exception(f"Erro ao deletar colunas: {e}")
    
    def aplicar_filtros_adicionais(self, df):
        """Aplicar todos os filtros adicionais selecionados pelo usuário"""
        linhas_iniciais = len(df)
        
        # Filtro 1: Remover linhas duplicadas
        if self.remover_duplicadas.get():
            df = self.filtro_remover_duplicadas(df)
        
        # Filtro 2: Remover linhas vazias/incompletas
        if self.remover_vazias.get():
            df = self.filtro_remover_vazias(df)
        
        # Filtro 3: Filtrar por valor mínimo
        if self.filtrar_por_valor.get():
            df = self.filtro_por_valor_minimo(df)
        
        # Filtro 4: Filtrar por texto/SKU/Categoria
        if self.filtrar_por_texto.get():
            df = self.filtro_por_texto(df)
        
        linhas_finais = len(df)
        linhas_removidas = linhas_iniciais - linhas_finais
        
        self.log(f"✓ Filtros aplicados com sucesso")
        self.log(f"  Linhas removidas pelos filtros: {linhas_removidas}")
        self.log(f"  Linhas restantes: {linhas_finais}")
        
        return df
    
    def filtro_remover_duplicadas(self, df):
        """Remover linhas duplicadas"""
        linhas_antes = len(df)
        df_sem_dup = df.drop_duplicates()
        duplicadas_removidas = linhas_antes - len(df_sem_dup)
        
        if duplicadas_removidas > 0:
            self.log(f"  🔄 Removidas {duplicadas_removidas} linhas duplicadas")
        else:
            self.log(f"  🔄 Nenhuma linha duplicada encontrada")
        
        return df_sem_dup
    
    def filtro_remover_vazias(self, df):
        """Remover linhas vazias ou com dados insuficientes"""
        linhas_antes = len(df)
        
        # Remove linhas onde TODAS as colunas são nulas
        df_sem_vazias = df.dropna(how='all')
        
        # Remove linhas onde a MAIORIA das colunas são nulas (>50%)
        threshold = len(df.columns) // 2  # pelo menos 50% dos dados devem estar presentes
        df_sem_vazias = df_sem_vazias.dropna(thresh=threshold)
        
        vazias_removidas = linhas_antes - len(df_sem_vazias)
        
        if vazias_removidas > 0:
            self.log(f"  🗑️  Removidas {vazias_removidas} linhas vazias/incompletas")
        else:
            self.log(f"  🗑️  Nenhuma linha vazia encontrada")
        
        return df_sem_vazias
    
    def filtro_por_valor_minimo(self, df):
        """Filtrar por valor mínimo com normalização de dados"""
        try:
            valor_min = float(self.valor_minimo.get())
            coluna_filtro = self.coluna_valor_selecionada.get()
            
            if not coluna_filtro or coluna_filtro not in df.columns:
                self.log(f"  ⚠️ Coluna '{coluna_filtro}' não disponível - Filtro ignorado")
                return df
            
            linhas_antes = len(df)
            
            # NORMALIZAÇÃO: Converter coluna para numérico limpando strings de moeda/formatos BR
            def limpar_valor(x):
                if isinstance(x, str):
                    x = x.replace('R$', '').replace('.', '').replace(',', '.').strip()
                try:
                    return float(x)
                except:
                    return 0.0

            df_temp = df.copy()
            df_temp[coluna_filtro] = df_temp[coluna_filtro].apply(limpar_valor)
            
            df_filtrado = df[df_temp[coluna_filtro] >= valor_min]
            linhas_removidas = linhas_antes - len(df_filtrado)
            
            if linhas_removidas > 0:
                self.log(f"  📊 Removidas {linhas_removidas} linhas com {coluna_filtro} < {valor_min}")
            else:
                self.log(f"  📊 Nenhuma linha removida (todas >= {valor_min})")
            
            return df_filtrado
            
        except ValueError:
            self.log(f"  ⚠️ Valor mínimo inválido: '{self.valor_minimo.get()}' - Filtro ignorado")
            return df
    
    def filtro_por_texto(self, df):
        """Filtrar por texto/SKU/Categoria de forma otimizada"""
        texto = self.texto_filtro.get().strip()
        
        if not texto:
            self.log(f"  ⚠️ Texto de filtro vazio - Filtro ignorado")
            return df
        
        linhas_antes = len(df)
        
        # OTIMIZAÇÃO: Buscar apenas em colunas de texto (object) para evitar cast desnecessário
        colunas_texto = df.select_dtypes(include=['object']).columns
        
        if len(colunas_texto) == 0:
            self.log(f"  ⚠️ Nenhuma coluna de texto encontrada para filtrar")
            return df

        # Criar máscara booleana vetorizada
        mascara = pd.Series(False, index=df.index)
        for col in colunas_texto:
            mascara |= df[col].astype(str).str.contains(texto, case=False, na=False)
        
        df_filtrado = df[mascara]
        linhas_removidas = linhas_antes - len(df_filtrado)
        
        if linhas_removidas > 0:
            self.log(f"  🔍 Filtradas {len(df_filtrado)} linhas contendo '{texto}' ({linhas_removidas} removidas)")
        else:
            self.log(f"  🔍 Nenhuma linha contém '{texto}' - DataFrame vazio!")
        
        return df_filtrado
    
    def salvar_planilha(self, df, caminho_saida):
        """Salvar planilha processada"""
        try:
            if df.empty:
                raise ValueError("DataFrame vazio. Nada para salvar.")
            
            df.to_excel(caminho_saida, index=False)
            
            if os.path.exists(caminho_saida):
                tamanho = os.path.getsize(caminho_saida) / 1024  # KB
                self.log(f"✓ Arquivo salvo com sucesso: {os.path.basename(caminho_saida)}")
                self.log(f"  Tamanho: {tamanho:.2f} KB")
        
        except Exception as e:
            raise Exception(f"Erro ao salvar planilha: {e}")


def main():
    """Função principal para iniciar a aplicação GUI"""
    root = tk.Tk()
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    # Criar aplicação
    app = LimpadorPlanilhaGUI(root)
    
    # Iniciar loop da GUI
    root.mainloop()


if __name__ == "__main__":
    main()
