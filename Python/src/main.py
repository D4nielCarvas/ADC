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
import xlrd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class LimpadorPlanilhaGUI:
    """
    MELHORIA GUI: Interface gráfica para facilitar o uso do script
    O usuário pode selecionar arquivos visualmente e acompanhar o progresso
    """
    
    def __init__(self, root):
        self.root = root
        self.root.title("ADC v2.5 Pro")
        self.root.geometry("1000x850") # Aumentado para acomodar a sidebar
        self.root.configure(bg="#1e1e2e")
        
        # Variáveis de Estado
        self.caminho_entrada = tk.StringVar()
        self.caminho_saida = tk.StringVar()
        self.processando = False
        self.pagina_atual = None
        self.cache_excel = {}  # OTIMIZAÇÃO: Cache de objetos pd.ExcelFile
        self.dashboard_canvas = None # Guardar referência do gráfico
        self.is_fullscreen = False # INICIALIZAÇÃO CORRIGIDA
        self.df_resultado = None  # Armazena resultado processado
        
        # Novas Variáveis Multi-Relatório
        self.presets = self.carregar_presets()
        self.nome_preset = tk.StringVar()
        self.aba_selecionada = tk.StringVar()
        self.coluna_valor_selecionada = tk.StringVar()
        self.lista_abas = []
        self.lista_colunas = []
        
        # Variáveis para filtros adicionais
        self.remover_duplicadas = tk.BooleanVar(value=True)
        self.remover_vazias = tk.BooleanVar(value=True)
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
        accent = "#89b4fa"      
        success = "#a6e3a1"     
        warning = "#fab387"     
        text_color = "#cdd6f4"
        entry_bg = "#45475a"
        sidebar_bg = "#181825"

        style.theme_use('clam')
        
        # Estilos Gerais
        style.configure("TFrame", background=bg_dark)
        style.configure("Sidebar.TFrame", background=sidebar_bg)
        style.configure("Card.TFrame", background=bg_card, relief="flat")
        
        style.configure("TLabel", 
            background=bg_dark, 
            foreground=text_color, 
            font=("Segoe UI", 10)
        )
        
        style.configure("Title.TLabel", 
            background=sidebar_bg, 
            foreground=accent, 
            font=("Segoe UI", 20, "bold")
        )
        
        style.configure("Header.TLabel", 
            background=bg_card, 
            foreground=accent, 
            font=("Segoe UI", 11, "bold")
        )

        style.configure("Nav.TButton", 
            padding=10, 
            font=("Segoe UI", 11),
            background=sidebar_bg,
            foreground=text_color,
            anchor="w"
        )
        style.map("Nav.TButton",
            background=[('active', '#313244'), ('pressed', accent)],
            foreground=[('active', accent)]
        )

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
        
        style.configure("TEntry", fieldbackground=entry_bg, foreground=text_color)
        style.configure("TCombobox", fieldbackground=entry_bg, background=bg_card, foreground=text_color)
        style.configure("TLabelframe", background=bg_card, foreground=accent, font=("Segoe UI", 10, "bold"))
        style.configure("TLabelframe.Label", background=bg_card, foreground=accent)
        style.configure("TCheckbutton", background=bg_card, foreground=text_color)

        style.configure("Stat.TLabel", 
            background=bg_card, 
            foreground=success, 
            font=("Segoe UI", 16, "bold")
        )
        style.configure("StatDesc.TLabel", 
            background=bg_card, 
            foreground=text_color, 
            font=("Segoe UI", 9)
        )

    def criar_interface(self):
        """Criar interface com sidebar e páginas"""
        self.configurar_estilos()
        
        # Layout Principal
        self.sidebar = ttk.Frame(self.root, style="Sidebar.TFrame", width=200)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        self.main_content = ttk.Frame(self.root, padding="30")
        self.main_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- SIDEBAR CONTENT ---
        ttk.Label(self.sidebar, text="✨ ADC", style="Title.TLabel").pack(pady=(30, 10), padx=20, anchor="w")
        ttk.Label(self.sidebar, text="PRO VERSION", font=("Segoe UI", 8, "bold"), foreground="#585b70", background="#181825").pack(padx=20, anchor="w", pady=(0, 40))

        self.btn_nav_limpeza = ttk.Button(self.sidebar, text="🧹 Limpeza", style="Nav.TButton", command=lambda: self.mudar_pagina("limpeza"))
        self.btn_nav_limpeza.pack(fill=tk.X, padx=10, pady=2)

        self.btn_nav_resumo = ttk.Button(self.sidebar, text="📊 Resumo", style="Nav.TButton", command=lambda: self.mudar_pagina("resumo"))
        self.btn_nav_resumo.pack(fill=tk.X, padx=10, pady=2)

        self.btn_nav_config = ttk.Button(self.sidebar, text="⚙️ Configurações", style="Nav.TButton", command=lambda: self.mudar_pagina("config"))
        self.btn_nav_config.pack(fill=tk.X, padx=10, pady=2)

        # Espaçador inferior
        ttk.Frame(self.sidebar, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)

        self.btn_full = ttk.Button(self.sidebar, text="🔲 Tela Cheia", command=self.toggle_fullscreen, style="Secondary.TButton")
        self.btn_full.pack(fill=tk.X, padx=10, pady=5)
        self.root.bind("<F11>", lambda e: self.toggle_fullscreen())

        # --- CONTAINERS DE PÁGINAS ---
        self.container_limpeza = ttk.Frame(self.main_content)
        self.container_resumo = ttk.Frame(self.main_content)
        self.container_config = ttk.Frame(self.main_content)

        self.montar_pagina_limpeza()
        self.montar_pagina_resumo()
        self.montar_pagina_config()

        # Barra de Progresso Global
        self.progress_frame = ttk.Frame(self.main_content)
        self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, mode='determinate', length=100)
        self.progress_bar.pack(fill=tk.X)
        self.progress_bar.pack_forget() # Oculta inicialmente

        # Log e Status
        self.status_label = tk.Label(self.root, text="✨ Sistema Pronto", bg="#313244", fg="#cdd6f4", font=("Segoe UI", 9), anchor=tk.W, padx=10, pady=3)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.mudar_pagina("limpeza")

    def mudar_pagina(self, pagina):
        """Alternar entre os containers de página com animação suave"""
        if self.pagina_atual == pagina:
            return

        def efeito_transicao():
            # Limpa todos os containers de forma rápida
            self.container_limpeza.pack_forget()
            self.container_resumo.pack_forget()
            self.container_config.pack_forget()
            
            # Reset visual dos botões da sidebar (opcional: destacar o ativo)
            
            if pagina == "limpeza":
                self.container_limpeza.pack(fill=tk.BOTH, expand=True)
                self.status_label.config(text="✨ Modo Limpeza Ativo")
            elif pagina == "resumo":
                self.container_resumo.pack(fill=tk.BOTH, expand=True)
                self.status_label.config(text="✨ Modo Resumo Ativo")
                # Forçar o preset de resumo se existir
                for p in self.presets:
                    if p.get("tipo") == "resumo":
                        self.nome_preset.set(p["nome"])
                        self.aplicar_preset()
                        break
            elif pagina == "config":
                self.container_config.pack(fill=tk.BOTH, expand=True)
                self.status_label.config(text="⚙️ Configurações de Sistema")

            self.pagina_atual = pagina
            # Efeito de Fade-in (Simulado via update)
            self.root.update_idletasks()

        # Pequeno delay para suavizar
        self.root.after(50, efeito_transicao)

    def montar_pagina_header(self, parent, titulo, icone):
        header = ttk.Frame(parent)
        header.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(header, text=f"{icone} {titulo}", font=("Segoe UI", 18, "bold"), foreground="#89b4fa").pack(side=tk.LEFT)

    def montar_pagina_limpeza(self):
        """Constrói os widgets da página de limpeza"""
        self.montar_pagina_header(self.container_limpeza, "Limpeza de Planilhas", "🧹")
        
        # Card Arquivos
        file_card = ttk.Frame(self.container_limpeza, style="Card.TFrame", padding=15)
        file_card.pack(fill=tk.X, pady=5)
        ttk.Label(file_card, text="📂 ARQUIVOS", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))

        # Entrada
        input_frame = ttk.Frame(file_card, style="Card.TFrame")
        input_frame.pack(fill=tk.X, pady=5)
        tk.Label(input_frame, text="Entrada:", background="#313244", foreground="#cdd6f4").pack(side=tk.LEFT)
        ttk.Entry(input_frame, textvariable=self.caminho_entrada, font=("Segoe UI", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(input_frame, text="📁", width=3, command=self.selecionar_arquivo_entrada, style="Secondary.TButton").pack(side=tk.LEFT)

        # Saída (REMOVIDO DO INICIO - SERÁ MOSTRADO NO FINAL)
        self.output_frame = ttk.Frame(file_card, style="Card.TFrame")
        # self.output_frame.pack(fill=tk.X, pady=5) # Oculto inicialmente
        tk.Label(self.output_frame, text="Salvar em:", background="#313244", foreground="#cdd6f4").pack(side=tk.LEFT)
        ttk.Entry(self.output_frame, textvariable=self.caminho_saida, font=("Segoe UI", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(self.output_frame, text="💾", width=3, command=self.selecionar_arquivo_saida, style="Secondary.TButton").pack(side=tk.LEFT)

        # Configurações
        config_card = ttk.Frame(self.container_limpeza, style="Card.TFrame", padding=15)
        config_card.pack(fill=tk.X, pady=15)
        ttk.Label(config_card, text="⚙️ CONFIGURAÇÕES", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))

        row1 = ttk.Frame(config_card, style="Card.TFrame")
        row1.pack(fill=tk.X, pady=5)
        tk.Label(row1, text="Preset:", background="#313244", foreground="#cdd6f4").pack(side=tk.LEFT)
        self.preset_combo = ttk.Combobox(row1, textvariable=self.nome_preset, values=[p["nome"] for p in self.presets], state="readonly", width=25)
        self.preset_combo.pack(side=tk.LEFT, padx=10)
        self.preset_combo.bind("<<ComboboxSelected>>", self.aplicar_preset)

        tk.Label(row1, text="Aba:", background="#313244", foreground="#cdd6f4").pack(side=tk.LEFT, padx=15)
        self.aba_combo = ttk.Combobox(row1, textvariable=self.aba_selecionada, state="readonly", width=15)
        self.aba_combo.pack(side=tk.LEFT, padx=10)
        self.aba_combo.bind("<<ComboboxSelected>>", self.atualizar_colunas_aba)

        row2 = ttk.Frame(config_card, style="Card.TFrame")
        row2.pack(fill=tk.X, pady=10)
        tk.Label(row2, text="Deletar Colunas:", background="#313244", foreground="#cdd6f4").pack(side=tk.LEFT)
        self.indices_entry = ttk.Entry(row2, width=30)
        self.indices_entry.pack(side=tk.LEFT, padx=10)

        # Refinamento
        filters_card = ttk.LabelFrame(self.container_limpeza, text=" ✨ REFINAMENTO ", padding=10, style="TLabelframe")
        filters_card.pack(fill=tk.X, pady=5)
        f_grid = ttk.Frame(filters_card, style="Card.TFrame")
        f_grid.pack(fill=tk.X)
        ttk.Checkbutton(f_grid, text="🔄 Duplicadas", variable=self.remover_duplicadas).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        ttk.Checkbutton(f_grid, text="🗑️ Vazias", variable=self.remover_vazias).grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        self.btn_processar_limpeza = ttk.Button(self.container_limpeza, text="🚀 INICIAR LIMPEZA", command=self.iniciar_processamento, style="Accent.TButton")
        self.btn_processar_limpeza.pack(fill=tk.X, pady=10)

        # Botão de Salvamento (Aparece após processar)
        self.btn_salvar_limpeza = ttk.Button(self.container_limpeza, text="💾 SALVAR PLANILHA LIMPA", command=self.salvar_resultado_processado, style="Secondary.TButton")
        # self.btn_salvar_limpeza.pack_forget() # Oculto

        # Log com tamanho fixo para não sumir (AGORA ACIMA DO DASHBOARD)
        self.montar_log(self.container_limpeza)

        # Container para os gráficos do Dashboard na Limpeza
        self.dash_container_limpeza = ttk.Frame(self.container_limpeza, style="Card.TFrame")
        self.dash_container_limpeza.pack(fill=tk.BOTH, expand=True, pady=5)

    def montar_pagina_resumo(self):
        """Constrói os widgets da página de resumo"""
        self.montar_pagina_header(self.container_resumo, "Resumo de Pedidos", "📊")

        file_card = ttk.Frame(self.container_resumo, style="Card.TFrame", padding=15)
        file_card.pack(fill=tk.X, pady=5)
        ttk.Label(file_card, text="� SELECIONAR PLANILHA", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))

        input_frame = ttk.Frame(file_card, style="Card.TFrame")
        input_frame.pack(fill=tk.X, pady=5)
        tk.Label(input_frame, text="Arquivo:", background="#313244", foreground="#cdd6f4").pack(side=tk.LEFT)
        ttk.Entry(input_frame, textvariable=self.caminho_entrada, font=("Segoe UI", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(input_frame, text="�", width=3, command=self.selecionar_arquivo_entrada, style="Secondary.TButton").pack(side=tk.LEFT)

        config_card = ttk.Frame(self.container_resumo, style="Card.TFrame", padding=15)
        config_card.pack(fill=tk.X, pady=15)
        
        r1 = ttk.Frame(config_card, style="Card.TFrame")
        r1.pack(fill=tk.X, pady=5)
        tk.Label(r1, text="Selecione a Aba:", background="#313244", foreground="#cdd6f4").pack(side=tk.LEFT)
        self.aba_resumo_combo = ttk.Combobox(r1, textvariable=self.aba_selecionada, state="readonly", width=30)
        self.aba_resumo_combo.pack(side=tk.LEFT, padx=10)

        info_box = ttk.Frame(self.container_resumo, style="Card.TFrame", padding=15)
        info_box.pack(fill=tk.X, pady=10)
        tk.Label(info_box, text="💡 Esta função calcula automaticamente:", background="#313244", foreground="#89b4fa", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W)
        tk.Label(info_box, text="• Total de Itens (Coluna Z)\n• Total de Pedidos (Coluna B)\n• Valor Total (Z * AA)", background="#313244", foreground="#cdd6f4", justify=tk.LEFT).pack(anchor=tk.W, pady=5)

        # Cards de Estatísticas
        stats_frame = ttk.Frame(self.container_resumo, style="TFrame")
        stats_frame.pack(fill=tk.X, pady=10)
        
        # Card Itens
        self.card_itens = ttk.Frame(stats_frame, style="Card.TFrame", padding=15)
        self.card_itens.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        ttk.Label(self.card_itens, text="ITENS TOTAL", style="StatDesc.TLabel").pack()
        self.lbl_stat_itens = ttk.Label(self.card_itens, text="0", style="Stat.TLabel")
        self.lbl_stat_itens.pack()

        # Card Pedidos
        self.card_pedidos = ttk.Frame(stats_frame, style="Card.TFrame", padding=15)
        self.card_pedidos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        ttk.Label(self.card_pedidos, text="PEDIDOS", style="StatDesc.TLabel").pack()
        self.lbl_stat_pedidos = ttk.Label(self.card_pedidos, text="0", style="Stat.TLabel")
        self.lbl_stat_pedidos.pack()

        # Card Valor
        self.card_valor = ttk.Frame(stats_frame, style="Card.TFrame", padding=15)
        self.card_valor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        ttk.Label(self.card_valor, text="VALOR TOTAL", style="StatDesc.TLabel").pack()
        self.lbl_stat_valor = ttk.Label(self.card_valor, text="R$ 0,00", style="Stat.TLabel")
        self.lbl_stat_valor.pack()

        self.btn_processar_resumo = ttk.Button(self.container_resumo, text="📊 GERAR RESUMO AGORA", command=self.iniciar_processamento, style="Accent.TButton")
        self.btn_processar_resumo.pack(fill=tk.X, pady=20)

        # Container para os gráficos do Dashboard
        self.dash_container = ttk.Frame(self.container_resumo, style="Card.TFrame")
        self.dash_container.pack(fill=tk.BOTH, expand=True)

        self.montar_log(self.container_resumo)

    def montar_pagina_config(self):
        """Página de Configurações (Editor de Presets)"""
        self.montar_pagina_header(self.container_config, "Configurações e Presets", "⚙️")

        main_card = ttk.Frame(self.container_config, style="Card.TFrame", padding=20)
        main_card.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_card, text="📝 GERENCIAR PRESETS", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 15))
        
        # Lista de Presets existentes
        list_frame = ttk.Frame(main_card, style="Card.TFrame")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.preset_listbox = tk.Listbox(list_frame, bg="#1e1e2e", fg="#cdd6f4", font=("Segoe UI", 10), borderwidth=0, highlightthickness=1, highlightcolor="#89b4fa")
        self.preset_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.preset_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preset_listbox.config(yscrollcommand=scrollbar.set)
        
        # Botões de Ação
        btn_frame = ttk.Frame(main_card, style="Card.TFrame")
        btn_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(btn_frame, text="➕ Novo", style="Secondary.TButton", width=10, command=self.adicionar_preset_ui).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ Excluir", style="Secondary.TButton", width=10, command=self.excluir_preset_ui).pack(side=tk.LEFT, padx=5)
        
        # Info
        inf = tk.Label(main_card, text="💡 As alterações são salvas automaticamente no arquivo config.json.", background="#313244", foreground="#585b70", font=("Segoe UI", 9, "italic"))
        inf.pack(side=tk.BOTTOM, pady=10)
        
        self.atualizar_lista_presets_config()

    def adicionar_preset_ui(self):
        """Abre prompt para criar novo preset básico"""
        from tkinter import simpledialog
        nome = simpledialog.askstring("Novo Preset", "Nome do Preset:")
        if nome:
            novo = {"nome": nome, "deletar": "", "tipo": "limpeza"}
            self.presets.append(novo)
            self.salvar_presets()
            self.atualizar_lista_presets_config()
            self.preset_combo['values'] = [p["nome"] for p in self.presets]
            messagebox.showinfo("Sucesso", f"Preset '{nome}' criado!")

    def excluir_preset_ui(self):
        """Exclui o preset selecionado na lista"""
        selecao = self.preset_listbox.curselection()
        if not selecao:
            messagebox.showwarning("Aviso", "Selecione um preset para excluir.")
            return
        
        index = selecao[0]
        nome = self.presets[index]["nome"]
        
        if messagebox.askyesno("Confirmar", f"Deseja excluir o preset '{nome}'?"):
            del self.presets[index]
            self.salvar_presets()
            self.atualizar_lista_presets_config()
            self.preset_combo['values'] = [p["nome"] for p in self.presets]

    def salvar_presets(self):
        """Salva a lista atual de presets no config.json"""
        try:
            caminho = os.path.join(os.path.dirname(__file__), "config.json")
            with open(caminho, 'w', encoding='utf-8') as f:
                json.dump({"presets": self.presets}, f, indent=4, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar presets: {e}")

    def atualizar_lista_presets_config(self):
        """Atualiza a listagem visual de presets na página de config"""
        self.preset_listbox.delete(0, tk.END)
        for p in self.presets:
            self.preset_listbox.insert(tk.END, f" {p['nome']} ({p.get('tipo', 'limpeza')})")

    def montar_log(self, parent):
        log_frame = ttk.Frame(parent)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=5, font=("Consolas", 10), bg="#11111b", fg="#a6e3a1", borderwidth=0, padx=10, pady=10)
        self.log_text.pack(fill=tk.BOTH, expand=False, pady=5)

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
        """Atualizar a lista de colunas baseada na aba selecionada com Fallback"""
        arquivo = self.caminho_entrada.get()
        aba = self.aba_selecionada.get()
        
        if not arquivo or not aba: return

        try:
            # Tenta carregar apenas o cabeçalho para ser rápido
            try:
                if arquivo.lower().endswith('.xls'):
                    df_temp = pd.read_excel(arquivo, sheet_name=aba, nrows=0, engine='xlrd')
                    df_exemplo = pd.read_excel(arquivo, sheet_name=aba, nrows=5, engine='xlrd')
                else:
                    df_temp = pd.read_excel(arquivo, sheet_name=aba, nrows=0, engine='openpyxl')
                    df_exemplo = pd.read_excel(arquivo, sheet_name=aba, nrows=5, engine='openpyxl')
            except Exception:
                # FALLBACK: Tenta o motor oposto
                engine_alt = 'openpyxl' if arquivo.lower().endswith('.xls') else 'xlrd'
                self.log(f"⚠️ Tentando motor alternativo ({engine_alt}) para colunas...")
                df_temp = pd.read_excel(arquivo, sheet_name=aba, nrows=0, engine=engine_alt)
                df_exemplo = pd.read_excel(arquivo, sheet_name=aba, nrows=5, engine=engine_alt)
            
            self.lista_colunas = list(df_temp.columns)
            
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
            
            # Carregar abas dinamicamente com OTIMIZAÇÃO (Cache)
            try:
                if arquivo in self.cache_excel:
                    self.log(f"⚡ Usando cache para: {os.path.basename(arquivo)}")
                    excel_file = self.cache_excel[arquivo]
                else:
                    self.set_progress(20, "Lendo estrutura do arquivo...")
                    try:
                        if arquivo.lower().endswith('.xls'):
                            excel_file = pd.ExcelFile(arquivo, engine='xlrd', engine_kwargs={'ignore_workbook_corruption': True})
                        else:
                            excel_file = pd.ExcelFile(arquivo, engine='openpyxl')
                    except Exception as e:
                        # FALLBACK AUTOMÁTICO
                        engine_alt = 'openpyxl' if arquivo.lower().endswith('.xls') else 'xlrd'
                        self.log(f"⚠️ Motor primário falhou. Tentando {engine_alt}...")
                        excel_file = pd.ExcelFile(arquivo, engine=engine_alt)
                    
                    self.cache_excel[arquivo] = excel_file
                
                self.lista_abas = excel_file.sheet_names
                
                # Atualiza ambos os comboboxes (Limpeza e Resumo)
                self.aba_combo['values'] = self.lista_abas
                self.aba_resumo_combo['values'] = self.lista_abas
                
                if self.lista_abas:
                    self.aba_selecionada.set(self.lista_abas[0])
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
    
    def set_progress(self, value, text=None):
        """Atualiza a barra de progresso de forma segura"""
        def atualizar():
            if value >= 100 or value < 0:
                self.progress_bar.pack_forget()
            else:
                if not self.progress_bar.winfo_viewable():
                    self.progress_bar.pack(fill=tk.X, pady=(5, 0))
                self.progress_bar['value'] = value
            
            if text:
                self.status_label.config(text=f"⏳ {text}")
        self.root.after(0, atualizar)

    def log(self, mensagem):
        """Adicionar mensagem ao log visual de forma segura para threads"""
        def atualizar():
            if hasattr(self, 'log_text'):
                self.log_text.insert(tk.END, mensagem + "\n")
                self.log_text.see(tk.END)
        self.root.after(0, atualizar)
    
    def limpar_log(self):
        """Limpar área de log"""
        if hasattr(self, 'log_text'):
            self.log_text.delete(1.0, tk.END)
    
    def iniciar_processamento(self):
        """Iniciar processamento em thread separada para não travar a GUI"""
        if self.processando:
            return
        
        # Validações básicas
        if not self.caminho_entrada.get():
            messagebox.showerror("Erro", "Selecione o arquivo de entrada!")
            return
        
        # O caminho de saída agora é opcional no início (fluxo post-save)
        # if not self.caminho_saida.get() and self.get_preset_tipo() != "resumo":
        #     messagebox.showerror("Erro", "Defina o arquivo de saída!")
        #     return
        
        # Executar em thread separada
        self.processando = True
        
        # Desabilitar botão correto
        btn = self.btn_processar_limpeza if self.pagina_atual == "limpeza" else self.btn_processar_resumo
        btn.config(state="disabled", text="⏳ Processando...")
        
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
            
            preset_atual = self.get_preset_atual()
            is_resumo = preset_atual.get("tipo") == "resumo" if preset_atual else False
            
            inicio = datetime.now()
            
            # Obter configurações
            caminho_entrada = self.caminho_entrada.get()
            caminho_saida = self.caminho_saida.get()
            
            if is_resumo:
                self.log("📋 MODO: RESUMO DE PEDIDOS")
                df = self.carregar_planilha(caminho_entrada)
                self.processar_resumo_pedidos(df)
                
                tempo_execucao = (datetime.now() - inicio).total_seconds()
                self.log("")
                self.log("=" * 60)
                self.log("✅ RESUMO CONCLUÍDO!")
                self.log(f"   Tempo de execução: {tempo_execucao:.2f} segundos")
                self.log("=" * 60)
                self.status_label.config(text="✅ Resumo concluído!")
                return
            
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
            self.set_progress(10, "Lendo planilha...")
            df = self.carregar_planilha(caminho_entrada)
            
            # Passo 4: Validar índices
            self.set_progress(30, "Validando colunas...")
            self.validar_indices_colunas(df, index_delet)
            
            # Passo 5: Deletar colunas
            self.set_progress(50, "Limpando dados...")
            df = self.deletar_colunas(df, index_delet)
            
            # Passo 6: Aplicar Filtros Adicionais
            self.set_progress(70, "Aplicando filtros...")
            df_limpo = self.aplicar_filtros_adicionais(df)
            self.df_resultado = df_limpo # ARMAZENA PARA SALVAMENTO POSTERIOR
            
            # OTIMIZAÇÃO: Preparar dados para Dashboard se for o preset de Mais Vendidos
            if preset_atual and "Mais Vendidos" in preset_atual.get("nome", ""):
                 try:
                     # BUSCA INTELIGENTE DE COLUNA DE QUANTIDADE
                     col_quantidade = -1
                     possiveis_nomes = ['Quantidade', 'Qty', 'Qtd', 'Venda', 'Total']
                     
                     # Tenta achar pelo nome
                     for i, col in enumerate(df_limpo.columns):
                         if any(token.lower() in str(col).lower() for token in possiveis_nomes):
                             col_quantidade = i
                             break
                     
                     # Se não achou pelo nome, tenta a coluna 25 original (pode ter mudado)
                     if col_quantidade == -1:
                         if len(df_limpo.columns) > 25: col_quantidade = 25
                         else: col_quantidade = len(df_limpo.columns) - 1 # Última coluna como fallback
                     
                     self.log(f"📊 Gerando dashboard (Coluna detectada: {df_limpo.columns[col_quantidade]})")

                     def clean_numeric(val):
                         if pd.isna(val): return 0.0
                         if isinstance(val, (int, float)): return float(val)
                         s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
                         try: return float(s)
                         except: return 0.0
                         
                     df_limpo['qty_clean'] = df_limpo.iloc[:, col_quantidade].apply(clean_numeric)
                     self.exibir_dashboard(df_limpo, container=self.dash_container_limpeza)
                 except Exception as dex:
                     self.log(f"⚠️ Não foi possível gerar gráfico: {dex}")

            # Finalizar sem salvar automaticamente
            tempo_execucao = (datetime.now() - inicio).total_seconds()
            
            self.log("")
            self.log("=" * 60)
            self.log("✅ PROCESSAMENTO CONCLUÍDO!")
            self.log(f"   Total de linhas: {len(df_limpo)}")
            self.log(f"   Tempo: {tempo_execucao:.2f} segundos")
            self.log("=" * 60)
            
            self.status_label.config(text="✅ Processamento concluído! Agora você pode salvar.")
            
            # Mostrar botão de salvar e container de saída
            self.root.after(0, lambda: [
                self.output_frame.pack(fill=tk.X, pady=5),
                self.btn_salvar_limpeza.pack(fill=tk.X, pady=5)
            ])
            
        except Exception as e:
            self.log("")
            self.log(f"❌ ERRO: {str(e)}")
            self.log("")
            self.status_label.config(text="❌ Erro no processamento")
            messagebox.showerror("Erro", f"Erro ao processar planilha:\n\n{str(e)}")
        
        finally:
            self.processando = False
            self.btn_processar_limpeza.config(state="normal", text="🚀 INICIAR LIMPEZA")
            self.btn_processar_resumo.config(state="normal", text="📊 GERAR RESUMO AGORA")
    
    # Funções de processamento (mesmas do código anterior, adaptadas para GUI)
    
    def salvar_resultado_processado(self):
        """Método para salvar o resultado após o processamento"""
        if self.df_resultado is None:
            messagebox.showwarning("Aviso", "Nenhum resultado processado para salvar!")
            return
            
        if not self.caminho_saida.get():
            self.selecionar_arquivo_saida()
            
        if self.caminho_saida.get():
            try:
                self.set_progress(50, "Salvando arquivo...")
                self.salvar_planilha(self.df_resultado, self.caminho_saida.get())
                messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso:\n{os.path.basename(self.caminho_saida.get())}")
                self.set_progress(100, "Salvo!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar: {e}")
    
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
        """Carregar planilha do Excel com aba específica, tratando corrupção em .xls"""
        try:
            aba = self.aba_selecionada.get()
            self.log(f"   Aba selecionada: {aba}")
            
            try:
                if caminho.lower().endswith('.xls'):
                    # Motor xlrd com flag para ignorar corrupções leves de cabeçalho
                    df = pd.read_excel(caminho, sheet_name=aba, engine='xlrd', engine_kwargs={'ignore_workbook_corruption': True})
                else:
                    df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
            except Exception as e:
                # FALLBACK AUTOMÁTICO DE MOTOR
                engine_alt = 'openpyxl' if caminho.lower().endswith('.xls') else 'xlrd'
                self.log(f"⚠️ Falha no motor primário ({e}). Tentando {engine_alt}...")
                
                if engine_alt == 'xlrd':
                    df = pd.read_excel(caminho, sheet_name=aba, engine='xlrd', engine_kwargs={'ignore_workbook_corruption': True})
                else:
                    df = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl')
                
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

    def get_preset_atual(self):
        """Retorna o dicionário do preset selecionado"""
        nome = self.nome_preset.get()
        for p in self.presets:
            if p["nome"] == nome:
                return p
        return None

    def get_preset_tipo(self):
        """Retorna o tipo do preset selecionado"""
        p = self.get_preset_atual()
        return p.get("tipo", "limpeza") if p else "limpeza"

    def processar_resumo_pedidos(self, df):
        """Calcula e exibe o resumo dos pedidos com Dashboard"""
        try:
            self.set_progress(30, "Calculando métricas...")
            
            col_pedidos = 1  # B
            col_quantidade = 25 # Z
            col_preco = 26 # AA
            
            def clean_numeric(val):
                if pd.isna(val): return 0.0
                if isinstance(val, (int, float)): return float(val)
                s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
                try: return float(s)
                except: return 0.0

            df_calc = df.copy()
            pedidos_unicos = df.iloc[:, col_pedidos].nunique()
            df_calc['qty_clean'] = df.iloc[:, col_quantidade].apply(clean_numeric)
            total_itens = int(df_calc['qty_clean'].sum())
            df_calc['price_clean'] = df.iloc[:, col_preco].apply(clean_numeric)
            df_calc['total_row'] = df_calc['qty_clean'] * df_calc['price_clean']
            valor_total = df_calc['total_row'].sum()
            
            self.set_progress(70, "Gerando dashboard...")
            
            # Exibição no Log
            self.log("-" * 40)
            self.log(f"� ITENS: {total_itens} | 🎫 PEDIDOS: {pedidos_unicos}")
            self.log(f"💰 VALOR TOTAL: R$ {valor_total:,.2f}")
            self.log("-" * 40)
            
            self.exibir_dashboard(df_calc, total_itens, pedidos_unicos, valor_total)
            self.set_progress(100, "Resumo concluído!")

        except Exception as e:
            self.log(f"❌ Erro no resumo: {e}")
            raise e

    def exibir_dashboard(self, df, total_itens=0, pedidos_unicos=0, valor_total=0, container=None):
        """Gera gráficos usando matplotlib e integra ao tkinter"""
        target_container = container if container else self.dash_container
        
        def atualizar_gui():
            # Atualizar Labels dos Cards (apenas se estivermos no resumo)
            if target_container == self.dash_container:
                self.lbl_stat_itens.config(text=f"{total_itens:,}".replace(",", "."))
                self.lbl_stat_pedidos.config(text=f"{pedidos_unicos:,}".replace(",", "."))
                self.lbl_stat_valor.config(text=f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            # Limpar gráficos anteriores no container específico
            for widget in target_container.winfo_children():
                widget.destroy()
            
            plt.rcParams['text.color'] = '#cdd6f4'
            plt.rcParams['axes.labelcolor'] = '#cdd6f4'
            plt.rcParams['xtick.color'] = '#cdd6f4'
            plt.rcParams['ytick.color'] = '#cdd6f4'

            fig, ax = plt.subplots(figsize=(6, 3), dpi=100)
            fig.patch.set_facecolor('#1e1e2e')
            ax.set_facecolor('#313244')
            
            try:
                # Determina coluna de nome (geralmente a primeira após limpeza ou a 3ª na original)
                # Na limpeza o SKU costuma ser mantido. Vamos tentar inferir.
                col_nome = df.columns[0]
                if 'SKU' in df.columns: col_nome = 'SKU'
                elif 'Produto' in df.columns: col_nome = 'Produto'
                
                if 'qty_clean' in df.columns:
                    top_10 = df.groupby(col_nome)['qty_clean'].sum().nlargest(10)
                    top_10.plot(kind='barh', ax=ax, color='#89b4fa')
                    ax.set_title("Top 10 Itens por Qtd", color='#89b4fa', weight='bold')
                    ax.invert_yaxis()
                    plt.tight_layout()
                    
                    canvas = FigureCanvasTkAgg(fig, master=target_container)
                    canvas.draw()
                    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            except Exception as e:
                self.log(f"⚠️ Erro no gráfico: {e}")

        self.root.after(0, atualizar_gui)


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
