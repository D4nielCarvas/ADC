import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter import ttk
import threading
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
from core_logic import ADCLogic


class LimpadorPlanilhaGUI:
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
        self.log_limpeza = None   # Buffer para log de limpeza
        self.log_resumo = None    # Buffer para log de resumo
        
        # Lógica central
        self.logic = ADCLogic()

        # Novas Variáveis Multi-Relatório
        self.presets = self.logic.presets
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
        caminho_config = os.path.join(os.path.dirname(__file__), "config.json")
        if os.path.exists(caminho_config):
            try:
                with open(caminho_config, 'r', encoding='utf-8') as f:
                    return json.load(f).get("presets", [])
            except Exception:
                return []
        return []
        
    def configurar_estilos(self):
        """Configurar o tema moderno do aplicativo (Catppuccin Mocha Inspired)"""
        style = ttk.Style()
        
        # --- PALETA DE CORES (Catppuccin Mocha) ---
        self.colors = {
            "base": "#1e1e2e",       # Fundo Principal
            "mantle": "#181825",     # Sidebar / Fundo Secundário
            "crust": "#11111b",      # Inputs / Logs
            "surface0": "#313244",   # Cards Baixos
            "surface1": "#45475a",   # Cards Altos / Hover
            "overlay0": "#6c7086",   # Bordas sutis
            "text": "#cdd6f4",       # Texto Principal
            "subtext": "#a6adc8",    # Texto Secundário
            "mauve": "#cba6f7",      # Acento Principal (Roxo Suave)
            "blue": "#89b4fa",       # Acento Secundário (Azul)
            "green": "#a6e3a1",      # Sucesso
            "red": "#f38ba8",        # Erro
            "peach": "#fab387",      # Aviso
            "sapphire": "#74c7ec",   # Detalhes
        }

        self.root.configure(bg=self.colors["base"])
        style.theme_use('clam')
        
        # --- CONFIGURAÇÃO DE ESTILOS TTK ---
        
        # 1. Frames e Containers
        style.configure("TFrame", background=self.colors["base"])
        style.configure("Sidebar.TFrame", background=self.colors["mantle"])
        style.configure("Card.TFrame", background=self.colors["surface0"], relief="flat", borderwidth=0)
        
        # 2. Labels
        style.configure("TLabel", 
            background=self.colors["base"], 
            foreground=self.colors["text"], 
            font=("Segoe UI", 10)
        )
        style.configure("Card.TLabel",
            background=self.colors["surface0"], 
            foreground=self.colors["text"],
            font=("Segoe UI", 10)
        )
        
        # Títulos
        style.configure("Title.TLabel", 
            background=self.colors["mantle"], 
            foreground=self.colors["mauve"], 
            font=("Segoe UI", 22, "bold")
        )
        style.configure("Header.TLabel", 
            background=self.colors["surface0"], 
            foreground=self.colors["blue"], 
            font=("Segoe UI", 12, "bold")
        )
        style.configure("Sub.TLabel",
            background=self.colors["base"],
            foreground=self.colors["subtext"],
            font=("Segoe UI", 9)
        )

        # 3. Botões de Navegação (Sidebar)
        style.configure("Nav.TButton", 
            padding=(20, 12), 
            font=("Segoe UI", 11),
            background=self.colors["mantle"],
            foreground=self.colors["subtext"], 
            anchor="w",
            borderwidth=0
        )
        style.map("Nav.TButton",
            background=[('active', self.colors["surface0"]), ('pressed', self.colors["surface1"])],
            foreground=[('active', self.colors["mauve"])]
        )

        # 4. Botões de Ação (Accent)
        style.configure("Accent.TButton", 
            padding=(15, 12), 
            font=("Segoe UI", 11, "bold"),
            background=self.colors["mauve"],
            foreground=self.colors["crust"], 
            borderwidth=0,
            relief="flat"
        )
        style.map("Accent.TButton",
            background=[('active', '#d0bef9'), ('pressed', '#bfa0e8')]
        )

        # 5. Botões Secundários
        style.configure("Secondary.TButton", 
            padding=(10, 8), 
            font=("Segoe UI", 9),
            background=self.colors["surface1"],
            foreground=self.colors["text"],
            borderwidth=0
        )
        style.map("Secondary.TButton",
            background=[('active', '#585b70')]
        )
        
        # 6. Inputs e Widgets
        style.configure("TEntry", 
            fieldbackground=self.colors["crust"], 
            foreground=self.colors["text"],
            insertcolor=self.colors["text"],
            borderwidth=0,
            padding=5
        )
        
        style.configure("TCombobox", 
            fieldbackground=self.colors["crust"], 
            background=self.colors["surface0"], 
            foreground=self.colors["text"],
            arrowcolor=self.colors["mauve"],
            borderwidth=0
        )
        # Hack para Combobox ficar escura no drop
        self.root.option_add('*TCombobox*Listbox.background', self.colors["crust"])
        self.root.option_add('*TCombobox*Listbox.foreground', self.colors["text"])
        self.root.option_add('*TCombobox*Listbox.selectBackground', self.colors["mauve"])
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.colors["crust"])

        # 7. LabelFrames
        style.configure("TLabelframe", 
            background=self.colors["surface0"], 
            foreground=self.colors["blue"], 
            font=("Segoe UI", 10, "bold"),
            borderwidth=1,
            relief="solid",
            bordercolor=self.colors["surface1"]
        )
        style.configure("TLabelframe.Label", background=self.colors["surface0"], foreground=self.colors["blue"])
        style.configure("TCheckbutton", background=self.colors["surface0"], foreground=self.colors["text"])

        # 8. Stats Cards (Resumo)
        style.configure("Stat.TLabel", 
            background=self.colors["surface0"], 
            foreground=self.colors["green"], 
            font=("Segoe UI", 24, "bold")
        )
        style.configure("StatDesc.TLabel", 
            background=self.colors["surface0"], 
            foreground=self.colors["subtext"], 
            font=("Segoe UI", 10, "bold")
        )

    def criar_interface(self):
        """Criar interface com sidebar e páginas"""
        self.configurar_estilos()
        
        # Layout Principal
        self.sidebar = ttk.Frame(self.root, style="Sidebar.TFrame", width=260) # Sidebar mais larga
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        self.main_content = ttk.Frame(self.root, style="TFrame", padding="40") # Mais padding no conteúdo
        self.main_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- SIDEBAR CONTENT ---
        # Logo Area
        logo_area = ttk.Frame(self.sidebar, style="Sidebar.TFrame", padding=(20, 40, 20, 20))
        logo_area.pack(fill=tk.X)
        ttk.Label(logo_area, text="✨ ADC", style="Title.TLabel").pack(anchor="w")
        ttk.Label(logo_area, text="Advanced Data Cleaner", font=("Segoe UI", 9), foreground=self.colors["subtext"], background=self.colors["mantle"]).pack(anchor="w")

        # Nav Buttons
        nav_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        nav_frame.pack(fill=tk.X, pady=20)

        self.btn_nav_limpeza = ttk.Button(nav_frame, text="   🧹  Limpeza", style="Nav.TButton", command=lambda: self.mudar_pagina("limpeza"))
        self.btn_nav_limpeza.pack(fill=tk.X, padx=10, pady=5)

        self.btn_nav_resumo = ttk.Button(nav_frame, text="   📊  Resumo", style="Nav.TButton", command=lambda: self.mudar_pagina("resumo"))
        self.btn_nav_resumo.pack(fill=tk.X, padx=10, pady=5)

        self.btn_nav_config = ttk.Button(nav_frame, text="   ⚙️  Configurações", style="Nav.TButton", command=lambda: self.mudar_pagina("config"))
        self.btn_nav_config.pack(fill=tk.X, padx=10, pady=5)

        # Espaçador inferior
        ttk.Frame(self.sidebar, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)

        self.btn_full = ttk.Button(self.sidebar, text="🔲 Tela Cheia", command=self.toggle_fullscreen, style="Secondary.TButton")
        self.btn_full.pack(fill=tk.X, padx=20, pady=20)
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

        # Log e Status (Barra inferior embutida no conteúdo principal)
        self.status_label = tk.Label(self.root, text="✨ Sistema Pronto", bg=self.colors["mantle"], fg=self.colors["subtext"], font=("Segoe UI", 9), anchor=tk.W, padx=15, pady=5)
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
            
            # Reset visual dos botões
            style = ttk.Style()
            # Reset não é trivial no tkinter stock, mas vamos focar no conteúdo

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

    def montar_pagina_header(self, parent, titulo, icone, subtitulo=""):
        header = ttk.Frame(parent)
        header.pack(fill=tk.X, pady=(0, 30))
        
        # Icone grande se possivel? Não, texto mesmo.
        ttk.Label(header, text=f"{icone}  {titulo}", font=("Segoe UI", 26, "bold"), foreground=self.colors["text"]).pack(anchor="w")
        if subtitulo:
            ttk.Label(header, text=subtitulo, style="Sub.TLabel").pack(anchor="w", pady=(5, 0))

    def montar_pagina_limpeza(self):
        """Constrói os widgets da página de limpeza"""
        self.montar_pagina_header(self.container_limpeza, "Limpeza de Planilhas", "🧹", "Processamento e higienização automática de dados.")
        
        # Card Arquivos
        file_card = ttk.Frame(self.container_limpeza, style="Card.TFrame", padding=20)
        file_card.pack(fill=tk.X, pady=10)
        ttk.Label(file_card, text="ARQUIVO DE ENTRADA", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 15))

        # Entrada
        input_frame = ttk.Frame(file_card, style="Card.TFrame")
        input_frame.pack(fill=tk.X)
        
        # Container de Input estiloso
        entry_container = ttk.Frame(input_frame, style="Card.TFrame") # Wrapper
        entry_container.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Entry(entry_container, textvariable=self.caminho_entrada, font=("Segoe UI", 10)).pack(fill=tk.X, ipady=5)
        
        ttk.Button(input_frame, text="� Selecionar", width=12, command=self.selecionar_arquivo_entrada, style="Accent.TButton").pack(side=tk.LEFT, padx=(15, 0))

        # Configurações
        config_card = ttk.Frame(self.container_limpeza, style="Card.TFrame", padding=20)
        config_card.pack(fill=tk.X, pady=20)
        ttk.Label(config_card, text="REGRAS DE LIMPEZA", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 15))

        grid = ttk.Frame(config_card, style="Card.TFrame")
        grid.pack(fill=tk.X)
        
        # Linha 1
        ttk.Label(grid, text="Preset de Configuração", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 5))
        ttk.Label(grid, text="Aba da Planilha", style="Card.TLabel").grid(row=0, column=1, sticky="w", pady=(0, 5), padx=20)
        
        self.preset_combo = ttk.Combobox(grid, textvariable=self.nome_preset, values=[p["nome"] for p in self.presets], state="readonly", width=30)
        self.preset_combo.grid(row=1, column=0, sticky="ew", ipady=3)
        self.preset_combo.bind("<<ComboboxSelected>>", self.aplicar_preset)

        self.aba_combo = ttk.Combobox(grid, textvariable=self.aba_selecionada, state="readonly", width=20)
        self.aba_combo.grid(row=1, column=1, sticky="ew", padx=20, ipady=3)
        self.aba_combo.bind("<<ComboboxSelected>>", self.atualizar_colunas_aba)

        # Linha 2
        ttk.Label(grid, text="Colunas Para Deletar (Índices)", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(20, 5))
        self.indices_entry = ttk.Entry(grid, width=30)
        self.indices_entry.grid(row=3, column=0, columnspan=2, sticky="ew", ipady=3)
        ttk.Label(grid, text="Ex: 1, 2, 5 (Separado por vírgula)", style="Sub.TLabel", background=self.colors["surface0"]).grid(row=4, column=0, sticky="w")

        # Refinamento (Checkboxes)
        chk_frame = ttk.Frame(config_card, style="Card.TFrame")
        chk_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Checkbutton(chk_frame, text="Remover Duplicadas", variable=self.remover_duplicadas).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(chk_frame, text="Remover Linhas Vazias", variable=self.remover_vazias).pack(side=tk.LEFT)

        # Ações
        action_frame = ttk.Frame(self.container_limpeza)
        action_frame.pack(fill=tk.X, pady=10)
        
        self.btn_processar_limpeza = ttk.Button(action_frame, text="🚀  INICIAR PROCESSAMENTO", command=self.iniciar_processamento, style="Accent.TButton")
        self.btn_processar_limpeza.pack(side=tk.RIGHT)

        # Output e Save (Inicialmente ocultos)
        self.output_frame = ttk.Frame(file_card, style="Card.TFrame", padding=(0, 20, 0, 0))
        # self.output_frame.pack(fill=tk.X) 
        
        ttk.Label(self.output_frame, text="SALVAR RESULTADO EM", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))
        save_inner = ttk.Frame(self.output_frame, style="Card.TFrame")
        save_inner.pack(fill=tk.X)
        self.entry_saida = ttk.Entry(save_inner, textvariable=self.caminho_saida, font=("Segoe UI", 10))
        self.entry_saida.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        self.btn_salvar_limpeza = ttk.Button(save_inner, text="💾 Salvar", command=self.salvar_resultado_processado, style="Accent.TButton")
        self.btn_salvar_limpeza.pack(side=tk.LEFT, padx=(10, 0))

        # Log
        self.log_limpeza = self.montar_log(self.container_limpeza)

    def montar_pagina_resumo(self):
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

        self.log_resumo = self.montar_log(self.container_resumo)

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
            self.logic.salvar_presets(self.presets)
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
            ok, msg = self.logic.salvar_presets(self.presets)
            if not ok:
                 messagebox.showerror("Erro", msg)
            self.atualizar_lista_presets_config()
            self.preset_combo['values'] = [p["nome"] for p in self.presets]

    # Método salvar_presets foi removido (agora em ADCLogic)

    def atualizar_lista_presets_config(self):
        """Atualiza a listagem visual de presets na página de config"""
        self.preset_listbox.delete(0, tk.END)
        for p in self.presets:
            self.preset_listbox.insert(tk.END, f" {p['nome']} ({p.get('tipo', 'limpeza')})")

    def montar_log(self, parent):
        log_frame = ttk.Frame(parent, padding=(0, 20, 0, 0))
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(log_frame, text="DIÁRIO DE ATIVIDADES", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 5))
        
        widget = scrolledtext.ScrolledText(log_frame, height=5, font=("Consolas", 9), 
            bg=self.colors["crust"], 
            fg=self.colors["green"], 
            borderwidth=0, 
            padx=10, pady=10
        )
        widget.pack(fill=tk.BOTH, expand=True)
        return widget

    def toggle_fullscreen(self):
        """Alternar entre modo tela cheia e janela"""
        self.is_fullscreen = not self.is_fullscreen
        self.root.attributes("-fullscreen", self.is_fullscreen)
        if not self.is_fullscreen:
            self.root.geometry("1000x850") # Sincronizado com o __init__
        
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
            # Usa o listar_abas do logic (que usa cache) para pegar info basica ou ler header
            # Mas aqui precisamos das colunas.
            # Vamos simplificar e ler apenas o header
            df_temp = self.logic.carregar_planilha(arquivo, aba)
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
            # Carregar abas dinamicamente com OTIMIZAÇÃO (Cache)
            try:
                # Usa método padronizado do logic que já trara cache
                self.lista_abas = self.logic.listar_abas(arquivo, log_callback=self.log)
                
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
        """Adicionar mensagem ao log visual da página ativa"""
        def atualizar():
            alvo = None
            if self.pagina_atual == "limpeza":
                alvo = self.log_limpeza
            elif self.pagina_atual == "resumo":
                alvo = self.log_resumo
                
            if alvo:
                alvo.insert(tk.END, mensagem + "\n")
                alvo.see(tk.END)
        self.root.after(0, atualizar)
    
    def limpar_log(self):
        """Limpar área de log da página ativa"""
        alvo = self.log_limpeza if self.pagina_atual == "limpeza" else self.log_resumo
        if alvo:
            alvo.delete(1.0, tk.END)
    
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
            
            # Passo 2: Carregar planilha
            self.set_progress(10, "Lendo planilha...")
            df = self.carregar_planilha(caminho_entrada)
            
            # Passo 3: Validar índices
            self.set_progress(30, "Validando colunas...")
            self.validar_indices_colunas(df, index_delet)
            
            # Passo 4: Deletar colunas
            self.set_progress(50, "Limpando dados...")
            df = self.deletar_colunas(df, index_delet)
            
            # Passo 5: Aplicar Filtros Adicionais
            self.set_progress(70, "Aplicando filtros...")
            df_limpo = self.aplicar_filtros_adicionais(df)
            self.df_resultado = df_limpo # ARMAZENA PARA SALVAMENTO POSTERIOR
            

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
    
    
    # Métodos de entrada/carregamento removidos (Delegado para ADCLogic)
    
    # Métodos de processamento removidos (Delegado para ADCLogic)
    
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

    def processar_resumo_pedidos(self, resultado_dict):
        """Calcula e exibe o resumo dos pedidos com Dashboard"""
        try:
            self.set_progress(70, "Gerando dashboard...")
            
            total_itens = resultado_dict["total_itens"]
            pedidos_unicos = resultado_dict["total_pedidos"]
            valor_total = resultado_dict["valor_total"]
            df_calc = resultado_dict.get("df")

            # Exibição no Log
            self.log("-" * 40)
            self.log(f"📦 ITENS: {total_itens} | 🎫 PEDIDOS: {pedidos_unicos}")
            self.log(f"💰 VALOR TOTAL: R$ {valor_total:,.2f}")
            self.log("-" * 40)
            
            if df_calc is not None:
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
                # 1. Detecção Inteligente da Coluna de Nome/Produto
                col_nome = df.columns[0]
                termos_produto = ['nome', 'produto', 'sku', 'descri', 'item', 'material']
                for col in df.columns:
                    if any(termo in str(col).lower() for termo in termos_produto):
                        col_nome = col
                        break
                
                if 'qty_clean' in df.columns:
                    # 2. Agrupamento e Ordenação Robusta (Top 10 real)
                    # Normalização profunda: strip e upper para evitar duplicatas por caixa
                    df_plot = df.copy()
                    df_plot[col_nome] = df_plot[col_nome].astype(str).str.strip().str.upper()
                    
                    if target_container == self.dash_container:
                        # ABA RESUMO: Ranking por frequência de pedidos (contagem)
                        top_10 = df_plot.groupby(col_nome).size().sort_values(ascending=False).head(10)
                        ax.set_title("🏆 TOP 10 SKUs MAIS PEDIDOS (Frequência)", color='#89b4fa', weight='bold', pad=15)
                    else:
                        # FALLBACK / OUTRAS ABAS: Ranking por volume de quantidade (soma)
                        top_10 = df_plot.groupby(col_nome)['qty_clean'].sum().sort_values(ascending=False).head(10)
                        ax.set_title(f"Top 10 por {col_nome} (Volume)", color='#89b4fa', weight='bold', pad=15)
                    
                    # Se não houver dados, avisar
                    if top_10.empty:
                        ax.text(0.5, 0.5, "Sem dados de quantidade para exibir", ha='center', va='center', color='#f38ba8')
                    else:
                        top_10.plot(kind='barh', ax=ax, color='#89b4fa')
                        ax.set_title(f"Top 10 por {col_nome}", color='#89b4fa', weight='bold', pad=15)
                        ax.invert_yaxis()
                        plt.tight_layout()
                    
                    # 3. Ordem de empilhamento: Botão primeiro (topo) para não sumir
                    btn_expand = ttk.Button(target_container, text="🔍 EXPANDIR DASHBOARD", command=lambda f=fig: self.expandir_dashboard(f), style="Secondary.TButton")
                    btn_expand.pack(pady=(0, 10))

                    canvas = FigureCanvasTkAgg(fig, master=target_container)
                    canvas.draw()
                    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            except Exception as e:
                self.log(f"⚠️ Erro no gráfico: {e}")

        self.root.after(0, atualizar_gui)

    def expandir_dashboard(self, fig_orig):
        """Abre o gráfico em uma nova janela em tela cheia"""
        top = tk.Toplevel(self.root)
        top.title("Dashboard Expandido - ADC Pro")
        top.attributes("-fullscreen", True)
        top.configure(bg="#1e1e2e")
        
        # Container
        frame = ttk.Frame(top, style="TFrame", padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Recriar a figura para a nova janela (evita conflitos de canvas)
        canvas = FigureCanvasTkAgg(fig_orig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Botões de controle flutuantes
        btn_close = ttk.Button(top, text="[ ESC ] FECHAR", command=top.destroy, style="Accent.TButton")
        btn_close.place(relx=0.95, rely=0.05, anchor="ne")
        
        top.bind("<Escape>", lambda e: top.destroy())
        top.focus_set()


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
