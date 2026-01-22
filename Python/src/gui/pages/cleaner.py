import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from datetime import datetime
from gui.styles import ThemeConfig

class CleanerPage(ttk.Frame):
    def __init__(self, parent, logic, status_callback):
        super().__init__(parent)
        self.logic = logic
        self.status_callback = status_callback
        self.colors = ThemeConfig.get_colors()
        
        # UI Elements Control
        self.progress_bar = None
        self.progress_frame = None
        self.lbl_status = None # será injetado

        # State Variables
        self.caminho_entrada = tk.StringVar()
        self.caminho_saida = tk.StringVar()
        self.nome_preset = tk.StringVar()
        self.aba_selecionada = tk.StringVar()
        self.lista_abas = []
        
        self.remover_duplicadas = tk.BooleanVar(value=True)
        self.remover_vazias = tk.BooleanVar(value=True)
        
        self.processando = False
        self.df_resultado = None

        self._setup_ui()

    def set_ui_controllers(self, progress_bar, lbl_status, progress_frame):
        self.progress_bar = progress_bar
        self.lbl_status = lbl_status
        self.progress_frame = progress_frame

    def on_show(self):
        # Atualiza presets caso tenham mudado na config
        self.preset_combo['values'] = [p["nome"] for p in self.logic.presets]

    def _setup_ui(self):
        # Header
        self._montar_header("Limpeza de Planilhas", "🧹", "Processamento e higienização automática de dados.")
        
        # Card Arquivo
        file_card = ttk.Frame(self, style="Card.TFrame", padding=20)
        file_card.pack(fill=tk.X, pady=10)
        ttk.Label(file_card, text="ARQUIVO DE ENTRADA", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 15))
        
        input_frame = ttk.Frame(file_card, style="Card.TFrame")
        input_frame.pack(fill=tk.X)
        
        entry_container = ttk.Frame(input_frame, style="Card.TFrame")
        entry_container.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Entry(entry_container, textvariable=self.caminho_entrada, font=("Segoe UI", 10)).pack(fill=tk.X, ipady=5)
        ttk.Button(input_frame, text="📂 Selecionar", width=12, command=self.selecionar_arquivo_entrada, style="Accent.TButton").pack(side=tk.LEFT, padx=(15, 0))

        # Config Card
        config_card = ttk.Frame(self, style="Card.TFrame", padding=20)
        config_card.pack(fill=tk.X, pady=20)
        ttk.Label(config_card, text="REGRAS DE LIMPEZA", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 15))

        grid = ttk.Frame(config_card, style="Card.TFrame")
        grid.pack(fill=tk.X)
        
        ttk.Label(grid, text="Preset de Configuração", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 5))
        ttk.Label(grid, text="Aba da Planilha", style="Card.TLabel").grid(row=0, column=1, sticky="w", pady=(0, 5), padx=20)
        
        self.preset_combo = ttk.Combobox(grid, textvariable=self.nome_preset, values=[p["nome"] for p in self.logic.presets], state="readonly", width=30)
        self.preset_combo.grid(row=1, column=0, sticky="ew", ipady=3)
        self.preset_combo.bind("<<ComboboxSelected>>", self.aplicar_preset)

        self.aba_combo = ttk.Combobox(grid, textvariable=self.aba_selecionada, state="readonly", width=20)
        self.aba_combo.grid(row=1, column=1, sticky="ew", padx=20, ipady=3)
        # self.aba_combo.bind("<<ComboboxSelected>>", self.atualizar_colunas_aba) # Opcional se quisermos mostrar colunas

        ttk.Label(grid, text="Colunas Para Deletar (Índices)", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(20, 5))
        self.indices_entry = ttk.Entry(grid, width=30)
        self.indices_entry.grid(row=3, column=0, columnspan=2, sticky="ew", ipady=3)
        ttk.Label(grid, text="Ex: 1, 2, 5 (Separado por vírgula)", style="Sub.TLabel", background=self.colors["surface0"]).grid(row=4, column=0, sticky="w")

        # Filtros
        chk_frame = ttk.Frame(config_card, style="Card.TFrame")
        chk_frame.pack(fill=tk.X, pady=(20, 0))
        ttk.Checkbutton(chk_frame, text="Remover Duplicadas", variable=self.remover_duplicadas).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(chk_frame, text="Remover Linhas Vazias", variable=self.remover_vazias).pack(side=tk.LEFT)

        # Actions
        action_frame = ttk.Frame(self)
        action_frame.pack(fill=tk.X, pady=10)
        self.btn_processar = ttk.Button(action_frame, text="🚀  INICIAR PROCESSAMENTO", command=self.iniciar_processamento, style="Accent.TButton")
        self.btn_processar.pack(side=tk.RIGHT)

        # Output (Hidden)
        self.output_frame = ttk.Frame(file_card, style="Card.TFrame", padding=(0, 20, 0, 0))
        ttk.Label(self.output_frame, text="SALVAR RESULTADO EM", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))
        save_inner = ttk.Frame(self.output_frame, style="Card.TFrame")
        save_inner.pack(fill=tk.X)
        self.entry_saida = ttk.Entry(save_inner, textvariable=self.caminho_saida, font=("Segoe UI", 10))
        self.entry_saida.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        self.btn_salvar = ttk.Button(save_inner, text="💾 Salvar", command=self.salvar_resultado, style="Accent.TButton")
        self.btn_salvar.pack(side=tk.LEFT, padx=(10, 0))

        # Log
        self.log_widget = self._montar_log(self)

    def _montar_header(self, titulo, icone, subtitulo):
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 30))
        ttk.Label(header, text=f"{icone}  {titulo}", font=("Segoe UI", 26, "bold"), foreground=self.colors["text"]).pack(anchor="w")
        ttk.Label(header, text=subtitulo, style="Sub.TLabel").pack(anchor="w", pady=(5, 0))

    def _montar_log(self, parent):
        log_frame = ttk.Frame(parent, padding=(0, 20, 0, 0))
        log_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(log_frame, text="DIÁRIO DE ATIVIDADES", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 5))
        widget = scrolledtext.ScrolledText(log_frame, height=5, font=("Consolas", 9), 
             bg=self.colors["crust"], fg=self.colors["green"], borderwidth=0, padx=10, pady=10)
        widget.pack(fill=tk.BOTH, expand=True)
        return widget

    def log(self, msg):
        def _u():
            self.log_widget.insert(tk.END, msg + "\n")
            self.log_widget.see(tk.END)
        self.after(0, _u)
        
    def limpar_log(self):
        self.log_widget.delete(1.0, tk.END)

    def selecionar_arquivo_entrada(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if arquivo:
            self.caminho_entrada.set(arquivo)
            try:
                self.lista_abas = self.logic.listar_abas(arquivo, log_callback=self.log)
                self.aba_combo['values'] = self.lista_abas
                if self.lista_abas: self.aba_selecionada.set(self.lista_abas[0])
                self.log(f"✓ Arquivo carregado: {os.path.basename(arquivo)}")
            except Exception as e:
                self.log(f"❌ Erro: {e}")

    def aplicar_preset(self, event=None):
        nome = self.nome_preset.get()
        if nome == "Personalizado": return
        for p in self.logic.presets:
            if p["nome"] == nome:
                self.indices_entry.delete(0, tk.END)
                self.indices_entry.insert(0, p.get("colunas_deletar", ""))

    def set_progress(self, val, text=None):
        def _u():
            if not self.progress_bar: return
            if val < 0 or val >= 100: self.progress_bar.pack_forget()
            else:
                if not self.progress_bar.winfo_viewable(): self.progress_bar.pack(fill=tk.X)
                self.progress_bar['value'] = val
            if text: self.lbl_status.config(text=text)
        self.after(0, _u)

    def iniciar_processamento(self):
        if self.processando: return
        if not self.caminho_entrada.get():
             messagebox.showerror("Erro", "Selecione o arquivo!")
             return
             
        self.processando = True
        self.btn_processar.config(state="disabled")
        threading.Thread(target=self._processar_thread, daemon=True).start()

    def _processar_thread(self):
        try:
            self.limpar_log()
            self.log("INICIANDO PROCESSAMENTO...")
            
            caminho = self.caminho_entrada.get()
            indices_str = self.indices_entry.get()
            
            # Parse indices
            try:
                indices = [int(i.strip()) - 1 for i in indices_str.split(",") if i.strip()]
            except ValueError:
                raise ValueError("Índices inválidos. Use números separados por vírgula.")

            opcoes = {
                "remover_duplicadas": self.remover_duplicadas.get(),
                "remover_vazias": self.remover_vazias.get()
            }

            self.set_progress(20, "Lendo arquivo...")
            df = self.logic.processar_limpeza(
                caminho, 
                self.aba_selecionada.get(), 
                indices, 
                opcoes, 
                log_callback=self.log
            )
            
            self.df_resultado = df
            
            self.set_progress(100, "Concluído!")
            self.log("✅ Concluído com sucesso!")
            
            self.after(0, lambda: [
                self.output_frame.pack(fill=tk.X, pady=5),
                self.btn_salvar.pack(fill=tk.X, pady=5)
            ])
            
        except Exception as e:
            self.log(f"❌ Erro: {e}")
            self.set_progress(0, "Erro")
        finally:
            self.processando = False
            self.after(0, lambda: self.btn_processar.config(state="normal"))

    def salvar_resultado(self):
        if self.df_resultado is None: return
        
        if not self.caminho_saida.get():
             path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
             if path: self.caminho_saida.set(path)
        
        if self.caminho_saida.get():
            try:
                self.logic.salvar_planilha(self.df_resultado, self.caminho_saida.get())
                messagebox.showinfo("Sucesso", "Arquivo salvo!")
                self.log(f"💾 Salvo em: {self.caminho_saida.get()}")
            except Exception as e:
                messagebox.showerror("Erro", str(e))
