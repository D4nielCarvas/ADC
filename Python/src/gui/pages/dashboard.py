import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from gui.styles import ThemeConfig
# Para gráficos se quiser implementar depois: import matplotlib.pyplot as plt

class DashboardPage(ttk.Frame):
    def __init__(self, parent, logic, status_callback):
        super().__init__(parent)
        self.logic = logic
        self.status_callback = status_callback
        self.colors = ThemeConfig.get_colors()
        
        self.arquivos_selecionados = []  # Lista de arquivos para processar
        
        self.lbl_stats = {} # refs para labels de numeros
        
        self._setup_ui()

    def set_ui_controllers(self, progress_bar, lbl_status, progress_frame):
        self.progress_bar = progress_bar
        self.lbl_status = lbl_status

    def on_show(self):
        pass

    def _setup_ui(self):
        # Header
        self._montar_header()
        
        # Inputs - SIMPLIFIED: Only file selection, no sheet selection
        card = ttk.Frame(self, style="Card.TFrame", padding=15)
        card.pack(fill=tk.X, pady=10)
        
        f1 = ttk.Frame(card, style="Card.TFrame")
        f1.pack(fill=tk.X)
        ttk.Label(f1, text="Arquivos:", background=self.colors["surface0"]).pack(side=tk.LEFT)
        self.lbl_arquivos = ttk.Label(f1, text="Nenhum arquivo selecionado", 
                                      foreground=self.colors["subtext"], background=self.colors["surface0"])
        self.lbl_arquivos.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(f1, text="📂", width=4, command=self.selecionar_arquivos, style="Secondary.TButton").pack(side=tk.LEFT)

        # Info label
        info_label = ttk.Label(card, text="💡 Selecione um ou mais arquivos. O sistema usará a primeira aba de cada arquivo.", 
                              foreground=self.colors["subtext"], font=("Segoe UI", 9))
        info_label.pack(pady=(10,0))

        # Stats Row
        stats = ttk.Frame(self)
        stats.pack(fill=tk.X, pady=20)
        self.lbl_stats['itens'] = self._criar_card_stat(stats, "ITENS", "0")
        self.lbl_stats['pedidos'] = self._criar_card_stat(stats, "PEDIDOS", "0")
        self.lbl_stats['valor'] = self._criar_card_stat(stats, "VALOR TOTAL", "R$ 0,00")

        # Action
        ttk.Button(self, text="📊 GERAR DASHBOARD", command=self.gerar_dashboard, style="Accent.TButton").pack(fill=tk.X)
        
        # Log (Simples)
        self.log_label = ttk.Label(self, text="Aguardando...", foreground=self.colors["subtext"])
        self.log_label.pack(pady=20)

    def _montar_header(self):
        h = ttk.Frame(self)
        h.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(h, text="📊 Resumo de Pedidos", font=("Segoe UI", 26, "bold"), foreground=self.colors["text"]).pack(anchor="w")

    def _criar_card_stat(self, parent, titulo, valor_inicial):
        card = ttk.Frame(parent, style="Card.TFrame", padding=15)
        card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        ttk.Label(card, text=titulo, style="StatDesc.TLabel").pack()
        lbl = ttk.Label(card, text=valor_inicial, style="Stat.TLabel")
        lbl.pack()
        return lbl

    def selecionar_arquivos(self):
        """
        Open file dialog to select one or more Excel files.
        """
        arquivos = filedialog.askopenfilenames(
            title="Selecionar Planilha(s)",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        if arquivos:
            self.arquivos_selecionados = list(arquivos)
            count = len(self.arquivos_selecionados)
            if count == 1:
                self.lbl_arquivos.config(
                    text=f"1 arquivo selecionado",
                    foreground=self.colors["green"]
                )
            else:
                self.lbl_arquivos.config(
                    text=f"{count} arquivos selecionados",
                    foreground=self.colors["green"]
                )
            self.log_label.config(
                text="Arquivos selecionados. Clique em GERAR DASHBOARD para processar.", 
                foreground=self.colors["green"]
            )

    def gerar_dashboard(self):
        """
        Generate dashboard statistics from selected Excel file(s).
        Automatically uses the first sheet of each file.
        """
        if not self.arquivos_selecionados:
            self.log_label.config(text="Selecione pelo menos um arquivo primeiro!", foreground=self.colors["red"])
            return
        
        self.log_label.config(text="Calculando...", foreground=self.colors["subtext"])
        self.update_idletasks()
        
        threading.Thread(target=self._calc_thread, daemon=True).start()

    def _calc_thread(self):
        """
        Background thread for calculating dashboard statistics.
        Handles both successful results and error cases.
        Processes multiple files and combines results:
        - Unique order IDs across all files (no duplicates)
        - Sum of items from all files
        - Sum of values from all files
        """
        try:
            # Combined results
            pedidos_unicos = set()  # Set para garantir IDs unicos
            total_itens = 0
            total_valor = 0.0
            erros = []
            
            # Process each file
            for idx, arquivo in enumerate(self.arquivos_selecionados, 1):
                try:
                    # Use empty string to auto-load first sheet
                    res = self.logic.gerar_resumo(arquivo, "")
                    
                    if 'erro' in res:
                        erros.append(f"{arquivo}: {res['erro']}")
                    else:
                        # Get unique order IDs from this file
                        df = res.get('df')
                        if df is not None:
                            # Column B (index 1) contains order IDs
                            ids_deste_arquivo = set(df.iloc[:, 1].dropna().unique())
                            pedidos_unicos.update(ids_deste_arquivo)
                        
                        # Sum items and values
                        total_itens += res.get('total_itens', 0)
                        total_valor += res.get('valor_total', 0.0)
                        
                except Exception as e:
                    erros.append(f"{arquivo}: {str(e)}")
            
            # Update UI
            def _u():
                if erros and not pedidos_unicos:
                    # All files failed
                    self.log_label.config(text=f"Erro: {erros[0]}", foreground=self.colors["red"])
                    self.lbl_stats['itens'].config(text="0")
                    self.lbl_stats['pedidos'].config(text="0")
                    self.lbl_stats['valor'].config(text="R$ 0,00")
                else:
                    # Success (at least some files processed)
                    self.lbl_stats['itens'].config(text=str(total_itens))
                    self.lbl_stats['pedidos'].config(text=str(len(pedidos_unicos)))
                    self.lbl_stats['valor'].config(text=f"R$ {total_valor:,.2f}")
                    
                    if erros:
                        self.log_label.config(
                            text=f"✓ Processado com {len(erros)} erro(s). Ver console.", 
                            foreground=self.colors["yellow"]
                        )
                        for erro in erros:
                            print(f"[WARNING] {erro}")
                    else:
                        self.log_label.config(
                            text=f"✓ {len(self.arquivos_selecionados)} arquivo(s) processado(s) com sucesso!", 
                            foreground=self.colors["green"]
                        )
                
            self.after(0, _u)
        except Exception as e:
            # Capture exception message before creating closure
            error_message = str(e)
            def _err():
                self.log_label.config(text=f"Erro: {error_message}", foreground=self.colors["red"])
                # Reset stats to zero
                self.lbl_stats['itens'].config(text="0")
                self.lbl_stats['pedidos'].config(text="0")
                self.lbl_stats['valor'].config(text="R$ 0,00")
            self.after(0, _err)
