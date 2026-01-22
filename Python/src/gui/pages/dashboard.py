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
        
        self.caminho_entrada = tk.StringVar()
        self.aba_selecionada = tk.StringVar()
        
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
        
        # Inputs
        card = ttk.Frame(self, style="Card.TFrame", padding=15)
        card.pack(fill=tk.X, pady=10)
        
        f1 = ttk.Frame(card, style="Card.TFrame")
        f1.pack(fill=tk.X)
        ttk.Label(f1, text="Arquivo:", background=self.colors["surface0"]).pack(side=tk.LEFT)
        ttk.Entry(f1, textvariable=self.caminho_entrada).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        ttk.Button(f1, text="📂", width=4, command=self.selecionar_arquivo, style="Secondary.TButton").pack(side=tk.LEFT)

        f2 = ttk.Frame(card, style="Card.TFrame")
        f2.pack(fill=tk.X, pady=(10,0))
        ttk.Label(f2, text="Aba:", background=self.colors["surface0"]).pack(side=tk.LEFT)
        self.aba_combo = ttk.Combobox(f2, textvariable=self.aba_selecionada, state="readonly")
        self.aba_combo.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

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

    def selecionar_arquivo(self):
        arq = filedialog.askopenfilename()
        if arq:
            self.caminho_entrada.set(arq)
            try:
                abas = self.logic.listar_abas(arq)
                self.aba_combo['values'] = abas
                if abas: self.aba_selecionada.set(abas[0])
            except: pass

    def gerar_dashboard(self):
        if not self.caminho_entrada.get(): return
        
        self.log_label.config(text="Calculando...")
        self.update_idletasks()
        
        threading.Thread(target=self._calc_thread, daemon=True).start()

    def _calc_thread(self):
        try:
            res = self.logic.gerar_resumo(self.caminho_entrada.get(), self.aba_selecionada.get())
            
            def _u():
                self.lbl_stats['itens'].config(text=str(res['total_itens']))
                self.lbl_stats['pedidos'].config(text=str(res['total_pedidos']))
                self.lbl_stats['valor'].config(text=f"R$ {res['valor_total']:,.2f}")
                self.log_label.config(text="Atualizado!")
                
            self.after(0, _u)
        except Exception as e:
            self.after(0, lambda: self.log_label.config(text=f"Erro: {e}"))
