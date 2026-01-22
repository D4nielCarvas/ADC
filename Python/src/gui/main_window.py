import tkinter as tk
from tkinter import ttk
from gui.styles import ThemeConfig
from gui.pages.cleaner import CleanerPage
from gui.pages.dashboard import DashboardPage
from gui.pages.config import ConfigPage
from core.cleaner import ADCLogic

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("ADC v2.5 Pro")
        self.root.geometry("1000x850")
        
        # Inicializa lógica central (compartilhada)
        self.logic = ADCLogic()
        
        # Aplica estilos
        ThemeConfig.apply_styles(self.root)
        self.colors = ThemeConfig.get_colors()
        
        # Estado
        self.pagina_atual = None
        self.pages = {}
        self.is_fullscreen = False

        self.criar_layout()
        self.navegar("limpeza")

    def criar_layout(self):
        # Sidebar
        self.sidebar = ttk.Frame(self.root, style="Sidebar.TFrame", width=260)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        # Content
        self.main_content = ttk.Frame(self.root, style="TFrame", padding="40")
        self.main_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.montar_sidebar()
        self.montar_paginas()
        self.montar_barra_inferior()

    def montar_sidebar(self):
        # Logo
        logo_area = ttk.Frame(self.sidebar, style="Sidebar.TFrame", padding=(20, 40, 20, 20))
        logo_area.pack(fill=tk.X)
        ttk.Label(logo_area, text="✨ ADC", style="Title.TLabel").pack(anchor="w")
        ttk.Label(logo_area, text="Advanced Data Cleaner", font=("Segoe UI", 9), foreground=self.colors["subtext"], background=self.colors["mantle"]).pack(anchor="w")

        # Nav
        nav_frame = ttk.Frame(self.sidebar, style="Sidebar.TFrame")
        nav_frame.pack(fill=tk.X, pady=20)

        self.btn_nav_limpeza = self._criar_botao_nav(nav_frame, "   🧹  Limpeza", "limpeza")
        self.btn_nav_resumo = self._criar_botao_nav(nav_frame, "   📊  Resumo", "resumo")
        self.btn_nav_config = self._criar_botao_nav(nav_frame, "   ⚙️  Configurações", "config")

        # Espaçador
        ttk.Frame(self.sidebar, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)

        # Fullscreen
        ttk.Button(self.sidebar, text="🔲 Tela Cheia", command=self.toggle_fullscreen, style="Secondary.TButton").pack(fill=tk.X, padx=20, pady=20)
        self.root.bind("<F11>", lambda e: self.toggle_fullscreen())

    def _criar_botao_nav(self, parent, text, page_key):
        btn = ttk.Button(parent, text=text, style="Nav.TButton", command=lambda: self.navegar(page_key))
        btn.pack(fill=tk.X, padx=10, pady=5)
        return btn

    def montar_paginas(self):
        # Instancia as páginas passando o logic compartilhado
        self.pages["limpeza"] = CleanerPage(self.main_content, self.logic, self.atualizar_status)
        self.pages["resumo"] = DashboardPage(self.main_content, self.logic, self.atualizar_status)
        self.pages["config"] = ConfigPage(self.main_content, self.logic, self.atualizar_status)
        
        # Configurar callback de preset update se necessário
        # Ex: Quando config mudar algo, limpar cache ou avisar outras abas. 
        # Por enquanto, eles leem self.logic.presets que é atualizado.

    def montar_barra_inferior(self):
        self.progress_frame = ttk.Frame(self.main_content)
        self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
         # Passar progress bar para as páginas usarem?
         # Melhor: Pages emitem eventos ou chamam método do MainWindow. 
         # Para simplificar, vou injetar a progress_bar nas pages ou criar método aqui.
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, mode='determinate', length=100)
        # Pages acessam via callback ou referência se passarmos self.
        
        self.status_label = tk.Label(self.root, text="✨ Sistema Pronto", bg=self.colors["mantle"], fg=self.colors["subtext"], font=("Segoe UI", 9), anchor=tk.W, padx=15, pady=5)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        # Injeta dependências de UI nas pages
        for page in self.pages.values():
            page.set_ui_controllers(self.progress_bar, self.status_label, self.progress_frame)

    def navegar(self, pagina_key):
        if self.pagina_atual == pagina_key: return

        # Hide all
        for p in self.pages.values():
            p.pack_forget()

        # Show selected with animation simulation
        self.root.after(50, lambda: self._animar_entrada(pagina_key))
        
    def _animar_entrada(self, pagina_key):
        page = self.pages[pagina_key]
        page.pack(fill=tk.BOTH, expand=True)
        self.pagina_atual = pagina_key
        
        # Hooks de entrada
        page.on_show()
        
        msgs = {
            "limpeza": "✨ Modo Limpeza Ativo",
            "resumo": "✨ Modo Resumo Ativo",
            "config": "⚙️ Configurações de Sistema"
        }
        self.atualizar_status(msgs.get(pagina_key, ""))

    def toggle_fullscreen(self):
        self.is_fullscreen = not self.is_fullscreen
        self.root.attributes("-fullscreen", self.is_fullscreen)
        if not self.is_fullscreen:
            self.root.geometry("1000x850")

    def atualizar_status(self, text):
        self.status_label.config(text=text)
