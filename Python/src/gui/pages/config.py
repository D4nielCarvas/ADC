import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from gui.styles import ThemeConfig

class ConfigPage(ttk.Frame):
    def __init__(self, parent, logic, status_callback):
        super().__init__(parent)
        self.logic = logic
        self.status_callback = status_callback
        self.colors = ThemeConfig.get_colors()
        self._setup_ui()

    def set_ui_controllers(self, *args):
        pass

    def on_show(self):
        self.atualizar_lista()

    def _setup_ui(self):
        ttk.Label(self, text="⚙️ Configurações e Presets", font=("Segoe UI", 26, "bold"), foreground=self.colors["text"]).pack(anchor="w", pady=(0, 20))

        card = ttk.Frame(self, style="Card.TFrame", padding=20)
        card.pack(fill=tk.BOTH, expand=True)

        ttk.Label(card, text="GERENCIAR PRESETS", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))

        # Listbox
        list_frame = ttk.Frame(card, style="Card.TFrame")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.listbox = tk.Listbox(list_frame, bg=self.colors["base"], fg=self.colors["text"], borderwidth=0, highlightthickness=1)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=sb.set)

        # Buttons
        btns = ttk.Frame(card, style="Card.TFrame")
        btns.pack(fill=tk.X, pady=15)
        ttk.Button(btns, text="➕ Novo", style="Secondary.TButton", command=self.novo_preset).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="🗑️ Excluir", style="Secondary.TButton", command=self.excluir_preset).pack(side=tk.LEFT, padx=5)

    def atualizar_lista(self):
        self.listbox.delete(0, tk.END)
        for p in self.logic.presets:
            self.listbox.insert(tk.END, f" {p['nome']}")

    def novo_preset(self):
        nome = simpledialog.askstring("Novo Preset", "Nome:")
        if nome:
            self.logic.presets.append({"nome": nome, "colunas_deletar": "", "tipo": "limpeza"})
            self.logic.salvar_presets(self.logic.presets)
            self.atualizar_lista()

    def excluir_preset(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        if messagebox.askyesno("Confirmar", "Excluir preset?"):
            del self.logic.presets[idx]
            self.logic.salvar_presets(self.logic.presets)
            self.atualizar_lista()
