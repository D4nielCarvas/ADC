import tkinter as tk
from tkinter import ttk

class ThemeConfig:
    @staticmethod
    def get_colors():
        return {
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

    @staticmethod
    def apply_styles(root):
        style = ttk.Style()
        colors = ThemeConfig.get_colors()

        root.configure(bg=colors["base"])
        style.theme_use('clam')
        
        # --- CONFIGURAÇÃO DE ESTILOS TTK ---
        
        # 1. Frames e Containers
        style.configure("TFrame", background=colors["base"])
        style.configure("Sidebar.TFrame", background=colors["mantle"])
        style.configure("Card.TFrame", background=colors["surface0"], relief="flat", borderwidth=0)
        
        # 2. Labels
        style.configure("TLabel", 
            background=colors["base"], 
            foreground=colors["text"], 
            font=("Segoe UI", 10)
        )
        style.configure("Card.TLabel",
            background=colors["surface0"], 
            foreground=colors["text"],
            font=("Segoe UI", 10)
        )
        
        # Títulos
        style.configure("Title.TLabel", 
            background=colors["mantle"], 
            foreground=colors["mauve"], 
            font=("Segoe UI", 22, "bold")
        )
        style.configure("Header.TLabel", 
            background=colors["surface0"], 
            foreground=colors["blue"], 
            font=("Segoe UI", 12, "bold")
        )
        style.configure("Sub.TLabel",
            background=colors["base"],
            foreground=colors["subtext"],
            font=("Segoe UI", 9)
        )

        # 3. Botões de Navegação (Sidebar)
        style.configure("Nav.TButton", 
            padding=(20, 12), 
            font=("Segoe UI", 11),
            background=colors["mantle"],
            foreground=colors["subtext"], 
            anchor="w",
            borderwidth=0
        )
        style.map("Nav.TButton",
            background=[('active', colors["surface0"]), ('pressed', colors["surface1"])],
            foreground=[('active', colors["mauve"])]
        )

        # 4. Botões de Ação (Accent)
        style.configure("Accent.TButton", 
            padding=(15, 12), 
            font=("Segoe UI", 11, "bold"),
            background=colors["mauve"],
            foreground=colors["crust"], 
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
            background=colors["surface1"],
            foreground=colors["text"],
            borderwidth=0
        )
        style.map("Secondary.TButton",
            background=[('active', '#585b70')]
        )
        
        # 6. Inputs e Widgets
        style.configure("TEntry", 
            fieldbackground=colors["crust"], 
            foreground=colors["text"],
            insertcolor=colors["text"],
            borderwidth=0,
            padding=5
        )
        
        style.configure("TCombobox", 
            fieldbackground=colors["crust"], 
            background=colors["surface0"], 
            foreground=colors["text"],
            arrowcolor=colors["mauve"],
            borderwidth=0
        )
        # Hack para Combobox fica no root.option_add abaixo

        # 7. LabelFrames
        style.configure("TLabelframe", 
            background=colors["surface0"], 
            foreground=colors["blue"], 
            font=("Segoe UI", 10, "bold"),
            borderwidth=1,
            relief="solid",
            bordercolor=colors["surface1"]
        )
        style.configure("TLabelframe.Label", background=colors["surface0"], foreground=colors["blue"])
        style.configure("TCheckbutton", background=colors["surface0"], foreground=colors["text"])

        # 8. Stats Cards (Resumo)
        style.configure("Stat.TLabel", 
            background=colors["surface0"], 
            foreground=colors["green"], 
            font=("Segoe UI", 24, "bold")
        )
        style.configure("StatDesc.TLabel", 
            background=colors["surface0"], 
            foreground=colors["subtext"], 
            font=("Segoe UI", 10, "bold")
        )
        
        # Configurações globais opcionais
        root.option_add('*TCombobox*Listbox.background', colors["crust"])
        root.option_add('*TCombobox*Listbox.foreground', colors["text"])
        root.option_add('*TCombobox*Listbox.selectBackground', colors["mauve"])
        root.option_add('*TCombobox*Listbox.selectForeground', colors["crust"])
