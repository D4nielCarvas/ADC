"""
Script para limpeza de planilha de itens mais vendidos por SKU
Versão com Interface Gráfica (GUI) usando tkinter
"""

import pandas as pd
import os
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
        self.root.title("Limpador de Planilha - Itens Mais Vendidos")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        
        # Variáveis
        self.caminho_entrada = tk.StringVar()
        self.caminho_saida = tk.StringVar()
        self.processando = False
        
        self.criar_interface()
        
    def criar_interface(self):
        """Criar todos os elementos da interface gráfica"""
        
        # Frame principal com padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        titulo = ttk.Label(
            main_frame, 
            text="🧹 Limpador de Planilha - Itens Mais Vendidos por SKU",
            font=("Arial", 14, "bold")
        )
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Seção 1: Arquivo de Entrada
        ttk.Label(
            main_frame, 
            text="1️⃣ Arquivo de Entrada (.xls):",
            font=("Arial", 10, "bold")
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        
        entrada_frame = ttk.Frame(main_frame)
        entrada_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.entrada_entry = ttk.Entry(
            entrada_frame, 
            textvariable=self.caminho_entrada, 
            width=60,
            state="readonly"
        )
        self.entrada_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_entrada = ttk.Button(
            entrada_frame, 
            text="📁 Procurar", 
            command=self.selecionar_arquivo_entrada
        )
        btn_entrada.pack(side=tk.LEFT, padx=(5, 0))
        
        # Seção 2: Arquivo de Saída
        ttk.Label(
            main_frame, 
            text="2️⃣ Salvar Como (.xlsx):",
            font=("Arial", 10, "bold")
        ).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        
        saida_frame = ttk.Frame(main_frame)
        saida_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.saida_entry = ttk.Entry(
            saida_frame, 
            textvariable=self.caminho_saida, 
            width=60,
            state="readonly"
        )
        self.saida_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_saida = ttk.Button(
            saida_frame, 
            text="💾 Procurar", 
            command=self.selecionar_arquivo_saida
        )
        btn_saida.pack(side=tk.LEFT, padx=(5, 0))
        
        # Seção 3: Configuração
        ttk.Label(
            main_frame, 
            text="3️⃣ Colunas a Deletar (índices):",
            font=("Arial", 10, "bold")
        ).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        
        config_frame = ttk.Frame(main_frame)
        config_frame.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        self.indices_entry = ttk.Entry(config_frame, width=30)
        self.indices_entry.pack(side=tk.LEFT)
        self.indices_entry.insert(0, "1, 3, 4, 5")  # Valor padrão
        
        ttk.Label(
            config_frame, 
            text="  (Ex: 1, 3, 4, 5)",
            foreground="gray"
        ).pack(side=tk.LEFT)
        
        # Botão Processar
        self.btn_processar = ttk.Button(
            main_frame, 
            text="▶️ PROCESSAR PLANILHA", 
            command=self.iniciar_processamento,
            style="Accent.TButton"
        )
        self.btn_processar.grid(row=7, column=0, columnspan=3, pady=20)
        
        # Área de Log
        ttk.Label(
            main_frame, 
            text="📋 Log de Processamento:",
            font=("Arial", 10, "bold")
        ).grid(row=8, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        
        self.log_text = scrolledtext.ScrolledText(
            main_frame, 
            width=75, 
            height=15,
            font=("Consolas", 9),
            wrap=tk.WORD
        )
        self.log_text.grid(row=9, column=0, columnspan=3, pady=(0, 10))
        
        # Barra de status
        self.status_label = ttk.Label(
            main_frame, 
            text="✨ Pronto para processar",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
    def selecionar_arquivo_entrada(self):
        """Abrir diálogo para selecionar arquivo de entrada"""
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha de entrada",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~/Downloads")
        )
        
        if arquivo:
            self.caminho_entrada.set(arquivo)
            
            # Sugerir nome para arquivo de saída automaticamente
            if not self.caminho_saida.get():
                base, ext = os.path.splitext(arquivo)
                sugestao = f"{base}_limpo.xlsx"
                self.caminho_saida.set(sugestao)
            
            self.log(f"✓ Arquivo selecionado: {os.path.basename(arquivo)}")
    
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
        """Adicionar mensagem ao log visual"""
        self.log_text.insert(tk.END, mensagem + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
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
            
            # Processar índices
            try:
                indices_texto = self.indices_entry.get()
                index_delet = [int(i.strip()) for i in indices_texto.split(",")]
                self.log(f"✓ Índices de colunas: {index_delet}")
            except ValueError:
                raise ValueError("Formato de índices inválido! Use números separados por vírgula.")
            
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
            df_limpo = self.deletar_colunas(df, index_delet)
            
            # Passo 6: Salvar
            self.log("\n6️⃣ Salvando planilha limpa...")
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
        """Carregar planilha do Excel"""
        try:
            df = pd.read_excel(caminho)
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
        """Deletar colunas especificadas"""
        try:
            colunas_deletar = [df.columns[i] for i in indices]
            self.log(f"✓ Deletando colunas: {colunas_deletar}")
            
            df_limpo = df.drop(df.columns[indices], axis=1)
            
            self.log(f"✓ Colunas removidas com sucesso")
            self.log(f"  Colunas restantes: {len(df_limpo.columns)}")
            
            return df_limpo
        except Exception as e:
            raise Exception(f"Erro ao deletar colunas: {e}")
    
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