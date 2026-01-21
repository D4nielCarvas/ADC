# Documentação Técnica - ADC (Advanced Data Cleaner) v2.5 Pro

Este documento fornece uma visão técnica detalhada da arquitetura e do funcionamento interno do sistema ADC.

## 🏗️ Arquitetura Geral
O sistema é construído inteiramente em **Python 3.8+**, utilizando uma abordagem de Programação Orientada a Objetos (POO). A interface é baseada em **Tkinter** com um "wrapper" de estilização moderna via `ttk`.

### Principais Bibliotecas
- **Pandas**: Núcleo de processamento e manipulação de DataFrames.
- **Tkinter**: Interface gráfica e gerenciamento de eventos.
- **Matplotlib**: Geração de gráficos e dashboards.
- **Openpyxl/xlrd**: Motores de leitura/escrita de arquivos Excel.
- **Threading**: Utilizado para manter a interface fluida durante o processamento pesado.

---

## 📱 A Classe Principal: `LimpadorPlanilhaGUI`
Localizada em `src/main.py`, esta classe gerencia todo o ciclo de vida da aplicação.

### 1. Inicialização e Estado (`__init__`)
- Configura a janela raiz, variáveis de estado (`tk.StringVar`, `tk.BooleanVar`) e o cache dinâmico de arquivos Excel para evitar leituras repetitivas do disco.

### 2. Sistema de Temas e UI (`configurar_estilos` & `criar_interface`)
- Implementa um tema "Modern Dark" customizado.
- Utiliza uma **Sidebar** para navegação e um sistema de **Containers** (Frames) que são alternados via o método `mudar_pagina`, criando o efeito de multi-páginas.

---

## ⚙️ Lógica de Processamento de Dados

### Carregamento Robusto (`carregar_planilha`)
Implementa um sistema de **Fallback Automático**:
1. Tenta ler usando o motor preferencial (`openpyxl` para `.xlsx`, `xlrd` para `.xls`).
2. Se falhar (devido a corrupção de cabeçalho ou formato não padrão), tenta o motor alternativo.
3. Para arquivos `.xls`, utiliza a flag `ignore_workbook_corruption=True`.

### O Ciclo de Limpeza (`processar_planilha`)
O processamento segue um pipeline linear:
1. **Validação**: Verifica se o arquivo existe e se os índices de colunas solicitados são válidos no DataFrame atual.
2. **Deleção**: Remove as colunas baseadas nos índices (convertendo de base 1 para base 0).
3. **Filtros Adicionais**:
   - `dropna`: Remove linhas vazias com um limite de 50% de preenchimento.
   - `drop_duplicates`: Elimina linhas idênticas.
   - **Filtro de Valor**: Utiliza a função `limpar_valor` para converter strings financeiras ("R$ 1.200,00") em floats comparáveis.
   - **Filtro de Texto**: Aplica busca vetorizada `str.contains(case=False)` em todas as colunas do tipo objeto.

---

## 📊 Dashboard e Visualização

### Integração Matplotlib-Tkinter (`exibir_dashboard`)
- Cria figuras do Matplotlib (`plt.subplots`) com o fundo sincronizado ao tema escuro da GUI.
- **Ranking Inteligente**:
  - Na aba **Resumo**, o ranking é calculado por **frequência** (`groupby().size()`), ideal para ver recorrência de pedidos.
  - No fallback, o ranking é por **volume** (`groupby().sum()`).
- **Modo Cinema**: O método `expandir_dashboard` transfere a referência da figura para uma nova janela `Toplevel` em tela cheia.

---

## 💾 Persistência e Configurações
- **`config.json`**: Armazena os presets. O sistema utiliza `json.dump` e `json.load` para garantir que as regras de limpeza sejam salvas permanentemente.
- **Presets**: Estrutura flexível que define quais colunas devem ser deletadas e quais filtros devem ser ativados por padrão.

---

## 🚀 Concorrência e UX
- **Threading**: O método `iniciar_processamento` lança a lógica pesada em uma thread separada.
- **Thread Safety**: Todas as atualizações de interface de dentro da thread (Logs, Barra de Progresso) são enviadas via `self.root.after()` para garantir que o Tkinter não trave ou apresente comportamentos erráticos.

---
**Desenvolvido por D4nielCarvas**
