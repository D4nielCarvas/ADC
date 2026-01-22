# Documentação Técnica - ADC (Advanced Data Cleaner) v2.5 Pro

Este documento fornece uma visão técnica detalhada da arquitetura e do funcionamento interno do sistema ADC.

## 🏗️ Arquitetura Geral
O sistema é construído inteiramente em **Python 3.8+**, utilizando uma arquitetura modular. A interface é baseada em **Tkinter** com um "wrapper" de estilização moderna via `ttk`.

### Principais Bibliotecas
- **Pandas**: Núcleo de processamento e manipulação de DataFrames.
- **Tkinter**: Interface gráfica e gerenciamento de eventos.
- **Matplotlib**: Geração de gráficos e dashboards.
- **Openpyxl/xlrd**: Motores de leitura/escrita de arquivos Excel.
- **Threading**: Utilizado para manter a interface fluida durante o processamento pesado.

---

## 📱 Estrutura de Módulos (Refatorada)

O projeto foi refatorado para separar responsabilidades:

- **`src/main.py`**: Ponto de entrada leve. Apenas inicializa a `MainWindow`.
- **`src/core/cleaner.py` (`ADCLogic`)**:  Centraliza toda a regra de negócios, incluindo validação, carregamento, limpeza e filtros. Garante reutilização entre GUI e outros possíveis frontends.
- **`src/gui/`**: Contém toda a lógica de interface.
    - **`main_window.py`**: Controlador principal, gerencia a Sidebar e a troca de páginas.
    - **`styles.py`**: Definições de temas (Cores Catppuccin) e estilos TTK.
    - **`pages/`**: Módulos independentes para cada tela (`cleaner.py`, `dashboard.py`, `config.py`).

### 1. Inicialização e Estado
- `MainWindow` instancia `ADCLogic` uma única vez e a injeta nas páginas.
- Isso garante que o estado dos Presets e Caches seja compartilhado.

### 2. Sistema de Temas e UI
- Implementa um tema "Modern Dark" customizado em `styles.py`.
- Utiliza uma **Sidebar** para navegação e um sistema de **Containers** (Frames) que são alternados via o método `mudar_pagina`, criando o efeito de multi-páginas.

---

## ⚙️ Lógica de Processamento de Dados

### Carregamento Robusto (`carregar_planilha`)
Implementa um sistema de **Fallback Automático**:
1. Tenta ler usando o motor preferencial (`openpyxl` para `.xlsx`, `xlrd` para `.xls`).
2. Se falhar (devido a corrupção de cabeçalho ou formato não padrão), tenta o motor alternativo.

### O Ciclo de Limpeza (`ADCLogic.processar_limpeza`)
O processamento segue um pipeline linear dentro da classe lógica:
1. **Validação**: Verifica se o arquivo existe e se os índices de colunas solicitados são válidos no DataFrame atual.
2. **Deleção**: Remove as colunas baseadas nos índices (convertendo de base 1 para base 0).
3. **Filtros Adicionais**:
   - `dropna`: Remove linhas vazias e duplicadas.
   - **Filtro de Valor**: Utiliza a função `limpar_valor` para converter strings financeiras ("R$ 1.200,00") em floats comparáveis.
   - **Filtro de Texto**: Aplica busca vetorizada `str.contains(case=False)` em todas as colunas do tipo objeto.

---

## 📊 Dashboard e Visualização

### Integração Matplotlib-Tkinter (`DashboardPage`)
- A geração de resumos é feita em thread separada para não travar a UI.
- Os dados são processados em `ADCLogic.gerar_resumo` e retornados para a GUI apenas para exibição.

---

## 💾 Persistência e Configurações
- **`config/settings.json`**: Armazena os presets. O sistema utiliza `json.dump` e `json.load` para garantir que as regras de limpeza sejam salvas permanentemente.
- **Presets**: Estrutura flexível que define quais colunas devem ser deletadas e quais filtros devem ser ativados por padrão.

---

## 🚀 Concorrência e UX
- **Threading**: Operações de I/O (leitura de Excel) e processamento pesado são sempre executadas em threads.
- **Thread Safety**: Callbacks de atualização de UI (`log_callback`, `set_progress`) usam `root.after` ou métodos seguros do Tkinter.

---
**Desenvolvido por D4nielCarvas**
