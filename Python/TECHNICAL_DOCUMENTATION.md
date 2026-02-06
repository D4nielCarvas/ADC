# 📘 Documentação Técnica - ADC (Advanced Data Cleaner) v3.0

> **Versão:** 3.0 (Estável)  
> **Data:** 06/02/2026  
> **Desenvolvedor:** DanielCarvas (Mantido por Antigravity)

---

## 🚀 Visão Geral do Projeto

O **ADC (Advanced Data Cleaner)** é uma aplicação de desktop profissional desenvolvida para automação de limpeza, padronização e análise de planilhas de e-commerce (foco em Shopee/Marketplaces). O sistema remove tarefas manuais repetitivas, garantindo integridade de dados e fornecendo insights rápidos através de dashboards.

### 🎯 Principais Capacidades
1.  **Limpeza Inteligente**: Remove colunas inúteis, linhas vazias e duplicatas automaticamente baseada em presets.
2.  **Dashboard Financeiro**: Analisa múltiplos arquivos simultaneamente para calcular itens vendidos, pedidos únicos e receita total.
3.  **Compatibilidade Universal**: Sistema robusto de carregamento que suporta arquivos Excel modernos (`.xlsx`) e legados (`.xls`), mesmo com corrupções leves.
4.  **Interface Moderna**: UI baseada em Tkinter com tema personalizado (Catppuccin), responsiva e com feedback visual em tempo real.

---

## 🏗️ Arquitetura de Software

O projeto segue uma arquitetura modular inspirada no padrão **MVC (Model-View-Controller)**, separando rigidamente a lógica de negócios da interface gráfica.

### Diagrama de Camadas
```mermaid
graph TD
    A[GUI Layer (View)] --> B[Controller/Pages]
    B --> C[Core Logic (Model)]
    C --> D[Data Persistence (JSON/Excel)]
    
    style A fill:#f9f,stroke:#333
    style C fill:#bbf,stroke:#333
```

1.  **Core Logic (`src/core`)**: Contém toda a inteligência do negócio (`ADCLogic`). Não depende de nenhuma biblioteca gráfica, permitindo fácil portabilidade ou uso via CLI/API.
2.  **GUI (`src/gui`)**: Implementação visual usando `tkinter`. Gerencia eventos, threads e atualização de widgets.
3.  **Config (`config/`)**: Persistência de preferências do usuário e presets de limpeza.

---

## 📂 Estrutura de Diretórios

```
C:\Projetos\Codigos\Python\
├── config/                  # Arquivos de configuração (gerado automaticamente)
│   └── settings.json        # Presets de limpeza e preferências
├── scripts/                 # Scripts utilitários de build/manutenção
│   └── atualizar_executavel.bat
├── src/                     # Código fonte principal
│   ├── core/                # Camada de Regra de Negócios
│   │   └── cleaner.py       # CLASSE PRINCIPAL: ADCLogic
│   ├── gui/                 # Camada de Interface Gráfica
│   │   ├── assets/          # Ícones e Imagens (.ico, .png)
│   │   ├── pages/           # Módulos das Telas (Componentes)
│   │   │   ├── cleaner.py   # Lógica da tela de Limpeza
│   │   │   ├── dashboard.py # Lógica do Dashboard Financeiro
│   │   │   ├── config.py    # Tela de Configuração e Presets
│   │   │   └── home.py      # Tela inicial
│   │   ├── styles.py        # Definição de Temas e Cores
│   │   └── main_window.py   # Janela Principal (Container de Navegação)
│   └── main.py              # Ponto de Entrada (Entry Point)
├── ADC.spec                 # Arquivo de especificação PyInstaller
└── requirements.txt         # Dependências do Python
```

---

## 🧠 Núcleo Lógico (`src/core/cleaner.py`)

A classe `ADCLogic` é o coração do sistema.

### 1. Sistema de Carregamento Híbrido (`carregar_planilha`)
Implementa uma estratégia de "Fallback em Tripla Camada" para garantir que o usuário consiga abrir qualquer planilha:
1.  **Detecção de Extensão**: Escolhe o engine primário (`openpyxl` para `.xlsx`, `xlrd` para `.xls`).
2.  **Tentativa Primária**: Tenta carregar com o engine ideal.
3.  **Fallback Secundário**: Se falhar (devido a corrupção ou formato incorreto), tenta o engine alternativo.
4.  **Fallback Automático**: Deixa o Pandas decidir o engine.
5.  **Auto-Load**: Se nenhuma aba for especificada (`aba=""`), identifica e carrega automaticamente a primeira aba disponível iterando por todos os engines.

### 2. Pipeline de Limpeza (`processar_limpeza`)
Fluxo linear e determinístico:
1.  **Validação**: Verifica existência do arquivo e integridade.
2.  **Load**: Carrega DataFrame.
3.  **Drop Columns**: Remove colunas por índice (mapeado da interface 1-based para 0-based).
4.  **Filtros**:
    -   `remover_duplicadas`: `df.drop_duplicates()`
    -   `remover_vazias`: `df.dropna(how='all')`
    -   `filtro_valor`: Normaliza strings de moeda ("R$ 1.200,50") para float e filtra.
    -   `filtro_texto`: Busca case-insensitive em todas as colunas de texto.

### 3. Motor de Cálculo de Dashboard (`gerar_resumo`)
O sistema de cálculo financeiro foi padronizado para planilhas de vendas (Shopee):

| Métrica | Fonte de Dados | Lógica |
| :--- | :--- | :--- |
| **Total de Pedidos** | Coluna B (Índice 1) | Contagem de valores únicos (`nunique`) para evitar duplicatas de itens no mesmo pedido. |
| **Total de Itens** | Coluna Z (Índice 25) | Soma simples dos valores numéricos da coluna. |
| **Valor Total** | Coluna Z * Coluna AA | Multiplica **Quantidade (Z)** por **Preço Unitário (AA)** linha a linha e soma o resultado. |

> **Nota:** O sistema sanitiza dados numéricos (remove "R$", pontos e vírgulas) antes de qualquer cálculo matemático.

---

## 🖥️ Interface Gráfica (`src/gui`)

### Gerenciamento de Estado e Threads
Para manter a interface responsiva durante processamento pesado (ex: ler 10 arquivos Excel):
-   **Threading**: O processamento ocorre em uma `threading.Thread` separada (daemon=True).
-   **Safe UI Updates**: A atualização da UI (Labels, ProgressBars) é feita via `root.after()` ou através de um sistema de callbacks seguro, evitando *"RuntimeError: main thread is not in main loop"*.

### Dashboard Multiarquivo (`dashboard.py`)
Recurso avançado recém-implementado:
-   **Input**: Aceita N arquivos simultâneos.
-   **Lógica de Combinação**:
    -   **Pedidos**: Mantém um `Set` global de IDs de pedidos para garantir que o mesmo pedido em arquivos diferentes não seja contado duas vezes.
    -   **Soma**: Acumula `total_itens` e `valor_total` de cada arquivo processado.
    -   **Tratamento de Erro Individual**: Se 1 de 10 arquivos falhar, o sistema processa os outros 9 e relata o erro específico apenas do arquivo problemático.

---

## 🛠️ Tecnologias e Dependências

| Componente | Tecnologia | Versão Mínima | Uso |
| :--- | :--- | :--- | :--- |
| **Runtime** | Python | 3.10+ | Linguagem base |
| **Data Engine** | Pandas | 2.0+ | Manipulação de dados |
| **Excel (Modern)** | Openpyxl | 3.1+ | Leitura/Escrita .xlsx |
| **Excel (Legacy)** | Xlrd | 2.0.1 | Leitura .xls |
| **GUI** | Tkinter | (Built-in) | Interface Gráfica |
| **Plots** | Matplotlib | 3.7+ | (Opcional) Gráficos futuros |
| **Build** | PyInstaller | 6.0+ | Compilação para .exe |

---

## 🔧 Guia de Manutenção e Build

### Como rodar em desenvolvimento
```powershell
python src/main.py
```

### Como gerar novo executável
Utilize o script automatizado que limpa arquivos temporários, constrói e organiza a pasta `dist`:
```powershell
scripts/atualizar_executavel.bat
```
O executável final estará em: `dist/ADC/ADC.exe`

---

> Documentação gerada automaticamente por **Antigravity**.
