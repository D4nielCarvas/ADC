# 🧹 ADC - Advanced Data Cleaner

Aplicação profissional com Interface Gráfica (GUI) para limpeza, filtragem e análise de planilhas de produtos (**Itens Mais Vendidos por SKU**).

![Versão](https://img.shields.io/badge/Vers%C3%A3o-2.5%20Pro-blue)
![Python](https://img.shields.io/badge/Python-3.8%2B-green)

## 🚀 Funcionalidades Principais

- **🧹 Limpeza Inteligente**: Remova colunas desnecessárias, normalize dados e elimine duplicatas em segundos.
- **📊 Dashboard de SKUs**: No Modo Resumo, visualize instantaneamente os **10 SKUs Mais Pedidos** por frequência de ocorrência.
- **🔍 Modo Cinema**: Expanda qualquer gráfico para tela cheia com um clique para análise detalhada.
- **⚙️ Editor de Presets**: Crie e salve perfis de limpeza personalizados para diferentes tipos de relatórios via interface gráfica.
- **📁 Suporte Universal**: Leitura robusta de arquivos `.xlsx` e `.xls` com tratamento de números BR/US (vírgula/ponto).
- **✨ Interface Premium**: Design moderno em Modo Escuro com **Logs Isolados** por aba e suporte a Tela Cheia (F11).
- **🗂️ Indexação Humana**: Sistema de colunas seguindo o padrão Excel (Coluna A = 1, Coluna B = 2).

## 📁 Estrutura do Projeto

```text
Projeto/
├── config/                  # Configurações do sistema
│   └── settings.json
├── src/                     # Código fonte
│   ├── main.py              # Ponto de entrada
│   ├── core/                # Lógica de negócios (ADCLogic)
│   ├── gui/                 # Interface Gráfica
│   │   ├── components/      # Componentes reutilizáveis
│   │   ├── pages/           # Páginas (Limpeza, Dashboard, Config)
│   │   ├── main_window.py   # Janela Principal
│   │   └── styles.py        # Temas e Estilos
│   └── utils/               # Utilitários
├── build/                   # Arquivos de build (PyInstaller)
├── dist/                    # Executável final
├── scripts/                 # Scripts auxiliares
└── requirements.txt         # Dependências
```

## 🛠️ Requisitos e Instalação

### Pré-requisitos
- Python 3.8 ou superior.

### Instalação das Dependências
```bash
pip install -r requirements.txt
```

## 💻 Como Usar

### 1. Execução via Python
Basta rodar o script principal localizado na pasta `src/`:
```bash
python src/main.py
```

### 2. Modos de Operação
- **Modo Limpeza**: Selecione o arquivo, o destino e o preset desejado. Clique em "Iniciar Limpeza" para gerar a nova planilha.
- **Modo Resumo**: Analise métricas financeiras e de volume sem a necessidade de criar arquivos intermediários.
- **Configurações**: Gerencie seus presets, adicionando ou removendo regras de colunas para deletar.

## 🔨 Desenvolvimento e Build

Para gerar uma versão executável (`.exe`) para Windows:
1. Certifique-se de ter o `pyinstaller` instalado.
2. Utilize o script em `scripts/atualizar_executavel.bat` ou execute:
```bash
pyinstaller main.spec
```

---
Desenvolvido por **D4nielCarvas** para otimização de fluxos de E-commerce.
