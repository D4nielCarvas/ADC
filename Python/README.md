[README.md](https://github.com/user-attachments/files/24742250/README.md)
# 🧹 ADC

Aplicação profissional com Interface Gráfica (GUI) para limpeza e filtragem de planilhas de produtos (Itens Mais Vendidos por SKU).

## 🚀 Funcionalidades

- **Multicamadas e Presets**: Suporte dinâmico para diferentes tipos de relatórios via `config.json`.
- **Indexação Humana (A=1)**: Sistema de colunas intuitivo que segue o padrão do Excel (Coluna A = 1, Coluna B = 2).
- **Seleção de Abas**: Escolha qual aba (sheet) do Excel deseja processar diretamente na interface.
- **Controle de Saída Manual**: Total liberdade para escolher o nome e o local de salvamento dos arquivos.
- **Interface Premium**: Design modernizado com Modo Escuro, cards e suporte a Tela Cheia (F11).
- **Filtros Flexíveis**: Começam desmarcados para garantir que você só aplique o que realmente deseja.

## 📁 Estrutura do Projeto

```text
ADC/
├── docs/           # Documentação adicional
├── scripts/        # Scripts de automação (.bat)
├── src/            # Código-fonte Python
│   └── main.py     # Script principal da aplicação
├── tests/          # Scripts de teste e validação
├── README.md       # Documentação principal
└── .gitignore      # Arquivos ignorados pelo Git
```

## 🛠️ Requisitos

- **Python 3.8+**
- Dependências: `pandas`, `openpyxl`, `tkinter` (incluído no Python)

## 💻 Como Usar

### Usando o Executável (Windows)
1. Vá até a pasta `dist/`.
2. Execute o arquivo `LimpadorPlanilha.exe`.

### Executando via Python
1. Instale as dependências:
   ```bash
   pip install pandas openpyxl
   ```
2. Execute o script principal:
   ```bash
   python src/main.py
   ```

## 🔨 Desenvolvimento

Para gerar um novo executável, utilize o script na pasta `scripts/`:
- `scripts/atualizar_executavel.bat`: Gera uma nova versão do executável na pasta `dist/` (inclui o `config.json`).

---
Desenvolvido para otimizar o fluxo de trabalho com planilhas de SKU.
