[README.md](https://github.com/user-attachments/files/24742250/README.md)
# 🧹 ADC - Limpador de Planilha MVSKU

Aplicação profissional com Interface Gráfica (GUI) para limpeza e filtragem de planilhas de produtos (Itens Mais Vendidos por SKU).

## 🚀 Funcionalidades

- **Limpeza Automática**: Remove colunas desnecessárias com base em índices configuráveis.
- **Filtros Avançados**:
    - 🔄 Remoção de linhas duplicadas.
    - 🗑️ Remoção de linhas vazias ou incompletas.
    - 📊 Filtragem por valor mínimo (vendas/quantidade).
    - 🔍 Filtragem por texto parcial (SKU/Categoria).
- **Interface Intuitiva**: Seleção visual de arquivos e log de processamento em tempo real.
- **Segurança**: Criação automática de backup do arquivo original antes do processamento.

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
- `scripts/atualizar_executavel.bat`: Gera uma nova versão do executável na pasta `dist/`.

---
Desenvolvido para otimizar o fluxo de trabalho com planilhas de SKU.
