# CHANGELOG - ADC (Acompanhamento de Distribuição e Carga)

Todas as alterações notáveis neste projeto serão documentadas neste arquivo.

## [2.6 Refactor] - 2026-01-22
### Arquitetura
- **Centralização de Lógica**: Movida toda a regra de negócios para `ADCLogic` em `core_logic.py`.
- **Refatoração Main**: `main.py` agora atua apenas como camada de visualização (View).

### Visual (UI/UX)
- **Tema Catppuccin Mocha**: Nova paleta de cores moderna (Dark Mode suave).
- **Design Flat**: Elementos de interface planos com espaçamento generoso e tipografia Segoe UI.
- **KPI Cards**: Visualização de estatísticas em cartões destacados.

### Adicionado
- **Ranking de SKUs Mais Pedidos**: Novo dashboard na aba de Resumo focado na frequência de pedidos.
- **Modo Cinema**: Botão "EXPANDIR DASHBOARD" para visualização cinematográfica em tela cheia.
- **Logs Isolados**: Buffers de log independentes para as abas de Limpeza e Resumo.
- **Limpeza Numérica Universal**: Algoritmo inteligente que suporta separadores decimais brasileiros (vírgula) e americanos (ponto).
- **Normalização Profunda**: Agrupamento de produtos insensível a maiúsculas/minúsculas e espaços extras.
### Melhorado
- **Simplificação de Interface**: Removido dashboard da aba de limpeza para fluxo mais direto.
- **Performance**: Remoção da funcionalidade de backup automático e otimização de cache.
- **UX**: Geometria de janela consistente ao alternar o modo tela cheia global.

## [2.1 Pro] - 2026-01-15
### Adicionado
- **Arquitetura Multi-página**: Sidebar para navegação entre Limpeza, Resumo e Configurações.
- **Modo Resumo**: Funcionalidade para calcular métricas de pedidos sem gerar novos arquivos.
- **Suporte Legado**: Melhoria no tratamento de arquivos `.xls` corrompidos.
### Melhorado
- **Barra de Progresso**: Feedback visual em tempo real durante o processamento.

## [1.0.0] - 2026-01-05
- Versão inicial com funcionalidades básicas de limpeza de SKU.
