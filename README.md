# 📊 Delta Analytics - Engenharia

![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![PySide6](https://img.shields.io/badge/UI-PySide6-green.svg)
![Pandas](https://img.shields.io/badge/Data-Pandas-orange.svg)
![Status](https://img.shields.io/badge/Status-Em%20Desenvolvimento-yellow.svg)

O **Delta Analytics** é uma ferramenta desktop desenvolvida para otimizar o fluxo de trabalho da engenharia de processos. Ele automatiza a conversão de listas técnicas do SAP em listas de corte operacionais, gerencia especificações de componentes (cabos, terminais, selos) e calcula tempos de setup e produção (Rates).

## 🚀 Funcionalidades Principais

*   **Conversão SAP:** Transforma arquivos brutos do SAP em listas de circuitos amigáveis para a produção.
*   **Processamento Inteligente:**
    *   Definição automática de processos.
    *   Sequenciamento de circuitos.
    *   Cálculo de volumes e comunização.
*   **Gestão de Bases de Dados:** Interface para visualização e atualização de tabelas JSON (Cabos, Terminais, Selos, Máquinas, Rates e Setup).
*   **Monitoramento de Sistema:** Painel lateral com informações do colaborador, latência de rede, status da máquina e consumo de memória.
*   **Logs de Erro:** Sistema robusto de captura de exceções com interface dedicada para depuração técnica.

## 🛠️ Tecnologias Utilizadas

*   **Linguagem:** [Python 3](https://www.python.org/)
*   **Interface Gráfica (GUI):** [PySide6](https://doc.qt.io/qtforpython/) (Qt for Python)
*   **Manipulação de Dados:** [Pandas](https://pandas.pydata.org/)
*   **Monitoramento de Hardware:** [psutil](https://psutil.readthedocs.io/)
*   **Integração Windows:** `win32com.client` (para coleta de dados do Active Directory/Outlook)
*   **Gerenciador de Pacotes:** [uv](https://github.com/astral-sh/uv) (identificado pelo arquivo `uv.lock`)

## 📁 Estrutura do Projeto

```text
DELTA_ANALYTICS/
├── core/               # Regras de negócio centrais
├── data/               # Arquivos JSON (Bases de dados locais)
├── infra/              # Configurações e persistência
├── services/           # Lógica de integração e serviços externos
├── ui/                 # Arquivos de definição de interface
├── utils/              # Funções auxiliares (conversores, cálculos, info_maq)
├── main.py             # Ponto de entrada da aplicação
├── pyproject.toml      # Configurações de dependências
└── uv.lock             # Lockfile do gerenciador uv
```

## ⚙️ Instalação e Execução

Como o projeto utiliza o gerenciador **uv**, a execução é extremamente rápida:

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/seu-usuario/delta-analytics.git
    cd delta-analytics
    ```

2.  **Instale as dependências:**
    ```bash
    # Se você tem o uv instalado:
    uv sync
    ```

3.  **Execute a aplicação:**
    ```bash
    uv run main.py
    ```

## 📖 Como Usar

1.  **Configuração Inicial:** Na aba "Basic Information", verifique se as tabelas bases (ícones ✔️/❌) estão atualizadas. Caso não estejam, importe os arquivos CSV correspondentes.
2.  **Importação:** Vá até a aba "Lista de Circuitos - Corte" e clique em **Tool list** para carregar o Excel exportado do SAP.
3.  **Processamento:**
    *   Clique em **Converte para lista**.
    *   Utilize os checkboxes para aplicar as camadas de inteligência: **Definir Processos**, **Sequenciar** e **Volumes**.
4.  **Análise:** Use a barra de filtros para pesquisar circuitos específicos por coluna.
5.  **Logs:** Se algo não funcionar como esperado, a aba **Logs Erros** detalhará exatamente o que aconteceu para facilitar o suporte.

## 👤 Autor

*   **Desenvolvedor:** TMGods
*   **Ano:** 2026
*   **Versão:** 1.0

---
*Este projeto é de uso restrito ao departamento de Engenharia.*