app/
|____ui/         # Telas (PySide)
|____core/       # Regras de negócio
|____services/   # IA, sinas, dados
|____infra/      # Banco, arquivos
main.py

# Criar o projeto
uv init meu_projeto
cd meu_projeto

# Adicionar o Ruff
uv add ruff

# (Opcional) Ver dependências instaladas
uv pip list

# Analisar o código
uv run ruff check .

# Corrigir automaticamente problemas
uv run ruff check . --fix

# Formatar o código
uv run ruff format .

# Fluxo diário recomendado
uv run ruff format .
uv run ruff check . --fix

# (Opcional) Adicionar pre-commit
uv add pre-commit
pre-commit install


# Esse comando: cria/atualiza o ambiente virtual/ instala tudo automaticamente
uv sync


uv run pyside6-designer

uv run pyside6-designer ui\simulator.ui
uv run pyside6-uic ui\untitled.ui -o ui\untitled.py

