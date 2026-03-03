# 🧹 Excel Filter

Ferramenta desktop para filtrar e limpar arquivos Excel de forma simples e rápida, sem precisar de conhecimento técnico.

## ✨ Funcionalidades

- Abre qualquer arquivo `.xlsx` e lê as colunas automaticamente
- Permite adicionar múltiplos filtros por coluna
- Remove linhas com base nos valores informados
- Suporta colar vários valores de uma vez (vírgula, espaço ou um por linha)
- Salva o resultado em um novo arquivo Excel
- Interface gráfica intuitiva com tema escuro
- Zera os campos automaticamente após cada operação

## 🖥️ Como usar

1. Clique em **PROCURAR** e selecione o arquivo Excel
2. Clique em **+ ADICIONAR FILTRO**
3. Escolha a coluna no dropdown
4. Cole ou digite os valores que deseja remover
5. Repita os passos 2-4 para adicionar mais filtros se necessário
6. Clique em **FILTRAR E SALVAR** e escolha onde salvar o resultado

## 🚀 Como rodar o projeto

### Pré-requisitos

- Python 3.10+
- pip

### Instalação

```bash
# Clone o repositório
git clone https://github.com/MarroneDelta/limparExel.git
cd limparExel

# Crie o ambiente virtual
python -m venv .venv

# Ative o ambiente virtual
# Windows:
.venv\Scripts\activate.bat

# Instale as dependências
pip install -r requirements.txt

# Rode o programa
python limpa.py
```

## 📦 Gerar executável (.exe)

```bash
pyinstaller --onefile --windowed --name "ExcelFilter" limpa.py
```

O `.exe` será gerado na pasta `dist/`.

## 🛠️ Tecnologias

- [Python](https://www.python.org/)
- [Pandas](https://pandas.pydata.org/) — leitura e manipulação de dados
- [OpenPyXL](https://openpyxl.readthedocs.io/) — leitura e escrita de arquivos Excel
- [Tkinter](https://docs.python.org/3/library/tkinter.html) — interface gráfica

## 📄 Licença

MIT License — fique à vontade para usar e modificar.
