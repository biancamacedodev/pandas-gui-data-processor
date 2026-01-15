# 📊 Pandas GUI Data Processor

Aplicação desktop desenvolvida em **Python** para manipulação de planilhas Excel por meio de uma **interface gráfica**, utilizando **Pandas** e **Tkinter**.

O projeto foi criado com foco em **aprendizado prático**, automação de tarefas repetitivas e manipulação de dados sem a necessidade de escrever código diretamente.

> **📝 Nota:** Este projeto foi originalmente desenvolvido em **Jupyter Notebook** e posteriormente convertido para uma aplicação executável standalone, mantendo todas as funcionalidades e a mesma interface gráfica.

---

## 🚀 Sobre o Projeto

Este sistema permite abrir arquivos Excel, visualizar os dados em tabela e realizar diversas operações comuns de análise e tratamento de dados através de menus e janelas interativas.

Todas as ações são processadas com **Pandas**, enquanto a interface é construída com **Tkinter**, tornando a aplicação leve e fácil de executar.

---

## 🧠 O que eu aprendi com este projeto

- Manipulação de dados com **Pandas**
- Leitura e escrita de arquivos Excel
- Criação de interfaces gráficas com **Tkinter**
- Uso de **DataFrames** em aplicações desktop
- Agrupamento, filtros, merges e limpeza de dados
- Organização de código orientado a objetos em Python
- Conversão de projetos Jupyter Notebook para aplicações executáveis

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.13+**
- **Pandas** - Manipulação e análise de dados
- **NumPy** - Operações numéricas
- **Tkinter** - Interface gráfica (built-in Python)
- **pandastable** - Componente de tabela interativa
- **openpyxl** - Leitura e escrita de arquivos Excel

---

## 📋 Funcionalidades Principais

### 📁 Arquivos
- ✅ Abrir arquivos Excel (`.xlsx` e `.xls`)
- ✅ Salvar arquivos tratados em Excel

### ✏️ Manipulação de Dados
- ✅ Renomear colunas
- ✅ Remover colunas
- ✅ Filtrar dados por valor
- ✅ Remover linhas em branco
- ✅ Remover linhas duplicadas
- ✅ Remover intervalos de linhas

### 📊 Análise
- ✅ Agrupar dados por coluna e somar valores numéricos
- ✅ Cálculo automático da soma de colunas numéricas

### 🔗 Merges de Arquivos
- ✅ Inner Join
- ✅ Left Join
- ✅ Outer Join
- ✅ Join Full (concatenação)

### 📂 Relatórios
- ✅ Consolidar vários arquivos Excel de uma pasta
- ✅ Quebrar um arquivo em vários relatórios com base em uma coluna

### ✍️ Edição Manual
- ✅ Edição direta dos dados ao clicar duas vezes em uma linha

---

## ▶️ Como Executar o Projeto

### Pré-requisitos

- Python 3.13 ou superior
- pip (gerenciador de pacotes Python)

### Instalação

1. Clone o repositório:
```bash
git clone https://github.com/biancamacedodev/pandas-gui-data-processor.git
cd pandas-gui-data-processor
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

### Executar a Aplicação

Execute o arquivo principal:
```bash
python main.py
```

A interface gráfica será aberta automaticamente.

---

## 📦 Gerar Executável (Opcional)

Para gerar um executável `.exe` usando PyInstaller:

1. Instale o PyInstaller:
```bash
pip install pyinstaller
```

2. Gere o executável:
```bash
pyinstaller --onefile --noconsole --name "ExcelEditor" main.py
```

O arquivo executável estará na pasta `dist/`.


---

## 🎯 Casos de Uso

- **Análise de dados:** Processar planilhas Excel sem conhecimento avançado de programação
- **Limpeza de dados:** Remover duplicatas, linhas vazias e dados inconsistentes
- **Consolidação:** Unir múltiplos arquivos Excel em um único relatório
- **Divisão de dados:** Separar um arquivo grande em múltiplos arquivos menores
- **Transformação:** Renomear colunas, filtrar dados e realizar agrupamentos

