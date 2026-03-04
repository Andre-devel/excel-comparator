# 📊 Excel Comparator

> Compare two Excel files cell by cell, highlight divergences in red and export a full audit report.

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.0-000000?style=flat&logo=flask&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-2.2-150458?style=flat&logo=pandas&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-1D6F42?style=flat)
![AI Generated](https://img.shields.io/badge/generated%20with-Claude%20AI-orange?style=flat&logo=anthropic&logoColor=white)

---

## ✨ Funcionalidades

- Upload de dois arquivos `.xlsx` via interface web
- **Validação estrutural obrigatória** antes de qualquer comparação (colunas, ordem, quantidade de linhas)
- Comparação **linha a linha** ou por **chave primária** configurável
- Suporte a **colunas ignoradas**
- Comparação **literal e exata** — sem tolerâncias, sem normalização de datas, sem ignorar espaços ou capitalização
- Geração de relatório `.xlsx` com:
  - Células divergentes destacadas em **vermelho**
  - Aba **"Resumo"** com estatísticas e listagem detalhada de todas as diferenças
- Suporte a arquivos com até **100.000 linhas**
- Tratamento de erros claro e descritivo

---

## 🖥️ Interface

A interface web exibe, após a comparação:
- Total de linhas comparadas
- Total de divergências encontradas
- Número de linhas afetadas
- Preview das primeiras 200 divergências (linha, coluna, valor em cada arquivo)
- Botão para download do relatório completo

---

## 🚀 Como usar

### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

### 2. Iniciar o servidor

```bash
python app.py
```

### 3. Acessar

Abra o navegador em **http://localhost:5000**

### 4. Comparar

1. Selecione o **Arquivo 1** (referência) e o **Arquivo 2** (para comparar)
2. Opcionalmente, informe:
   - **Chave primária** — coluna que identifica cada linha de forma única (ex: `ID`, `cpf`)
   - **Colunas a ignorar** — separadas por vírgula (ex: `updated_at, log_date`)
3. Clique em **"Comparar arquivos"**
4. Baixe o relatório `.xlsx` gerado

---

## 📁 Estrutura do projeto

```
excel-comparator/
│
├── app.py                  # Servidor Flask + rotas da API
│
├── comparator/             # Lógica de negócio em 4 camadas
│   ├── reader.py           # Leitura segura dos arquivos
│   ├── validator.py        # Validação estrutural
│   ├── comparator.py       # Lógica de comparação
│   └── reporter.py         # Geração do relatório .xlsx
│
├── templates/
│   └── index.html          # Interface web (vanilla JS, sem dependências)
│
├── uploads/                # Arquivos temporários (limpos após uso)
├── outputs/                # Relatórios gerados
└── requirements.txt
```

---

## 🛡️ Validações estruturais

Antes de qualquer comparação, o sistema verifica se os dois arquivos possuem:

| Validação | Comportamento em caso de falha |
|---|---|
| Mesmo número de colunas | Aborta e informa a contagem de cada arquivo |
| Mesmos nomes de colunas | Aborta e lista as colunas ausentes em cada lado |
| Mesma ordem de colunas | Aborta e informa as posições divergentes |
| Mesmo número de linhas (sem chave) | Aborta e sugere usar chave primária |
| Existência da chave primária | Aborta e lista as colunas disponíveis |
| Existência das colunas a ignorar | Aborta e lista as colunas disponíveis |

---

## 🔌 API

### `POST /api/compare`

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `file1` | File | ✅ | Arquivo Excel de referência |
| `file2` | File | ✅ | Arquivo Excel para comparar |
| `primary_key` | String | ❌ | Nome da coluna chave |
| `ignore_columns` | String | ❌ | Colunas a ignorar (separadas por vírgula) |

### `GET /api/download/<filename>`

Retorna o relatório `.xlsx` para download.

---

## 🤖 Gerado com IA

Este projeto foi inteiramente gerado com o auxílio do [Claude](https://claude.ai) (Anthropic) — incluindo arquitetura, código, testes e documentação. O desenvolvimento foi conduzido via prompts descritivos especificando requisitos funcionais, regras de comparação, tratamento de erros e entregáveis esperados.

---

## 📄 Licença

MIT
