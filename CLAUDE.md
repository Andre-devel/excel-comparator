# CLAUDE.md — Excel Comparator

Guia de contexto para o Claude Code trabalhar neste projeto.

---

## O que é este projeto

Aplicação web para comparação célula a célula de dois arquivos `.xlsx`.
O usuário faz upload de dois arquivos, configura opções (chave primária, colunas a ignorar) e recebe um novo `.xlsx` com as divergências destacadas em vermelho e uma aba de resumo detalhado.

---

## Como rodar

```bash
# Instalar dependências
pip install -r requirements.txt

# Subir o servidor de desenvolvimento
python app.py
# → http://localhost:5000

# Rodar um teste rápido de ponta a ponta (sem servidor)
python -c "
from comparator import read_excel, validate_structure, compare, generate_report
import pandas as pd
df = pd.DataFrame({'ID':['1'],'Val':['a']})
df.to_excel('/tmp/t.xlsx', index=False)
d = read_excel('/tmp/t.xlsx', 'Arquivo 1')
print('OK', len(d), 'rows')
"
```

Não há testes automatizados com pytest ainda. Validações são feitas via scripts inline.

---

## Arquitetura — 4 camadas obrigatórias

O projeto é deliberadamente dividido em camadas isoladas. **Nunca misture responsabilidades entre elas.**

```
comparator/
├── reader.py       → CAMADA 1: leitura
├── validator.py    → CAMADA 2: validação estrutural
├── comparator.py   → CAMADA 3: comparação de dados
└── reporter.py     → CAMADA 4: geração do relatório
```

### Camada 1 — `reader.py`
- Única responsabilidade: abrir um `.xlsx` e retornar um `pd.DataFrame`
- **Sempre** usa `dtype=str` e `keep_default_na=False` → valores chegam como string pura, sem coerção
- Lança `FileReadError` para qualquer problema de leitura

### Camada 2 — `validator.py`
- Recebe dois DataFrames + parâmetros opcionais
- Executa 6 verificações em ordem (ver abaixo)
- Retorna `ValidationResult`; se `valid=False`, a comparação **não deve ser executada**
- Nunca acessa arquivos diretamente

**Ordem das verificações (não alterar):**
1. Mesmo número de colunas
2. Mesmos nomes de colunas
3. Mesma ordem de colunas
4. Mesmo número de linhas (somente sem `primary_key`)
5. Existência e unicidade da `primary_key`
6. Existência das `ignore_columns`

### Camada 3 — `comparator.py`
- Recebe dois DataFrames já validados
- Dois modos: **linha a linha** (padrão) e **chave primária**
- Comparação sempre literal: `str(v1) != str(v2)` — sem normalização, sem tolerância
- Retorna `ComparisonResult` com lista de `Divergence`
- Nunca grava arquivos

### Camada 4 — `reporter.py`
- Recebe `df1` + `ComparisonResult` e grava o `.xlsx` de saída
- Duas abas obrigatórias: `"Comparação"` (dados + vermelho) e `"Resumo"` (estatísticas)
- Usa `openpyxl` diretamente — não usa `df.to_excel()` para ter controle de estilo
- Nunca faz comparações de dados

---

## Convenções críticas

### Comparação de valores
```python
# CORRETO — comparação literal
if str(v1) != str(v2):

# ERRADO — nunca fazer isso
if v1.strip().lower() != v2.strip().lower()   # normalização proibida
if abs(float(v1) - float(v2)) < 0.01          # tolerância proibida
```

### Leitura de Excel
```python
# CORRETO — sempre estes parâmetros
pd.read_excel(path, dtype=str, keep_default_na=False)

# ERRADO — permite coerção de tipo pelo pandas
pd.read_excel(path)
```

### Erros
- Cada camada tem sua própria exceção (`FileReadError`, `ValidationError`)
- `app.py` captura e converte para JSON com `{"success": false, "error": "mensagem clara"}`
- Mensagens de erro devem ser descritivas: informar **o que** falhou, **em qual arquivo**, e **o que era esperado**

### Nomes de abas no relatório
- Aba principal: `"Comparação"` (com acento) — não alterar
- Aba de resumo: `"Resumo"` — não alterar
- Código downstream pode depender desses nomes

---

## API

### `POST /api/compare`
Form-data esperado:

| Campo | Tipo | Obrigatório |
|---|---|---|
| `file1` | File (.xlsx) | ✅ |
| `file2` | File (.xlsx) | ✅ |
| `primary_key` | string | ❌ |
| `ignore_columns` | string (vírgula) | ❌ |

Resposta de sucesso (`200`):
```json
{
  "success": true,
  "stats": { "total_rows": 0, "total_divergences": 0, "divergent_rows": 0 },
  "download_url": "/api/download/<uuid>.xlsx",
  "divergences": [{ "row": 1, "column": "Nome", "value1": "A", "value2": "B" }]
}
```

Resposta de erro (`400` / `422`):
```json
{ "success": false, "error": "mensagem descritiva" }
```

O preview de divergências na resposta JSON é limitado a **200 itens**. O arquivo `.xlsx` sempre contém tudo.

### `GET /api/download/<filename>`
- Retorna o `.xlsx` gerado
- Arquivos ficam em `outputs/` — não há expiração automática ainda

---

## Diretórios de runtime

```
uploads/    → arquivos temporários de upload (deletados após comparação)
outputs/    → relatórios gerados (persistem até limpeza manual)
```

Ambos são criados automaticamente pelo `app.py` na inicialização. Não commitar arquivos destes diretórios.

---

## Dependências

```
flask==3.0.3       → servidor web
pandas==2.2.2      → leitura e manipulação dos DataFrames
openpyxl==3.1.5    → escrita do .xlsx com controle de estilo
werkzeug==3.0.3    → utilitários Flask (secure_filename, etc.)
```

Não adicionar dependências sem necessidade clara. O projeto intencionalmente não usa `xlrd`, `xlwt` ou `xlsxwriter`.

---

## O que não fazer

- **Não normalizar valores** antes de comparar (espaços, case, datas, números)
- **Não executar a comparação** se `ValidationResult.valid == False`
- **Não misturar camadas** — reporter não compara, comparator não grava
- **Não usar `df.to_excel()`** no reporter (perde controle de estilo do openpyxl)
- **Não commitar** arquivos de `uploads/` ou `outputs/`
- **Não renomear** as abas `"Comparação"` e `"Resumo"` sem revisar o código cliente

---

## Melhorias planejadas (backlog)

- Suporte a múltiplas abas no Excel de entrada
- Expiração automática dos arquivos em `outputs/`
- Modo CLI: `python -m comparator file1.xlsx file2.xlsx`
- Testes automatizados com `pytest`
- Tolerância opcional para campos numéricos (opt-in explícito do usuário)
- Suporte a `.csv`
