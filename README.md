# Excel Comparator

Ferramenta web para comparação célula a célula de dois arquivos `.xlsx`, com destaque visual das divergências e relatório detalhado.

---

## Stack Tecnológica

| Camada | Tecnologia | Justificativa |
|--------|-----------|---------------|
| Backend | **Python + Flask** | Simples, direto, excelente suporte a arquivos |
| Leitura/Escrita Excel | **pandas + openpyxl** | Combinação padrão da indústria para Excel em Python |
| Frontend | **HTML/CSS/JS vanilla** | Zero dependências, rápido, sem build step |

---

## Instalação

```bash
# 1. Clone o repositório
cd excel-comparator

# 2. Crie um ambiente virtual (recomendado)
python -m venv .venv
source .venv/bin/activate      # Linux/Mac
.venv\Scripts\activate         # Windows

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Execute
python app.py
```

Acesse: http://localhost:5000

---

## Estrutura do Projeto

```
excel-comparator/
│
├── app.py                  # Servidor Flask + rotas da API
│
├── comparator/             # Pacote com as 4 camadas de lógica
│   ├── __init__.py
│   ├── reader.py           # CAMADA 1: Leitura dos arquivos
│   ├── validator.py        # CAMADA 2: Validação estrutural
│   ├── comparator.py       # CAMADA 3: Lógica de comparação
│   └── reporter.py         # CAMADA 4: Geração do relatório .xlsx
│
├── templates/
│   └── index.html          # Interface web
│
├── uploads/                # Arquivos temporários (limpos após uso)
├── outputs/                # Relatórios gerados
└── requirements.txt
```

---

## Lógica de Comparação

### Validação Estrutural (sempre executada primeiro)

Antes de qualquer comparação, o sistema verifica:

1. **Mesmo número de colunas** — aborta se divergir
2. **Mesmos nomes de colunas** — reporta quais faltam em cada arquivo
3. **Mesma ordem de colunas** — reporta as posições divergentes
4. **Mesmo número de linhas** — apenas quando não há chave primária
5. **Existência da chave primária** — se informada, verifica existência e unicidade
6. **Existência das colunas a ignorar** — valida cada uma antes de prosseguir

Se qualquer validação falhar, a comparação **não é executada** e o erro é retornado de forma clara.

### Modo Linha a Linha (padrão)

Quando nenhuma chave primária é informada:
- `df1.iloc[i]` é comparado com `df2.iloc[i]` para todo `i`
- Garante velocidade O(n × c) onde n=linhas, c=colunas

### Modo Chave Primária

Quando uma coluna chave é informada:
- Os DataFrames são indexados pela chave
- A interseção de chaves é comparada registro a registro
- Chaves presentes apenas em um arquivo são registradas como divergências especiais

### Regras de Comparação

- **Todos os valores são lidos como `str`** — sem coerção de tipo
- **Comparação exata**: `v1 != v2` (case-sensitive, espaços preservados)
- **Sem tolerância numérica** — `"1.0"` ≠ `"1"`
- **Sem normalização de datas** — `"2024-01-01"` ≠ `"01/01/2024"`

---

## Saída

O arquivo `.xlsx` gerado contém:

### Aba "Comparação"
- Todos os dados do Arquivo 1
- Células divergentes com **fundo vermelho**
- Cabeçalho azul escuro, linhas alternadas para leitura
- Legenda ao final

### Aba "Resumo"
- Total de linhas comparadas
- Total de divergências
- Linhas com ao menos uma divergência
- Tabela detalhada: linha | coluna | valor arquivo 1 | valor arquivo 2

---

## Tratamento de Erros

| Cenário | Mensagem retornada |
|---------|-------------------|
| Arquivo não enviado | "Arquivo N: nenhum arquivo enviado." |
| Formato inválido | "Arquivo N: formato inválido. Apenas .xlsx aceitos." |
| Arquivo corrompido | "Arquivo N: não foi possível ler o arquivo... [detalhe]" |
| Número de colunas diferente | "Número de colunas divergente: Arquivo 1 possui X, Arquivo 2 possui Y." |
| Colunas com nomes diferentes | "Coluna(s) presentes apenas no Arquivo N: [lista]" |
| Ordem de colunas diferente | "Ordem das colunas divergente. Diferenças: posição N: 'X' vs 'Y'" |
| Número de linhas diferente (sem PK) | "Número de linhas divergente... Informe uma chave primária." |
| Chave primária inexistente | "Chave primária 'X' não encontrada. Colunas disponíveis: [lista]" |
| Chave primária com duplicatas | "Arquivo N possui valores duplicados na chave primária 'X': [lista]" |
| Coluna a ignorar inexistente | "Coluna(s) a ignorar não encontradas: [lista]" |

---

## API

### POST /api/compare

**Form-data:**
| Campo | Tipo | Obrigatório | Descrição |
|-------|------|-------------|-----------|
| `file1` | File | ✅ | Arquivo Excel de referência |
| `file2` | File | ✅ | Arquivo Excel para comparar |
| `primary_key` | String | ❌ | Nome da coluna chave |
| `ignore_columns` | String | ❌ | Colunas a ignorar (vírgula) |

**Resposta de sucesso:**
```json
{
  "success": true,
  "stats": {
    "total_rows": 5000,
    "total_divergences": 37,
    "divergent_rows": 22
  },
  "download_url": "/api/download/comparacao_uuid.xlsx",
  "divergences": [...]
}
```

### GET /api/download/<filename>

Retorna o arquivo `.xlsx` para download.

---

## Melhorias Futuras

1. **Autenticação / expiração de arquivos** — remover relatórios antigos automaticamente
2. **Comparação multi-aba** — comparar planilhas com múltiplas abas
3. **Tolerância configurável** — para campos numéricos/datas, se o usuário optar
4. **Exportação para CSV** do resumo de divergências
5. **Modo CLI** — uso sem servidor via `python -m comparator file1.xlsx file2.xlsx`
6. **Paginação no backend** — para preview de mais de 200 divergências na UI
7. **Histórico de comparações** — salvar sessões com banco de dados leve (SQLite)
8. **Suporte a .csv** — expandir além de .xlsx
