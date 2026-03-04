"""
CAMADA DE VALIDAÇÃO ESTRUTURAL
Verifica se os dois arquivos têm estrutura compatível antes da comparação.
Nenhuma comparação de dados deve ocorrer se esta camada retornar erro.
"""

import pandas as pd
from dataclasses import dataclass, field
from typing import Optional


class ValidationError(Exception):
    """Erro de validação estrutural entre os dois arquivos."""
    pass


@dataclass
class ValidationResult:
    valid: bool
    errors: list[str] = field(default_factory=list)

    def add_error(self, msg: str):
        self.errors.append(msg)
        self.valid = False


def validate_structure(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    primary_key: Optional[str] = None,
    ignore_columns: Optional[list[str]] = None,
) -> ValidationResult:
    """
    Valida que os dois DataFrames possuem estrutura compatível para comparação.

    Verificações realizadas (em ordem):
      1. Mesmos nomes de colunas
      2. Mesma ordem de colunas
      3. Mesmo número de colunas
      4. Mesmo número de linhas (somente se primary_key não for informada)
      5. Existência da primary_key (se informada)
      6. Existência das colunas a ignorar (se informadas)

    Args:
        df1: DataFrame do Arquivo 1
        df2: DataFrame do Arquivo 2
        primary_key: Nome da coluna usada como chave primária (opcional)
        ignore_columns: Lista de colunas a serem ignoradas (opcional)

    Returns:
        ValidationResult com `valid=True` ou lista de erros descritivos.
    """
    result = ValidationResult(valid=True)

    cols1 = list(df1.columns)
    cols2 = list(df2.columns)

    # 1. Mesmo número de colunas
    if len(cols1) != len(cols2):
        result.add_error(
            f"Número de colunas divergente: "
            f"Arquivo 1 possui {len(cols1)} coluna(s), "
            f"Arquivo 2 possui {len(cols2)} coluna(s)."
        )
        # Sem sentido continuar as próximas checagens de colunas
        return result

    # 2. Mesmos nomes de colunas (independente de ordem)
    set1, set2 = set(cols1), set(cols2)
    only_in_1 = set1 - set2
    only_in_2 = set2 - set1

    if only_in_1:
        result.add_error(
            f"Coluna(s) presentes apenas no Arquivo 1: {sorted(only_in_1)}."
        )
    if only_in_2:
        result.add_error(
            f"Coluna(s) presentes apenas no Arquivo 2: {sorted(only_in_2)}."
        )

    if not result.valid:
        return result

    # 3. Mesma ordem de colunas
    if cols1 != cols2:
        mismatches = [
            f"posição {i+1}: '{c1}' vs '{c2}'"
            for i, (c1, c2) in enumerate(zip(cols1, cols2))
            if c1 != c2
        ]
        result.add_error(
            f"Ordem das colunas divergente. Diferenças: {'; '.join(mismatches)}."
        )
        return result

    # 4. Mesmo número de linhas (quando não há chave primária)
    if primary_key is None:
        if len(df1) != len(df2):
            result.add_error(
                f"Número de linhas divergente: "
                f"Arquivo 1 possui {len(df1)} linha(s), "
                f"Arquivo 2 possui {len(df2)} linha(s). "
                f"Para arquivos com número diferente de linhas, informe uma coluna como chave primária."
            )

    # 5. Existência da chave primária
    if primary_key is not None:
        if primary_key not in cols1:
            result.add_error(
                f"Chave primária '{primary_key}' não encontrada. "
                f"Colunas disponíveis: {cols1}."
            )
            return result

        # Verificar duplicatas na chave primária
        dup1 = df1[df1[primary_key].duplicated()][primary_key].tolist()
        dup2 = df2[df2[primary_key].duplicated()][primary_key].tolist()
        if dup1:
            result.add_error(
                f"Arquivo 1 possui valores duplicados na chave primária '{primary_key}': "
                f"{dup1[:10]}{'...' if len(dup1) > 10 else ''}."
            )
        if dup2:
            result.add_error(
                f"Arquivo 2 possui valores duplicados na chave primária '{primary_key}': "
                f"{dup2[:10]}{'...' if len(dup2) > 10 else ''}."
            )

    # 6. Existência das colunas a ignorar
    if ignore_columns:
        missing = [c for c in ignore_columns if c not in cols1]
        if missing:
            result.add_error(
                f"Coluna(s) a ignorar não encontradas: {missing}. "
                f"Colunas disponíveis: {cols1}."
            )

    return result
