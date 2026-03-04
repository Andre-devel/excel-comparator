"""
CAMADA DE COMPARAÇÃO
Realiza a comparação célula a célula entre dois DataFrames válidos.

Regras de comparação:
  - Comparação literal: os valores são comparados exatamente como lidos do Excel.
  - Espaços são preservados.
  - Maiúsculas/minúsculas são preservadas.
  - Sem tolerância numérica.
  - Sem normalização de datas.
"""

import pandas as pd
from dataclasses import dataclass
from typing import Optional


@dataclass
class Divergence:
    """Representa uma célula com valor divergente entre os dois arquivos."""
    row_number: int        # Número da linha (1-based, relativo aos dados)
    column_name: str       # Nome da coluna
    value_file1: str       # Valor no Arquivo 1
    value_file2: str       # Valor no Arquivo 2


@dataclass
class ComparisonResult:
    total_rows_compared: int
    divergences: list[Divergence]

    @property
    def total_divergences(self) -> int:
        return len(self.divergences)

    @property
    def divergent_rows(self) -> set[int]:
        return {d.row_number for d in self.divergences}


def compare(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    primary_key: Optional[str] = None,
    ignore_columns: Optional[list[str]] = None,
) -> ComparisonResult:
    """
    Compara dois DataFrames e retorna todas as divergências encontradas.

    Args:
        df1: DataFrame do Arquivo 1 (referência)
        df2: DataFrame do Arquivo 2 (comparado)
        primary_key: Coluna usada para alinhar registros. Se None, usa ordem das linhas.
        ignore_columns: Colunas que devem ser completamente ignoradas na comparação.

    Returns:
        ComparisonResult com total de linhas comparadas e lista de divergências.
    """
    ignore_set = set(ignore_columns or [])

    # Colunas que serão efetivamente comparadas
    compare_cols = [
        col for col in df1.columns
        if col not in ignore_set and col != primary_key
    ]

    divergences: list[Divergence] = []

    if primary_key:
        # ---- MODO CHAVE PRIMÁRIA ----
        # Alinha registros pelo valor da chave antes de comparar.
        df1_indexed = df1.set_index(primary_key)
        df2_indexed = df2.set_index(primary_key)

        # Considera apenas as chaves presentes em ambos os arquivos
        common_keys = df1_indexed.index.intersection(df2_indexed.index)
        only_in_1 = df1_indexed.index.difference(df2_indexed.index).tolist()
        only_in_2 = df2_indexed.index.difference(df1_indexed.index).tolist()

        rows_compared = len(common_keys)

        for key in common_keys:
            row1 = df1_indexed.loc[key]
            row2 = df2_indexed.loc[key]
            # row_number = posição no df1 original (1-based)
            row_number = df1.index[df1[primary_key] == key][0] + 1

            for col in compare_cols:
                v1 = str(row1[col]) if col in row1 else ""
                v2 = str(row2[col]) if col in row2 else ""
                if v1 != v2:
                    divergences.append(
                        Divergence(
                            row_number=row_number,
                            column_name=col,
                            value_file1=v1,
                            value_file2=v2,
                        )
                    )

        # Registrar linhas ausentes como divergências especiais
        for key in only_in_1:
            row_number = df1.index[df1[primary_key] == key][0] + 1
            divergences.append(
                Divergence(
                    row_number=row_number,
                    column_name=f"[CHAVE: {primary_key}]",
                    value_file1=str(key),
                    value_file2="<ausente no Arquivo 2>",
                )
            )

        for key in only_in_2:
            divergences.append(
                Divergence(
                    row_number=0,  # Não existe no Arquivo 1
                    column_name=f"[CHAVE: {primary_key}]",
                    value_file1="<ausente no Arquivo 1>",
                    value_file2=str(key),
                )
            )

    else:
        # ---- MODO LINHA A LINHA ----
        # Validação garantiu que len(df1) == len(df2).
        rows_compared = len(df1)

        for idx in range(rows_compared):
            row1 = df1.iloc[idx]
            row2 = df2.iloc[idx]
            row_number = idx + 1  # 1-based

            for col in compare_cols:
                v1 = str(row1[col])
                v2 = str(row2[col])
                if v1 != v2:
                    divergences.append(
                        Divergence(
                            row_number=row_number,
                            column_name=col,
                            value_file1=v1,
                            value_file2=v2,
                        )
                    )

    return ComparisonResult(
        total_rows_compared=rows_compared,
        divergences=divergences,
    )
