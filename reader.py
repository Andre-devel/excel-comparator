"""
CAMADA DE LEITURA
Responsável por carregar arquivos Excel de forma segura e isolada.
"""

import pandas as pd
from pathlib import Path


class FileReadError(Exception):
    """Erro ao tentar ler um arquivo Excel."""
    pass


def read_excel(filepath: str, label: str = "arquivo") -> pd.DataFrame:
    """
    Lê um arquivo .xlsx e retorna um DataFrame.

    Args:
        filepath: Caminho para o arquivo .xlsx
        label: Nome descritivo usado nas mensagens de erro ("Arquivo 1" ou "Arquivo 2")

    Returns:
        DataFrame com os dados do arquivo, sem alterações nos valores.

    Raises:
        FileReadError: Se o arquivo não puder ser lido por qualquer motivo.
    """
    path = Path(filepath)

    if not path.exists():
        raise FileReadError(f"{label}: arquivo não encontrado no caminho '{filepath}'.")

    if path.suffix.lower() not in (".xlsx", ".xls"):
        raise FileReadError(
            f"{label}: formato inválido '{path.suffix}'. Apenas arquivos .xlsx são suportados."
        )

    try:
        df = pd.read_excel(
            filepath,
            sheet_name=0,       # Sempre lê a primeira aba
            dtype=str,          # Lê tudo como string para comparação literal
            keep_default_na=False,  # Não converte strings vazias em NaN
        )
    except Exception as exc:
        raise FileReadError(
            f"{label}: não foi possível ler o arquivo. "
            f"O arquivo pode estar corrompido ou em formato inválido. Detalhe: {exc}"
        )

    return df
