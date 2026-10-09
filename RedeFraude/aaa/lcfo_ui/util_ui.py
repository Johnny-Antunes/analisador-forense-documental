"""
Pequenos utilitários de UI.
"""

from __future__ import annotations


import pandas as pd


def sanitizar_df_para_tabela(df: pd.DataFrame) -> pd.DataFrame:
    return df.astype(object).where(pd.notnull(df), None)
