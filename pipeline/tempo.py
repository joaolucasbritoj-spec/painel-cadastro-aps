"""Utilidades de competência (AAAA-MM)."""
import pandas as pd


def ultimo_dia_competencia(competencia: str) -> pd.Timestamp:
    return pd.Period(competencia, freq='M').end_time.normalize()


def competencia_anterior(competencia: str) -> str:
    periodo = pd.Period(competencia, freq='M') - 1
    return str(periodo)
