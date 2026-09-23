"""Teste de fidelidade obrigatório: score_legado() precisa reproduzir
exatamente os números que o Power BI já mostrava, calculados a partir do
histórico migrado do MAD.xlsx (data/historico.csv, dado agregado, sem
nenhuma informação nominal — por isso pode ficar versionado e ser lido
direto aqui, ao contrário dos testes das funções de detecção de
inconsistência por cadastro, que usam dado fictício gerado no teste).
"""
from pathlib import Path

import pandas as pd
import pytest

from pipeline.score import score_legado

RAIZ = Path(__file__).resolve().parent.parent
HISTORICO = RAIZ / 'data' / 'historico.csv'


@pytest.fixture
def historico() -> pd.DataFrame:
    return pd.read_csv(HISTORICO, dtype={'competencia': str, 'ine': str})


def test_score_legado_2026_07_bate_com_power_bi(historico):
    assert score_legado(historico, 'Santa Maria Madalena', '2026-07') == 44.31


def test_score_legado_2026_08_bate_com_power_bi(historico):
    assert score_legado(historico, 'Santa Maria Madalena', '2026-08') == 44.09


def test_score_legado_competencia_sem_dados_da_zero(historico):
    assert score_legado(historico, 'Santa Maria Madalena', '1999-01') == 0.0
