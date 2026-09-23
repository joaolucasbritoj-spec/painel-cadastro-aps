"""Teste do score() corrigido (pipeline/score.py) — histórico FICTÍCIO
montado aqui, cobrindo as correções em relação ao score_legado():
ausente ≠ zero, só indicadores ativos com entra_no_score contam, e o
resultado nunca sai de [0, 100].
"""
import pandas as pd
import pytest

from pipeline.score import score

INDICADORES_CFG = [
    {'id': 'sem_cpf', 'ativo': True, 'entra_no_score': True},
    {'id': 'sem_cns', 'ativo': True, 'entra_no_score': True},
    {'id': 'desatualizado', 'ativo': True, 'entra_no_score': False},
    {'id': 'indicador_desativado', 'ativo': False, 'entra_no_score': True},
]


def _linha(competencia, equipe, indicador_id, valor, total):
    return {
        'competencia': competencia, 'municipio': 'Município Teste', 'ine': '',
        'equipe': equipe, 'indicador_id': indicador_id, 'valor': valor, 'total_cadastrados': total,
    }


@pytest.fixture
def historico() -> pd.DataFrame:
    linhas = [
        # Equipe A, 2026-09: todos os indicadores presentes.
        _linha('2026-09', 'equipe-a', 'sem_cpf', 10, 100),
        _linha('2026-09', 'equipe-a', 'sem_cns', 5, 100),
        _linha('2026-09', 'equipe-a', 'desatualizado', 40, 100),

        # Equipe B, 2026-09: "sem_cns" AUSENTE (não tem linha nenhuma) —
        # diferente de vir com valor 0.
        _linha('2026-09', 'equipe-b', 'sem_cpf', 20, 100),
        _linha('2026-09', 'equipe-b', 'desatualizado', 5, 100),

        # Equipe C: soma dos indicadores passaria de 100% do total —
        # o score não pode ficar negativo.
        _linha('2026-09', 'equipe-c', 'sem_cpf', 90, 100),
        _linha('2026-09', 'equipe-c', 'sem_cns', 80, 100),
    ]
    return pd.DataFrame(linhas)


def test_indicador_ausente_nao_conta_como_zero_e_marca_dado_incompleto(historico):
    resultado = score(historico, 'Município Teste', '2026-09', 'equipe-b', INDICADORES_CFG)
    # sem_cns está AUSENTE — o score usa só sem_cpf (20/100), não soma 0 por ela.
    assert resultado['score'] == 80.0
    assert resultado['dado_incompleto'] is True
    assert resultado['indicadores_ausentes'] == ['sem_cns']


def test_indicador_presente_em_todos_da_dado_completo(historico):
    resultado = score(historico, 'Município Teste', '2026-09', 'equipe-a', INDICADORES_CFG)
    # sem_cpf (10) + sem_cns (5) = 15 -> score = 100 - 15 = 85. "desatualizado"
    # não entra (entra_no_score=false) mesmo estando presente.
    assert resultado['score'] == 85.0
    assert resultado['dado_incompleto'] is False
    assert resultado['indicadores_ausentes'] == []


def test_score_nunca_fica_abaixo_de_zero(historico):
    resultado = score(historico, 'Município Teste', '2026-09', 'equipe-c', INDICADORES_CFG)
    # 90 + 80 = 170, 170% > 100% do total -> score bruto seria -70.
    assert resultado['score'] == 0.0


def test_equipe_sem_nenhuma_linha_da_score_none(historico):
    resultado = score(historico, 'Município Teste', '2026-09', 'equipe-inexistente', INDICADORES_CFG)
    assert resultado['score'] is None
    assert resultado['dado_incompleto'] is True
