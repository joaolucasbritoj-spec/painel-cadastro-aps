"""Teste de cada inconsistência do catálogo (config/inconsistencias.yaml),
rodando a REGRA DE VERDADE (mesmo interpretador usado pelo pipeline mensal)
sobre um cadastro fictício montado aqui — nunca sobre um export real.

Cada linha do DataFrame fictício foi pensada para acender UMA inconsistência
específica (e nenhuma outra sem querer), mais uma linha "cadastro completo"
que não deveria acender nenhuma.
"""
import pandas as pd
import pytest

from pipeline.configuracao import carregar_inconsistencias
from pipeline.regras import avaliar_regra, flag_duplicatas
from pipeline.tempo import ultimo_dia_competencia

COMPETENCIA = '2026-09'


@pytest.fixture
def df_ficticio() -> pd.DataFrame:
    linhas = [
        # 0: cadastro completo — não deve acender nenhuma inconsistência.
        dict(nome='Cadastro Completo', cpf='11111111111', cns='111111111111111',
             logradouro='Rua Um, 10', vinc_familia='SIM',
             dt_criacao='01/01/2020', dt_edicao='01/06/2026', nascimento='01/01/1980'),
        # 1: sem CPF.
        dict(nome='Sem Cpf', cpf='', cns='222222222222222',
             logradouro='Rua Dois, 20', vinc_familia='SIM',
             dt_criacao='01/01/2020', dt_edicao='01/06/2026', nascimento='02/02/1980'),
        # 2: sem CNS.
        dict(nome='Sem Cns', cpf='33333333333', cns='',
             logradouro='Rua Três, 30', vinc_familia='SIM',
             dt_criacao='01/01/2020', dt_edicao='01/06/2026', nascimento='03/03/1980'),
        # 3: sem endereço.
        dict(nome='Sem Endereco', cpf='44444444444', cns='444444444444444',
             logradouro='', vinc_familia='SIM',
             dt_criacao='01/01/2020', dt_edicao='01/06/2026', nascimento='04/04/1980'),
        # 4: não vinculado à família.
        dict(nome='Sem Vinculo', cpf='55555555555', cns='555555555555555',
             logradouro='Rua Cinco, 50', vinc_familia='NÃO',
             dt_criacao='01/01/2020', dt_edicao='01/06/2026', nascimento='05/05/1980'),
        # 5: nunca editado — criação == edição (data recente de propósito,
        # para não cair TAMBÉM em "desatualizado" e confundir o teste — os
        # dois podem coexistir num cadastro real, só não neste caso de teste).
        dict(nome='Nunca Editado Criacao Igual', cpf='66666666666', cns='666666666666666',
             logradouro='Rua Seis, 60', vinc_familia='SIM',
             dt_criacao='01/06/2026', dt_edicao='01/06/2026', nascimento='06/06/1980'),
        # 6: nunca editado — edição vazia.
        dict(nome='Nunca Editado Sem Data', cpf='77777777777', cns='777777777777777',
             logradouro='Rua Sete, 70', vinc_familia='SIM',
             dt_criacao='01/06/2026', dt_edicao='', nascimento='07/07/1980'),
        # 7 e 8: possível duplicata — mesmo CPF nos dois.
        dict(nome='Duplicata Um', cpf='88888888888', cns='888888888888881',
             logradouro='Rua Oito, 80', vinc_familia='SIM',
             dt_criacao='01/01/2020', dt_edicao='01/06/2026', nascimento='08/08/1980'),
        dict(nome='Duplicata Dois', cpf='88888888888', cns='888888888888882',
             logradouro='Rua Oito, 81', vinc_familia='SIM',
             dt_criacao='01/01/2020', dt_edicao='01/06/2026', nascimento='08/08/1981'),
        # 9: desatualizado — última edição bem antes de 24 meses atrás.
        dict(nome='Desatualizado', cpf='99999999999', cns='999999999999999',
             logradouro='Rua Nove, 90', vinc_familia='SIM',
             dt_criacao='01/01/2015', dt_edicao='01/01/2020', nascimento='09/09/1980'),
    ]
    df = pd.DataFrame(linhas)
    df['equipe'] = 'EQUIPE TESTE'
    df['micro_area'] = 'MICROÁREA 01'
    df['funcionario'] = 'ACS Teste'
    df['sexo'] = 'FEMININO'
    df['saida'] = ''
    return df


def _regra_por_id(indicador_id: str) -> dict:
    todas = {i['id']: i for i in carregar_inconsistencias()}
    return todas[indicador_id]


def _contexto(df: pd.DataFrame) -> dict:
    return {
        'ultimo_dia_competencia': ultimo_dia_competencia(COMPETENCIA),
        'flag_duplicata': flag_duplicatas(df),
    }


def test_sem_cpf(df_ficticio):
    mask = avaliar_regra(_regra_por_id('sem_cpf')['regra'], df_ficticio, _contexto(df_ficticio))
    assert list(df_ficticio.loc[mask, 'nome']) == ['Sem Cpf']


def test_sem_cns(df_ficticio):
    mask = avaliar_regra(_regra_por_id('sem_cns')['regra'], df_ficticio, _contexto(df_ficticio))
    assert list(df_ficticio.loc[mask, 'nome']) == ['Sem Cns']


def test_sem_endereco(df_ficticio):
    mask = avaliar_regra(_regra_por_id('sem_endereco')['regra'], df_ficticio, _contexto(df_ficticio))
    assert list(df_ficticio.loc[mask, 'nome']) == ['Sem Endereco']


def test_sem_vinculo_familiar(df_ficticio):
    mask = avaliar_regra(_regra_por_id('sem_vinculo_familiar')['regra'], df_ficticio, _contexto(df_ficticio))
    assert list(df_ficticio.loc[mask, 'nome']) == ['Sem Vinculo']


def test_nunca_editado_dois_casos(df_ficticio):
    mask = avaliar_regra(_regra_por_id('nunca_editado')['regra'], df_ficticio, _contexto(df_ficticio))
    assert set(df_ficticio.loc[mask, 'nome']) == {'Nunca Editado Criacao Igual', 'Nunca Editado Sem Data'}


def test_possivel_duplicata(df_ficticio):
    mask = avaliar_regra(_regra_por_id('possivel_duplicata')['regra'], df_ficticio, _contexto(df_ficticio))
    assert set(df_ficticio.loc[mask, 'nome']) == {'Duplicata Um', 'Duplicata Dois'}


def test_desatualizado(df_ficticio):
    mask = avaliar_regra(_regra_por_id('desatualizado')['regra'], df_ficticio, _contexto(df_ficticio))
    assert list(df_ficticio.loc[mask, 'nome']) == ['Desatualizado']


def test_cadastro_completo_nao_acende_nenhuma_inconsistencia(df_ficticio):
    contexto = _contexto(df_ficticio)
    linha_completa = df_ficticio[df_ficticio['nome'] == 'Cadastro Completo'].index[0]
    for inc in carregar_inconsistencias():
        if not inc.get('ativo'):
            continue
        mask = avaliar_regra(inc['regra'], df_ficticio, contexto)
        assert not mask.loc[linha_completa], f'"{inc["id"]}" acendeu para o cadastro completo — regra excessiva.'
