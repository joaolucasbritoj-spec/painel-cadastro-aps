"""Leitura/gravação com upsert de data/historico.csv e data/historico_microarea.csv.

Upsert = para cada (competência, município, equipe[, microárea]) recalculado
nesta carga, substitui as linhas antigas daquele recorte por completo (todas
as suas linhas de indicador) antes de acrescentar as novas — nunca duplica
uma carga repetida da mesma competência, e um indicador que deixou de ser
gerado (ex.: ficou "ativo: false") desaparece do recorte, em vez de ficar
com um valor congelado do mês anterior.
"""
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
HISTORICO_CSV = RAIZ / 'data' / 'historico.csv'
HISTORICO_MICROAREA_CSV = RAIZ / 'data' / 'historico_microarea.csv'
HISTORICO_ESTADO_CSV = RAIZ / 'data' / 'historico_estado.csv'

COLUNAS_HISTORICO = ['competencia', 'municipio', 'ine', 'equipe', 'indicador_id', 'valor', 'total_cadastrados']
COLUNAS_HISTORICO_MICROAREA = [
    'competencia', 'municipio', 'ine', 'equipe', 'microarea', 'indicador_id', 'valor', 'total_cadastrados',
]
# "Estado do cadastro" — indicadores complementares calculados por cadastro
# (Etapa 6): quantos cadastros ativos da equipe têm 0/1/2/3+ inconsistências
# ativas simultâneas, e quantos estavam "FORA DE ÁREA" (contados à parte,
# fora do total_cadastrados de todo o resto do painel).
COLUNAS_HISTORICO_ESTADO = [
    'competencia', 'municipio', 'ine', 'equipe', 'total_cadastrados',
    'limpos', 'com_1', 'com_2', 'com_3_mais', 'media_inconsistencias',
    'fora_area', 'total_bruto',
]
# Formato LONGO, como historico.csv (indicador vira LINHA, nunca coluna) —
# além de ser a convenção do projeto, evita que "sem_cpf"/"sem_cns" virem
# cabeçalho de coluna (a verificação de privacidade reprova isso, mesmo
# sendo dado agregado: é a mesma regra pensada para nunca deixar passar
# uma coluna nominal de verdade por engano).
HISTORICO_FAIXA_ETARIA_CSV = RAIZ / 'data' / 'historico_faixa_etaria.csv'
COLUNAS_HISTORICO_FAIXA_ETARIA = [
    'competencia', 'municipio', 'ine', 'equipe', 'faixa', 'indicador_id', 'valor', 'total_cadastrados',
]

HISTORICO_GRUPOS_CSV = RAIZ / 'data' / 'historico_grupos.csv'
COLUNAS_HISTORICO_GRUPOS = [
    'competencia', 'municipio', 'ine', 'equipe', 'grupo_id', 'indicador_id', 'valor', 'total_grupo',
]

HISTORICO_CONDICOES_CSV = RAIZ / 'data' / 'historico_condicoes.csv'
COLUNAS_HISTORICO_CONDICOES = ['competencia', 'municipio', 'ine', 'equipe', 'condicao_id', 'valor', 'total_cadastrados']


def _ler(caminho: Path, colunas: list[str]) -> pd.DataFrame:
    if not caminho.exists():
        return pd.DataFrame(columns=colunas)
    return pd.read_csv(caminho, dtype={'competencia': str, 'ine': str}, keep_default_na=True)


def _gravar(df: pd.DataFrame, caminho: Path, colunas: list[str]) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    ordenar_por = [c for c in ['competencia', 'municipio', 'equipe', 'microarea', 'indicador_id'] if c in colunas]
    df[colunas].sort_values(ordenar_por).to_csv(caminho, index=False, encoding='utf-8')


def upsert_historico(novas_linhas: pd.DataFrame, competencia: str, municipio: str, equipes_slugs: set[str],
                      caminho: Path = HISTORICO_CSV) -> pd.DataFrame:
    """Substitui, em data/historico.csv, todas as linhas de
    (competencia, municipio, equipe) para cada equipe em `equipes_slugs`,
    pelas linhas novas fornecidas. Devolve o DataFrame completo já
    atualizado (quem chama decide quando gravar de fato)."""
    atual = _ler(caminho, COLUNAS_HISTORICO)
    mantidas = ~(
        (atual['competencia'] == competencia)
        & (atual['municipio'] == municipio)
        & (atual['equipe'].isin(equipes_slugs))
    )
    resultado = pd.concat([atual.loc[mantidas], novas_linhas[COLUNAS_HISTORICO]], ignore_index=True)
    return resultado


def upsert_historico_microarea(novas_linhas: pd.DataFrame, competencia: str, municipio: str, equipes_slugs: set[str],
                                caminho: Path = HISTORICO_MICROAREA_CSV) -> pd.DataFrame:
    atual = _ler(caminho, COLUNAS_HISTORICO_MICROAREA)
    mantidas = ~(
        (atual['competencia'] == competencia)
        & (atual['municipio'] == municipio)
        & (atual['equipe'].isin(equipes_slugs))
    )
    resultado = pd.concat([atual.loc[mantidas], novas_linhas[COLUNAS_HISTORICO_MICROAREA]], ignore_index=True)
    return resultado


def upsert_historico_estado(novas_linhas: pd.DataFrame, competencia: str, municipio: str, equipes_slugs: set[str],
                             caminho: Path = HISTORICO_ESTADO_CSV) -> pd.DataFrame:
    atual = _ler(caminho, COLUNAS_HISTORICO_ESTADO)
    mantidas = ~(
        (atual['competencia'] == competencia)
        & (atual['municipio'] == municipio)
        & (atual['equipe'].isin(equipes_slugs))
    )
    resultado = pd.concat([atual.loc[mantidas], novas_linhas[COLUNAS_HISTORICO_ESTADO]], ignore_index=True)
    return resultado


def upsert_historico_faixa_etaria(novas_linhas: pd.DataFrame, competencia: str, municipio: str, equipes_slugs: set[str],
                                   caminho: Path = HISTORICO_FAIXA_ETARIA_CSV) -> pd.DataFrame:
    atual = _ler(caminho, COLUNAS_HISTORICO_FAIXA_ETARIA)
    mantidas = ~(
        (atual['competencia'] == competencia)
        & (atual['municipio'] == municipio)
        & (atual['equipe'].isin(equipes_slugs))
    )
    resultado = pd.concat([atual.loc[mantidas], novas_linhas[COLUNAS_HISTORICO_FAIXA_ETARIA]], ignore_index=True)
    return resultado


def upsert_historico_grupos(novas_linhas: pd.DataFrame, competencia: str, municipio: str, equipes_slugs: set[str],
                             caminho: Path = HISTORICO_GRUPOS_CSV) -> pd.DataFrame:
    atual = _ler(caminho, COLUNAS_HISTORICO_GRUPOS)
    mantidas = ~(
        (atual['competencia'] == competencia)
        & (atual['municipio'] == municipio)
        & (atual['equipe'].isin(equipes_slugs))
    )
    resultado = pd.concat([atual.loc[mantidas], novas_linhas[COLUNAS_HISTORICO_GRUPOS]], ignore_index=True)
    return resultado


def upsert_historico_condicoes(novas_linhas: pd.DataFrame, competencia: str, municipio: str, equipes_slugs: set[str],
                                caminho: Path = HISTORICO_CONDICOES_CSV) -> pd.DataFrame:
    atual = _ler(caminho, COLUNAS_HISTORICO_CONDICOES)
    mantidas = ~(
        (atual['competencia'] == competencia)
        & (atual['municipio'] == municipio)
        & (atual['equipe'].isin(equipes_slugs))
    )
    resultado = pd.concat([atual.loc[mantidas], novas_linhas[COLUNAS_HISTORICO_CONDICOES]], ignore_index=True)
    return resultado


def gravar_historico(df: pd.DataFrame, caminho: Path = HISTORICO_CSV) -> None:
    _gravar(df, caminho, COLUNAS_HISTORICO)


def gravar_historico_microarea(df: pd.DataFrame, caminho: Path = HISTORICO_MICROAREA_CSV) -> None:
    _gravar(df, caminho, COLUNAS_HISTORICO_MICROAREA)


def gravar_historico_estado(df: pd.DataFrame, caminho: Path = HISTORICO_ESTADO_CSV) -> None:
    _gravar(df, caminho, COLUNAS_HISTORICO_ESTADO)


def gravar_historico_faixa_etaria(df: pd.DataFrame, caminho: Path = HISTORICO_FAIXA_ETARIA_CSV) -> None:
    _gravar(df, caminho, COLUNAS_HISTORICO_FAIXA_ETARIA)


def gravar_historico_grupos(df: pd.DataFrame, caminho: Path = HISTORICO_GRUPOS_CSV) -> None:
    _gravar(df, caminho, COLUNAS_HISTORICO_GRUPOS)


def gravar_historico_condicoes(df: pd.DataFrame, caminho: Path = HISTORICO_CONDICOES_CSV) -> None:
    _gravar(df, caminho, COLUNAS_HISTORICO_CONDICOES)
