"""Indicadores principais calculados sobre data/historico.csv."""
import pandas as pd

# Indicadores somados pelo Score de Qualidade LEGADO (fórmula DAX original
# do Power BI): Sem CPF + Sem CNS + Sem endereço + Não Vinculado à Família +
# Nunca editados + Possíveis duplicatas. Propositalmente NÃO inclui
# "desatualizado" (Mais de 2 anos sem edição).
INDICADORES_SCORE_LEGADO = [
    'sem_cpf', 'sem_cns', 'sem_endereco', 'sem_vinculo_familiar',
    'nunca_editado', 'possivel_duplicata',
]


def score_legado(historico: pd.DataFrame, municipio: str, competencia: str) -> float:
    """Réplica fiel da medida DAX original:

        Score = 100 − (Σ inconsistências / Total de cadastrados × 100)

    Agregação entre equipes é SEMPRE soma dos numeradores / soma do
    denominador, nunca média de percentuais — por isso a filtragem abaixo
    soma direto sobre as linhas do município inteiro, sem calcular um score
    por equipe e tirar média depois.

    Só existe para o teste de fidelidade (comparar com os números que o
    Power BI já mostrava). O painel novo usa score(), que respeita
    indicador ativo/entra_no_score e ausente-vs-zero.
    """
    filtro = (historico['municipio'] == municipio) & (historico['competencia'] == competencia)
    df = historico.loc[filtro]

    total = df.drop_duplicates(subset=['equipe'])['total_cadastrados'].sum()
    if total == 0:
        return 0.0

    soma_incons = df.loc[df['indicador_id'].isin(INDICADORES_SCORE_LEGADO), 'valor'].sum()
    return round(100 - (soma_incons / total * 100), 2)


def score(historico: pd.DataFrame, municipio: str, competencia: str, equipe: str,
          inconsistencias_cfg: list[dict]) -> dict:
    """Score de Qualidade corrigido de uma equipe, numa competência.

    Correções em relação ao score_legado():
      1. Soma só as inconsistências ATIVAS com entra_no_score=true em
         config/inconsistencias.yaml (não uma lista fixa no código).
      2. Indicador ausente para essa equipe/competência não entra como
         zero — o score é calculado sem ele, e dado_incompleto=True avisa
         que faltou informação (o painel mostra o selo "dado incompleto").
      3. Score sempre entre 0 e 100.

    Devolve um dict porque o painel precisa tanto do número quanto do selo
    de dado incompleto e de QUAIS indicadores faltaram.
    """
    ids_score = {c['id'] for c in inconsistencias_cfg if c.get('ativo') and c.get('entra_no_score')}

    filtro = (
        (historico['municipio'] == municipio)
        & (historico['competencia'] == competencia)
        & (historico['equipe'] == equipe)
    )
    df = historico.loc[filtro]

    if df.empty:
        return {'score': None, 'dado_incompleto': True, 'indicadores_ausentes': sorted(ids_score),
                'total_cadastrados': 0}

    total = int(df['total_cadastrados'].iloc[0])
    presentes = set(df['indicador_id']) & ids_score
    ausentes = ids_score - presentes

    soma = df.loc[df['indicador_id'].isin(presentes), 'valor'].sum()
    valor = 100 - (soma / total * 100) if total else 0.0
    valor = max(0.0, min(100.0, valor))

    return {
        'score': round(valor, 2),
        'dado_incompleto': bool(ausentes),
        'indicadores_ausentes': sorted(ausentes),
        'total_cadastrados': total,
    }
