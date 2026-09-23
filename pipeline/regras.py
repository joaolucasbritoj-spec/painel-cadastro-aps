"""Interpreta a `regra` de cada inconsistência (config/inconsistencias.yaml)
sobre um DataFrame de cadastros já carregado, devolvendo uma flag booleana
por cadastro (Series alinhada ao índice do df de entrada).

Isso é o que permite tirar/acrescentar inconsistência só editando o YAML:
nenhuma regra de negócio fica hardcoded aqui, só os TIPOS de regra que o
YAML pode expressar.
"""
from __future__ import annotations

import pandas as pd

from pipeline.legado import is_blank, is_blank_cpf, detectar_duplicatas

_FUNCOES = {
    'is_blank': is_blank,
    'is_blank_cpf': is_blank_cpf,
}


def flag_duplicatas(df_completo: pd.DataFrame) -> pd.Series:
    """Marca True nas linhas de df_completo que aparecem no resultado de
    detectar_duplicatas() (função original de gerar_relatorio.py, não
    reescrita aqui). detectar_duplicatas() devolve um DataFrame novo, sem
    o índice original — por isso o casamento de volta é feito pela
    combinação de colunas que ela preserva verbatim (equipe/nome/cpf/cns/
    nascimento/dt_criacao/dt_edicao), que já é o que define cada linha
    como candidata a duplicata.
    """
    flag = pd.Series(False, index=df_completo.index)
    resultado = detectar_duplicatas(df_completo)
    if resultado.empty:
        return flag

    chave_cols = [c for c in ['equipe', 'nome', 'cpf', 'cns', 'nascimento', 'dt_criacao', 'dt_edicao']
                  if c in resultado.columns and c in df_completo.columns]
    chaves_duplicadas = set(map(tuple, resultado[chave_cols].astype(str).values))
    chaves_df = df_completo[chave_cols].astype(str).apply(tuple, axis=1)
    return chaves_df.isin(chaves_duplicadas)


def avaliar_regra(regra: dict, df: pd.DataFrame, contexto: dict) -> pd.Series:
    """Avalia recursivamente uma regra do YAML sobre `df`.

    `contexto` traz o que não dá para calcular olhando só para uma coluna:
      - 'ultimo_dia_competencia': Timestamp do último dia da competência.
      - 'flag_duplicata': Series (mesmo índice de df) já calculada por
        flag_duplicatas(), para o tipo 'especial'/detectar_duplicatas.
    """
    tipo = regra['tipo']

    if tipo == 'funcao':
        funcao = _FUNCOES[regra['funcao']]
        return funcao(df[regra['coluna']])

    if tipo == 'comparacao':
        coluna = df[regra['coluna']]
        operador = regra['operador']
        referencia = regra.get('referencia')

        if referencia == 'ultimo_dia_competencia_menos_24_meses':
            limite = contexto['ultimo_dia_competencia'] - pd.DateOffset(months=24)
            coluna_dt = pd.to_datetime(coluna, dayfirst=True, errors='coerce')
            if operador == '<':
                return coluna_dt < limite
            raise ValueError(f'Operador não suportado para comparação de data: {operador}')

        valor = regra['valor']
        coluna_str = coluna.astype(str).str.strip()
        if operador == '==':
            return coluna_str == valor
        if operador == '!=':
            return coluna_str != valor
        raise ValueError(f'Operador não suportado: {operador}')

    if tipo == 'colunas_iguais':
        a = pd.to_datetime(df[regra['coluna_a']], dayfirst=True, errors='coerce')
        b = pd.to_datetime(df[regra['coluna_b']], dayfirst=True, errors='coerce')
        return a.dt.date == b.dt.date

    if tipo == 'ou':
        resultado = pd.Series(False, index=df.index)
        for sub in regra['condicoes']:
            resultado = resultado | avaliar_regra(sub, df, contexto)
        return resultado

    if tipo == 'especial':
        if regra['funcao'] == 'detectar_duplicatas':
            flag_completo = contexto['flag_duplicata']
            return flag_completo.loc[df.index]
        raise ValueError(f'Função especial não suportada: {regra["funcao"]}')

    # Os dois tipos abaixo são de config/grupos.yaml (grupos prioritários),
    # não de inconsistencias.yaml — mas como é o mesmo interpretador de
    # regra+condição, ganham o mesmo lugar em vez de duplicar a lógica.
    if tipo == 'contem':
        coluna = df[regra['coluna']]
        return coluna.str.upper().str.contains(regra['texto'].upper(), na=False)

    if tipo == 'faixa_etaria':
        coluna = df[regra['coluna']]
        minimo = regra.get('min')
        maximo = regra.get('max')
        resultado = pd.Series(True, index=df.index)
        if minimo is not None:
            resultado &= coluna >= minimo
        if maximo is not None:
            resultado &= coluna <= maximo
        return resultado & coluna.notna()

    raise ValueError(f'Tipo de regra não suportado: {tipo}')
