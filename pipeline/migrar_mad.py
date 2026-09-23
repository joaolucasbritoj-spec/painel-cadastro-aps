"""
Migração do histórico legado (MAD.xlsx, formato LARGO) para o formato LONGO
usado pelo painel (data/historico.csv). Script de uso único por fonte de
histórico legado — não faz parte do comando mensal (pipeline/atualizar.py).

Uso:
  python pipeline/migrar_mad.py --mad "<caminho para MAD.xlsx>"

  # Com recálculo dos indicadores conhecidamente incorretos no MAD.xlsx,
  # a partir da exportação bruta daquela competência/município:
  python pipeline/migrar_mad.py --mad "<MAD.xlsx>" \
      --bruto "2026-09|Santa Maria Madalena|<export bruto de 2026-09.xlsx>"

  # Por padrão só mostra o que seria gravado (modo de conferência). Para
  # gravar de fato em data/historico.csv:
  python pipeline/migrar_mad.py --mad "<MAD.xlsx>" --bruto "..." --aplicar

Por quê o recálculo existe
---------------------------
Duas colunas do MAD.xlsx — "Sem endereço" e "Nunca editados" — estão em
ZERO em todas as linhas de todas as competências, o que não bate com a
exportação bruta (quando existe uma para conferir). Uma terceira,
"Mais de 2 anos sem edição", também não bate: ela foi calculada com uma
data fixa (27/04/2025) que está hoje hardcoded em gerar_relatorio.py, em
vez de "24 meses antes do fim da própria competência" — por isso o valor
não varia mês a mês do jeito que deveria.

Para essas três colunas, o valor do MAD só é usado se não houver
exportação bruta correspondente para recalcular — e mesmo assim ele entra
marcado como AUSENTE (não como zero), porque não temos como confirmá-lo.
"Ausente" aqui significa literalmente não escrever a linha no CSV: o
painel trata indicador ausente diferente de indicador zerado (mostra o
selo "dado incompleto" em vez de sugerir que está tudo certo).
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.configuracao import buscar_equipe, carregar_equipes  # noqa: E402
from pipeline.legado import carregar_ativos, is_blank  # noqa: E402
from pipeline.tempo import ultimo_dia_competencia  # noqa: E402

# Nome da coluna no MAD.xlsx -> id do indicador em config/inconsistencias.yaml
COLUNAS_MAD = {
    'Sem CPF': 'sem_cpf',
    'Sem CNS': 'sem_cns',
    'Sem endereço': 'sem_endereco',
    'Não Vinculado a Família': 'sem_vinculo_familiar',
    'Nunca editados': 'nunca_editado',
    'Mais de 2 anos sem edição': 'desatualizado',
    'Possíveis duplicatas': 'possivel_duplicata',
}

# Indicadores cujo valor no MAD.xlsx é conhecidamente pouco confiável sem
# uma exportação bruta para recalcular (ver docstring do módulo).
SUSPEITOS_NO_MAD = {'sem_endereco', 'nunca_editado', 'desatualizado'}

COLUNAS_HISTORICO = ['competencia', 'municipio', 'ine', 'equipe', 'indicador_id', 'valor', 'total_cadastrados']


def recalcular_indicador(indicador_id: str, df_ativos: pd.DataFrame, equipe_nome_mad: str, competencia: str) -> int | None:
    """Recalcula um indicador "suspeito" a partir do export bruto (já
    filtrado para ativos/dentro-de-área), para a equipe informada. Devolve
    None se a equipe não aparecer no export bruto.
    """
    alvo = equipe_nome_mad.strip().casefold()
    g = df_ativos[df_ativos['equipe'].str.strip().str.casefold() == alvo]
    if g.empty:
        return None

    if indicador_id == 'sem_endereco':
        return int(is_blank(g['logradouro']).sum())

    if indicador_id == 'nunca_editado':
        dt_c = pd.to_datetime(g['dt_criacao'], dayfirst=True, errors='coerce')
        dt_e = pd.to_datetime(g['dt_edicao'], dayfirst=True, errors='coerce')
        mesma_data = dt_c.dt.date == dt_e.dt.date
        edicao_vazia = is_blank(g['dt_edicao'])
        return int((mesma_data | edicao_vazia).sum())

    if indicador_id == 'desatualizado':
        referencia = ultimo_dia_competencia(competencia) - pd.DateOffset(months=24)
        dt_e = pd.to_datetime(g['dt_edicao'], dayfirst=True, errors='coerce')
        return int((dt_e < referencia).sum())

    raise ValueError(f'Indicador sem recálculo definido: {indicador_id}')


def migrar(mad_path: str, brutos: dict[tuple[str, str], pd.DataFrame]) -> tuple[pd.DataFrame, list[str]]:
    """Lê o MAD.xlsx e devolve (dataframe no formato longo, lista de avisos/diferenças)."""
    mad = pd.read_excel(mad_path, dtype={'Competência': str, 'Município': str, 'Equipe': str})
    equipes_cfg = carregar_equipes()

    linhas: list[dict] = []
    avisos: list[str] = []

    # Validação: competência duplicada para a mesma equipe
    dup = mad[mad.duplicated(subset=['Competência', 'Município', 'Equipe'], keep=False)]
    if not dup.empty:
        avisos.append(f'ALERTA: {len(dup)} linhas com competência duplicada para a mesma equipe no MAD.xlsx.')

    for _, row in mad.iterrows():
        competencia = str(row['Competência']).strip()
        municipio = str(row['Município']).strip()
        equipe_mad = str(row['Equipe']).strip()
        total = row['Total de cadastrados']

        if pd.isna(total) or int(total) == 0:
            avisos.append(f'ALERTA: total de cadastrados zerado/ausente em {municipio} | {equipe_mad} | {competencia}.')
            continue

        equipe_cfg = buscar_equipe(municipio, equipe_mad.upper(), equipes_cfg)
        if equipe_cfg is None:
            avisos.append(
                f'ALERTA: equipe "{equipe_mad}" ({municipio}) não encontrada em config/equipes.yaml — '
                f'linha ignorada. Cadastre a equipe e rode a migração de novo.'
            )
            continue

        slug = equipe_cfg['slug']
        ine = equipe_cfg.get('ine') or ''
        bruto_df = brutos.get((competencia, municipio))

        for col_mad, indicador_id in COLUNAS_MAD.items():
            valor_mad = row.get(col_mad)
            ausente_no_mad = pd.isna(valor_mad)

            if indicador_id in SUSPEITOS_NO_MAD:
                if bruto_df is None:
                    if not ausente_no_mad:
                        avisos.append(
                            f'{municipio} | {equipe_mad} | {competencia} | {indicador_id}: '
                            f'MAD.xlsx tem {int(valor_mad)}, mas é um indicador suspeito e não há '
                            f'exportação bruta desta competência para confirmar — gravado como AUSENTE.'
                        )
                    continue  # ausente: não escreve linha

                valor_real = recalcular_indicador(indicador_id, bruto_df, equipe_mad, competencia)
                if valor_real is None:
                    avisos.append(
                        f'{municipio} | {equipe_mad} | {competencia} | {indicador_id}: '
                        f'equipe não encontrada na exportação bruta — gravado como AUSENTE.'
                    )
                    continue

                if not ausente_no_mad and int(valor_mad) != valor_real:
                    avisos.append(
                        f'DIFERENÇA {municipio} | {equipe_mad} | {competencia} | {indicador_id}: '
                        f'MAD.xlsx={int(valor_mad)}  →  recalculado do bruto={valor_real}'
                    )
                valor_final = valor_real
            else:
                if ausente_no_mad:
                    avisos.append(
                        f'{municipio} | {equipe_mad} | {competencia} | {indicador_id}: '
                        f'valor ausente no MAD.xlsx — gravado como AUSENTE (não como zero).'
                    )
                    continue
                valor_final = int(valor_mad)

            linhas.append({
                'competencia': competencia,
                'municipio': municipio,
                'ine': ine,
                'equipe': slug,
                'indicador_id': indicador_id,
                'valor': valor_final,
                'total_cadastrados': int(total),
            })

    # Validação: indicador com 0 em TODAS as equipes de uma competência —
    # nunca tratar como "tudo corrigido", sempre como possível erro de
    # extração (aqui vira aviso; o pipeline mensal, na Etapa 3, transforma
    # isso num alerta de verdade no resumo impresso).
    df_long = pd.DataFrame(linhas, columns=COLUNAS_HISTORICO)
    if not df_long.empty:
        agrupado = df_long.groupby(['competencia', 'municipio', 'indicador_id'])['valor'].sum()
        for (competencia, municipio, indicador_id), soma in agrupado.items():
            if soma == 0:
                avisos.append(
                    f'POSSÍVEL ERRO DE EXTRAÇÃO: "{indicador_id}" está em 0 em TODAS as equipes de '
                    f'{municipio} em {competencia}.'
                )

    return df_long, avisos


def _parse_bruto_arg(valor: str) -> tuple[tuple[str, str], pd.DataFrame]:
    partes = valor.split('|')
    if len(partes) != 3:
        raise argparse.ArgumentTypeError(
            f'--bruto precisa ser "AAAA-MM|Município|caminho_do_arquivo", recebido: {valor!r}'
        )
    competencia, municipio, caminho = (p.strip() for p in partes)
    df_ativos = carregar_ativos(caminho)
    return (competencia, municipio), df_ativos


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--mad', required=True, help='Caminho para o MAD.xlsx (histórico legado, formato largo).')
    ap.add_argument(
        '--bruto', action='append', default=[],
        help='Repetível. "AAAA-MM|Município|caminho.xlsx" — exportação bruta usada para recalcular '
             'os indicadores suspeitos daquela competência/município.',
    )
    ap.add_argument('--saida', default=str(Path(__file__).resolve().parent.parent / 'data' / 'historico.csv'))
    ap.add_argument('--aplicar', action='store_true', help='Grava data/historico.csv. Sem esta flag, só mostra o que seria feito.')
    args = ap.parse_args()

    brutos = dict(_parse_bruto_arg(b) for b in args.bruto)

    df_long, avisos = migrar(args.mad, brutos)

    print(f'{len(df_long)} linhas geradas a partir de {args.mad}.\n')
    if avisos:
        print('Avisos e diferenças encontradas:')
        for a in avisos:
            print(f'  - {a}')
        print()

    print('Prévia (5 primeiras linhas):')
    print(df_long.head().to_string(index=False))

    if args.aplicar:
        Path(args.saida).parent.mkdir(parents=True, exist_ok=True)
        df_long.to_csv(args.saida, index=False, encoding='utf-8')
        print(f'\n✓ Gravado em {args.saida}')
    else:
        print('\n(modo de conferência — nada foi gravado; rode de novo com --aplicar para gravar)')


if __name__ == '__main__':
    main()
