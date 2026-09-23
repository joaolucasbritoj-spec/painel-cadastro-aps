"""Monta o dicionário de dados do painel a partir de data/historico.csv,
data/historico_microarea.csv e dos YAMLs de configuração.

Quem grava os arquivos de verdade (as páginas HTML de site/, com este
dicionário embutido em cada uma) é pipeline/gerar_site.py — este módulo só
monta o dicionário, para poder ser testado isoladamente.

Convenção de nomes, para nunca colidir com a verificação de privacidade
(pipeline/privacidade.py): nenhuma CHAVE de dicionário usa palavras como
"nome"/"cpf"/"cns" (banidas em QUALQUER chave, em qualquer profundidade) —
id de indicador, slug de equipe etc. sempre viram o VALOR de um campo
"id", nunca o nome do campo em si. É por isso que "indicadores" é uma
lista de objetos com {"id": "sem_cpf", ...} e não um dicionário
{"sem_cpf": {...}}.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.configuracao import (  # noqa: E402
    carregar_condicoes_saude, carregar_equipes, carregar_grupos, carregar_inconsistencias, carregar_metas,
)
from pipeline.historico import (  # noqa: E402
    HISTORICO_CONDICOES_CSV, HISTORICO_CSV, HISTORICO_ESTADO_CSV, HISTORICO_FAIXA_ETARIA_CSV, HISTORICO_GRUPOS_CSV,
    HISTORICO_MICROAREA_CSV,
)
from pipeline.score import score  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent

# slug de município -> caminho do brasão dentro de assets/, relativo à raiz
# do site publicado. Preenchido conforme os brasões que já temos em
# assets/brasoes/; município sem export ainda não aparece no site de qualquer
# forma (não tem dado no histórico), então não precisa entrar aqui ainda.
BRASOES = {
    'santa-maria-madalena': 'assets/brasoes/santa_maria_madalena.jfif',
    'saquarema': 'assets/brasoes/saquarema.png',
    'arraial-do-cabo': 'assets/brasoes/arraial_do_cabo.png',
    'quissama': 'assets/brasoes/quissama.jpg',
}


def _slugificar(texto: str) -> str:
    import unicodedata
    s = unicodedata.normalize('NFKD', texto)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip()
    s = ''.join(c if c.isalnum() else '-' for c in s)
    while '--' in s:
        s = s.replace('--', '-')
    return s.strip('-')


def _serie_equipe(historico: pd.DataFrame, historico_ma: pd.DataFrame, historico_estado: pd.DataFrame,
                   historico_faixa: pd.DataFrame, historico_grupos: pd.DataFrame, historico_condicoes: pd.DataFrame,
                   municipio: str, equipe_slug: str, inconsistencias_cfg: list[dict]) -> list[dict]:
    filtro = (historico['municipio'] == municipio) & (historico['equipe'] == equipe_slug)
    df_eq = historico.loc[filtro]

    serie = []
    for competencia in sorted(df_eq['competencia'].unique()):
        df_comp = df_eq[df_eq['competencia'] == competencia]
        total = int(df_comp['total_cadastrados'].iloc[0])
        presentes = set(df_comp['indicador_id'])

        resultado_score = score(historico, municipio, competencia, equipe_slug, inconsistencias_cfg)

        indicadores = []
        for inc in inconsistencias_cfg:
            if inc['id'] in presentes:
                valor = int(df_comp.loc[df_comp['indicador_id'] == inc['id'], 'valor'].iloc[0])
                indicadores.append({'id': inc['id'], 'valor': valor, 'ausente': False})
            else:
                indicadores.append({'id': inc['id'], 'ausente': True})

        item = {
            'competencia': competencia,
            'total_cadastrados': total,
            'score': resultado_score['score'],
            'dado_incompleto': resultado_score['dado_incompleto'],
            'indicadores_ausentes': resultado_score['indicadores_ausentes'],
            'indicadores': indicadores,
        }

        if not historico_ma.empty:
            filtro_ma = (
                (historico_ma['municipio'] == municipio)
                & (historico_ma['equipe'] == equipe_slug)
                & (historico_ma['competencia'] == competencia)
            )
            df_ma = historico_ma.loc[filtro_ma]
            if not df_ma.empty:
                microareas = []
                for microarea in sorted(df_ma['microarea'].unique()):
                    df_m = df_ma[df_ma['microarea'] == microarea]
                    total_ma = int(df_m['total_cadastrados'].iloc[0])
                    presentes_ma = set(df_m['indicador_id'])
                    indicadores_ma = []
                    for inc in inconsistencias_cfg:
                        if inc['id'] in presentes_ma:
                            valor = int(df_m.loc[df_m['indicador_id'] == inc['id'], 'valor'].iloc[0])
                            indicadores_ma.append({'id': inc['id'], 'valor': valor, 'ausente': False})
                        else:
                            indicadores_ma.append({'id': inc['id'], 'ausente': True})
                    microareas.append({
                        'microarea': microarea, 'total_cadastrados': total_ma, 'indicadores': indicadores_ma,
                    })
                item['microareas'] = microareas

        if not historico_estado.empty:
            filtro_estado = (
                (historico_estado['municipio'] == municipio)
                & (historico_estado['equipe'] == equipe_slug)
                & (historico_estado['competencia'] == competencia)
            )
            df_estado = historico_estado.loc[filtro_estado]
            if not df_estado.empty:
                linha = df_estado.iloc[0]
                item['estado'] = {
                    'limpos': int(linha['limpos']),
                    'com_1': int(linha['com_1']),
                    'com_2': int(linha['com_2']),
                    'com_3_mais': int(linha['com_3_mais']),
                    'media_inconsistencias': float(linha['media_inconsistencias']),
                    'fora_area': int(linha['fora_area']),
                    'total_bruto': int(linha['total_bruto']),
                }

        if not historico_faixa.empty:
            filtro_faixa = (
                (historico_faixa['municipio'] == municipio)
                & (historico_faixa['equipe'] == equipe_slug)
                & (historico_faixa['competencia'] == competencia)
            )
            df_faixa = historico_faixa.loc[filtro_faixa]
            if not df_faixa.empty:
                item['faixa_etaria'] = [
                    {
                        'faixa': row['faixa'], 'indicador_id': row['indicador_id'],
                        'valor': int(row['valor']), 'total_cadastrados': int(row['total_cadastrados']),
                    }
                    for _, row in df_faixa.iterrows()
                ]

        if not historico_grupos.empty:
            filtro_grupos = (
                (historico_grupos['municipio'] == municipio)
                & (historico_grupos['equipe'] == equipe_slug)
                & (historico_grupos['competencia'] == competencia)
            )
            df_grupos = historico_grupos.loc[filtro_grupos]
            if not df_grupos.empty:
                grupos_out = []
                for grupo_id, df_g in df_grupos.groupby('grupo_id'):
                    total_grupo = int(df_g['total_grupo'].iloc[0])
                    metricas = [
                        {'id': row['indicador_id'], 'valor': int(row['valor'])}
                        for _, row in df_g.iterrows()
                    ]
                    grupos_out.append({'id': grupo_id, 'total_grupo': total_grupo, 'metricas': metricas})
                item['grupos'] = grupos_out

        if not historico_condicoes.empty:
            filtro_cond = (
                (historico_condicoes['municipio'] == municipio)
                & (historico_condicoes['equipe'] == equipe_slug)
                & (historico_condicoes['competencia'] == competencia)
            )
            df_cond = historico_condicoes.loc[filtro_cond]
            if not df_cond.empty:
                item['condicoes'] = [
                    {'id': row['condicao_id'], 'valor': int(row['valor']), 'total_cadastrados': int(row['total_cadastrados'])}
                    for _, row in df_cond.iterrows()
                ]

        serie.append(item)
    return serie


def compilar() -> dict:
    equipes_cfg = carregar_equipes()
    inconsistencias_cfg = [i for i in carregar_inconsistencias() if i.get('ativo')]
    grupos_cfg = carregar_grupos()
    condicoes_cfg = carregar_condicoes_saude()
    metas_cfg = carregar_metas()

    historico = pd.read_csv(HISTORICO_CSV, dtype={'competencia': str, 'ine': str}) \
        if HISTORICO_CSV.exists() else pd.DataFrame()
    historico_ma = pd.read_csv(HISTORICO_MICROAREA_CSV, dtype={'competencia': str, 'ine': str}) \
        if HISTORICO_MICROAREA_CSV.exists() else pd.DataFrame()
    historico_estado = pd.read_csv(HISTORICO_ESTADO_CSV, dtype={'competencia': str, 'ine': str}) \
        if HISTORICO_ESTADO_CSV.exists() else pd.DataFrame()
    historico_faixa = pd.read_csv(HISTORICO_FAIXA_ETARIA_CSV, dtype={'competencia': str, 'ine': str}) \
        if HISTORICO_FAIXA_ETARIA_CSV.exists() else pd.DataFrame()
    historico_grupos = pd.read_csv(HISTORICO_GRUPOS_CSV, dtype={'competencia': str, 'ine': str}) \
        if HISTORICO_GRUPOS_CSV.exists() else pd.DataFrame()
    historico_condicoes = pd.read_csv(HISTORICO_CONDICOES_CSV, dtype={'competencia': str, 'ine': str}) \
        if HISTORICO_CONDICOES_CSV.exists() else pd.DataFrame()

    municipios_out = []
    for municipio in sorted(historico['municipio'].unique()) if not historico.empty else []:
        slug_municipio = _slugificar(municipio)
        equipes_do_municipio = [e for e in equipes_cfg if e['municipio'] == municipio]

        equipes_out = []
        for eq_cfg in equipes_do_municipio:
            serie = _serie_equipe(
                historico, historico_ma, historico_estado, historico_faixa, historico_grupos, historico_condicoes,
                municipio, eq_cfg['slug'], inconsistencias_cfg,
            )
            if not serie:
                continue  # equipe cadastrada mas ainda sem nenhuma carga no histórico
            equipes_out.append({
                'slug': eq_cfg['slug'],
                'equipe': eq_cfg['nome_exibicao'],
                'ine': eq_cfg.get('ine') or None,
                'link_drive': eq_cfg.get('link_drive') or None,
                'link_padlet': eq_cfg.get('link_padlet') or None,
                'serie': serie,
            })

        if not equipes_out:
            continue

        municipios_out.append({
            'slug': slug_municipio,
            'municipio': municipio,
            'brasao': BRASOES.get(slug_municipio),
            'competencias': sorted(historico.loc[historico['municipio'] == municipio, 'competencia'].unique()),
            'equipes': equipes_out,
        })

    return {
        'gerado_em': datetime.now().isoformat(timespec='seconds'),
        'metas': metas_cfg,
        'indicadores': [
            {
                'id': i['id'],
                'titulo': i['nome'],
                'descricao_curta': i['descricao_curta'],
                'meta_pct': i['meta_pct'],
                'entra_no_score': i['entra_no_score'],
                'orientacao': i['orientacao'].strip(),
            }
            for i in inconsistencias_cfg
        ],
        'grupos': [{'id': g['id'], 'titulo': g['nome']} for g in grupos_cfg],
        'condicoes': [
            {
                'id': c['id'],
                'titulo': c['nome'],
                'taxa_esperada_min_pct': c.get('taxa_esperada_min_pct'),
                'taxa_esperada_nota': c.get('taxa_esperada_nota'),
            }
            for c in condicoes_cfg.get('condicoes', [])
        ],
        'minimo_cadastros_para_alerta': condicoes_cfg.get('minimo_cadastros_para_alerta', 300),
        'municipios': municipios_out,
    }


if __name__ == '__main__':
    # compilar_site.py só monta o dicionário de dados — quem grava as
    # páginas de verdade (com os dados embutidos, para abrir local sem
    # servidor) é pipeline/gerar_site.py. Rodar este arquivo direto serve
    # só para inspecionar o dicionário no console.
    print(json.dumps(compilar(), ensure_ascii=False, indent=2))
