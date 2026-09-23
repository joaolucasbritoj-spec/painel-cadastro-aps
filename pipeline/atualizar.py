"""
Pipeline mensal: lê uma exportação bruta, marca as inconsistências ativas
por cadastro, agrega por equipe e por microárea, faz upsert em
data/historico.csv e data/historico_microarea.csv, gera as listagens de
correção por equipe e imprime um resumo com a variação em relação ao mês
anterior e os alertas de validação.

Uso:
  python pipeline/atualizar.py --arquivo "<export.xlsx>" --municipio "Santa Maria Madalena" --competencia 2026-09
  python pipeline/atualizar.py --lote "<pasta>" --competencia 2026-09
    (a pasta precisa ter uma subpasta por município, com um .xlsx dentro de cada)

O município NUNCA é adivinhado a partir do nome do arquivo — vem sempre de
--municipio ou do nome da subpasta em --lote, porque o sistema de origem
não garante que o nome do arquivo traga o município (às vezes traz a
unidade). Ver config/equipes.yaml.

No final, regenera site/ inteiro (dados + páginas HTML) e roda a
verificação de privacidade antes de considerar a carga publicável.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.configuracao import (  # noqa: E402
    buscar_equipe, carregar_caminhos, carregar_condicoes_saude, carregar_equipes, carregar_grupos,
    carregar_inconsistencias, carregar_metas,
)
from pipeline.legado import _save_listing_xlsx, detectar_duplicatas, is_blank, is_blank_cpf, load_data  # noqa: E402
from pipeline.historico import (  # noqa: E402
    HISTORICO_CONDICOES_CSV, HISTORICO_CSV, HISTORICO_ESTADO_CSV, HISTORICO_FAIXA_ETARIA_CSV, HISTORICO_GRUPOS_CSV,
    HISTORICO_MICROAREA_CSV, gravar_historico, gravar_historico_condicoes, gravar_historico_estado,
    gravar_historico_faixa_etaria, gravar_historico_grupos, gravar_historico_microarea, upsert_historico,
    upsert_historico_condicoes, upsert_historico_estado, upsert_historico_faixa_etaria, upsert_historico_grupos,
    upsert_historico_microarea,
)
from pipeline.regras import avaliar_regra, flag_duplicatas  # noqa: E402
from pipeline.score import score  # noqa: E402
from pipeline.tempo import competencia_anterior, ultimo_dia_competencia  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent

# Faixas etárias do indicador complementar "Sem CPF por faixa etária".
FAIXAS_ETARIAS = [
    ('<1', 0, 0), ('1-4', 1, 4), ('5-9', 5, 9),
    ('10-17', 10, 17), ('18-59', 18, 59), ('60+', 60, 999),
]


def _sanitizar_nome_arquivo(nome: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', '-', nome)


def _filtrar_ativos(df_completo: pd.DataFrame) -> pd.DataFrame:
    """Mesmo filtro de gerar_relatorio.py: exclui saída de cadastro e
    "FORA DE ÁREA". Preserva o índice original de df_completo (importante
    para casar a flag de duplicata calculada sobre o df completo)."""
    df = df_completo
    if 'saida' in df.columns:
        df = df[is_blank(df['saida'])]
    if 'micro_area' in df.columns:
        df = df[~df['micro_area'].str.contains('FORA', case=False, na=False)]
    return df.copy()


def validar_equipes(municipio: str, df_completo: pd.DataFrame, equipes_cfg: list[dict]) -> list[str]:
    avisos = []
    equipes_no_arquivo = {str(e).strip().upper() for e in df_completo['equipe'].dropna().unique()}
    equipes_no_municipio = {
        e['nome_exportacao'].strip().upper(): e for e in equipes_cfg if e['municipio'] == municipio
    }

    nao_cadastradas = equipes_no_arquivo - set(equipes_no_municipio)
    for nome in sorted(nao_cadastradas):
        avisos.append(
            f'ALERTA: equipe "{nome}" aparece no arquivo mas não está cadastrada em config/equipes.yaml '
            f'para {municipio} — os cadastros dela foram ignorados. Confira se o arquivo é do município certo.'
        )

    ausentes = set(equipes_no_municipio) - equipes_no_arquivo
    for nome in sorted(ausentes):
        avisos.append(
            f'ALERTA: a equipe "{nome}" está cadastrada para {municipio} em config/equipes.yaml, mas não '
            f'aparece nesta exportação.'
        )
    return avisos


def processar_arquivo(caminho_arquivo: str, municipio: str, competencia: str) -> dict:
    equipes_cfg = carregar_equipes()
    inconsistencias_cfg = [i for i in carregar_inconsistencias() if i.get('ativo')]
    grupos_cfg = carregar_grupos()
    condicoes_cfg = carregar_condicoes_saude()
    metas = carregar_metas()
    caminhos_cfg = carregar_caminhos()

    print(f'\nCarregando {caminho_arquivo} ({municipio}, competência {competencia})...')
    df_completo = load_data(caminho_arquivo)

    avisos = validar_equipes(municipio, df_completo, equipes_cfg)

    df_ativos = _filtrar_ativos(df_completo)
    contexto = {
        'ultimo_dia_competencia': ultimo_dia_competencia(competencia),
        'flag_duplicata': flag_duplicatas(df_completo),  # usa df completo (com inativos), de propósito
    }

    # "FORA DE ÁREA" por equipe: cadastros ativos (sem saída) mas fora de
    # qualquer microárea real — contados à parte, nunca dentro do
    # total_cadastrados do resto do painel (ver COLUNAS_HISTORICO_ESTADO).
    fora_por_equipe: dict[str, int] = {}
    if 'saida' in df_completo.columns and 'micro_area' in df_completo.columns:
        df_sem_saida = df_completo[is_blank(df_completo['saida'])]
        mask_fora = df_sem_saida['micro_area'].str.contains('FORA', case=False, na=False)
        fora_por_equipe = df_sem_saida[mask_fora].groupby('equipe').size().to_dict()

    linhas_hist: list[dict] = []
    linhas_hist_ma: list[dict] = []
    linhas_estado: list[dict] = []
    linhas_faixa_etaria: list[dict] = []
    linhas_grupos: list[dict] = []
    linhas_condicoes: list[dict] = []
    listagens: dict[str, dict[str, pd.DataFrame]] = {}  # slug_equipe -> {indicador_id: df}
    equipes_slugs_processadas: set[str] = set()

    for eq_nome_raw, g in df_ativos.groupby('equipe', dropna=True):
        eq_cfg = buscar_equipe(municipio, eq_nome_raw, equipes_cfg)
        if eq_cfg is None:
            continue  # já avisado em validar_equipes
        slug = eq_cfg['slug']
        ine = eq_cfg.get('ine') or ''
        total = len(g)
        equipes_slugs_processadas.add(slug)

        if total == 0:
            avisos.append(f'ALERTA: total de cadastrados ativos zerado em {municipio} | {eq_cfg["nome_exibicao"]}.')
            continue

        # Nº de inconsistências ATIVAS (todas, não só as que entram no
        # score) por cadastro — base do "estado do cadastro" (Etapa 6):
        # % limpos, distribuição 0/1/2/3+, média.
        contagem_por_cadastro = pd.Series(0, index=g.index)

        for inc in inconsistencias_cfg:
            mask = avaliar_regra(inc['regra'], g, contexto)
            contagem_por_cadastro += mask.astype(int)
            valor = int(mask.sum())
            linhas_hist.append({
                'competencia': competencia, 'municipio': municipio, 'ine': ine,
                'equipe': slug, 'indicador_id': inc['id'], 'valor': valor, 'total_cadastrados': total,
            })

            if not inc.get('listagem', {}).get('gerar'):
                continue

            if inc['id'] == 'possivel_duplicata':
                # Listagem em grupos (Grupo/Motivo/alerta), igual ao comportamento
                # original de gerar_relatorio.py — usa o df COMPLETO da equipe
                # (com inativos), não só os ativos, para dar contexto ao grupo
                # (ex.: um dos dois cadastros do par já pode estar marcado como óbito).
                df_eq_completo = df_completo[
                    df_completo['equipe'].str.strip().str.upper() == eq_nome_raw.strip().upper()
                ]
                dup_eq = detectar_duplicatas(df_eq_completo)
                if not dup_eq.empty:
                    listagens.setdefault(slug, {})['possivel_duplicata'] = dup_eq
                continue

            colunas = [c for c in inc['listagem']['colunas'] if c in g.columns]
            subset = g.loc[mask, colunas].sort_values(
                [c for c in ['equipe', 'micro_area', 'nome'] if c in colunas] or colunas[:1]
            )
            listagens.setdefault(slug, {})[inc['id']] = subset

        fora_area = int(fora_por_equipe.get(eq_nome_raw, 0))
        linhas_estado.append({
            'competencia': competencia, 'municipio': municipio, 'ine': ine, 'equipe': slug,
            'total_cadastrados': total,
            'limpos': int((contagem_por_cadastro == 0).sum()),
            'com_1': int((contagem_por_cadastro == 1).sum()),
            'com_2': int((contagem_por_cadastro == 2).sum()),
            'com_3_mais': int((contagem_por_cadastro >= 3).sum()),
            'media_inconsistencias': round(float(contagem_por_cadastro.mean()), 3),
            'fora_area': fora_area,
            'total_bruto': total + fora_area,
        })

        # Sem CPF por faixa etária (formato longo: indicador_id sempre
        # 'sem_cpf' por enquanto, mas já preparado para outro indicador
        # por faixa no futuro sem precisar mudar o esquema do arquivo).
        if 'idade_num' in g.columns and 'cpf' in g.columns:
            mask_sem_cpf_geral = is_blank_cpf(g['cpf'])
            for faixa, ini, fim in FAIXAS_ETARIAS:
                mask_faixa = (g['idade_num'] >= ini) & (g['idade_num'] <= fim)
                linhas_faixa_etaria.append({
                    'competencia': competencia, 'municipio': municipio, 'ine': ine, 'equipe': slug,
                    'faixa': faixa, 'indicador_id': 'sem_cpf',
                    'valor': int((mask_faixa & mask_sem_cpf_geral).sum()),
                    'total_cadastrados': int(mask_faixa.sum()),
                })

        # Impacto em grupos prioritários (config/grupos.yaml).
        for grupo in grupos_cfg:
            mask_grupo = avaliar_regra(grupo['regra'], g, contexto)
            total_grupo = int(mask_grupo.sum())
            if total_grupo == 0:
                continue
            sem_cpf_g = is_blank_cpf(g['cpf']) if 'cpf' in g.columns else pd.Series(False, index=g.index)
            sem_cns_g = is_blank(g['cns']) if 'cns' in g.columns else pd.Series(False, index=g.index)
            sem_vinculo_g = (g['vinc_familia'].astype(str).str.strip() == 'NÃO') if 'vinc_familia' in g.columns else pd.Series(False, index=g.index)
            falha_g = sem_cpf_g | sem_cns_g | sem_vinculo_g
            for indicador_id, mask_metrica in [
                ('sem_cpf', sem_cpf_g), ('sem_cns', sem_cns_g), ('falha_identificacao_vinculo', falha_g),
            ]:
                linhas_grupos.append({
                    'competencia': competencia, 'municipio': municipio, 'ine': ine, 'equipe': slug,
                    'grupo_id': grupo['id'], 'indicador_id': indicador_id,
                    'valor': int((mask_grupo & mask_metrica).sum()), 'total_grupo': total_grupo,
                })

        # Condições de saúde + alerta de plausibilidade (config/condicoes_saude.yaml):
        # zero casos numa equipe grande o suficiente é estatisticamente
        # improvável e quase sempre indica campo não preenchido — vira
        # pergunta no resumo, nunca "tudo certo".
        if 'cond_saude' in g.columns:
            minimo_alerta = condicoes_cfg.get('minimo_cadastros_para_alerta', 300)
            for cond in condicoes_cfg.get('condicoes', []):
                mask_cond = avaliar_regra(cond['regra'], g, contexto)
                valor_cond = int(mask_cond.sum())
                linhas_condicoes.append({
                    'competencia': competencia, 'municipio': municipio, 'ine': ine, 'equipe': slug,
                    'condicao_id': cond['id'], 'valor': valor_cond, 'total_cadastrados': total,
                })

                taxa_min = cond.get('taxa_esperada_min_pct')
                if taxa_min is not None:
                    if total > 0 and (valor_cond / total * 100) < taxa_min:
                        avisos.append(
                            f'ALERTA DE PLAUSIBILIDADE: "{cond["nome"]}" em {eq_cfg["nome_exibicao"]} está em '
                            f'{valor_cond / total * 100:.2f}%, abaixo do esperado (≥{taxa_min}% — {cond.get("taxa_esperada_nota", "")}).'
                        )
                elif valor_cond == 0 and total >= minimo_alerta:
                    avisos.append(
                        f'ALERTA DE PLAUSIBILIDADE: ninguém em {eq_cfg["nome_exibicao"]} tem "{cond["nome"]}" '
                        f'registrado (0 de {total} cadastros) — confira se esse campo está sendo preenchido.'
                    )

        if 'micro_area' in g.columns:
            for ma_nome, g_ma in g.groupby('micro_area', dropna=False):
                ma_label = 'Sem microárea' if pd.isna(ma_nome) or not str(ma_nome).strip() else str(ma_nome).strip()
                total_ma = len(g_ma)
                for inc in inconsistencias_cfg:
                    mask_ma = avaliar_regra(inc['regra'], g_ma, contexto)
                    linhas_hist_ma.append({
                        'competencia': competencia, 'municipio': municipio, 'ine': ine, 'equipe': slug,
                        'microarea': ma_label, 'indicador_id': inc['id'],
                        'valor': int(mask_ma.sum()), 'total_cadastrados': total_ma,
                    })

    df_novas = pd.DataFrame(linhas_hist)
    df_novas_ma = pd.DataFrame(linhas_hist_ma)
    df_novas_estado = pd.DataFrame(linhas_estado)
    df_novas_faixa_etaria = pd.DataFrame(linhas_faixa_etaria)
    df_novas_grupos = pd.DataFrame(linhas_grupos)
    df_novas_condicoes = pd.DataFrame(linhas_condicoes)

    # Possível erro de extração: indicador em 0 em TODAS as equipes do município nesta competência.
    if not df_novas.empty:
        por_indicador = df_novas.groupby('indicador_id')['valor'].sum()
        for indicador_id, soma in por_indicador.items():
            if soma == 0:
                avisos.append(
                    f'POSSÍVEL ERRO DE EXTRAÇÃO: "{indicador_id}" está em 0 em TODAS as equipes de '
                    f'{municipio} em {competencia}.'
                )

    # ── Upsert em memória (para o resumo poder comparar com o mês anterior
    # já a partir do histórico atualizado) ──────────────────────────────
    historico_atualizado = upsert_historico(df_novas, competencia, municipio, equipes_slugs_processadas)
    gravar_historico(historico_atualizado)

    historico_estado_atualizado = upsert_historico_estado(df_novas_estado, competencia, municipio, equipes_slugs_processadas)
    gravar_historico_estado(historico_estado_atualizado)

    if not df_novas_faixa_etaria.empty:
        historico_faixa_atualizado = upsert_historico_faixa_etaria(df_novas_faixa_etaria, competencia, municipio, equipes_slugs_processadas)
        gravar_historico_faixa_etaria(historico_faixa_atualizado)

    if not df_novas_grupos.empty:
        historico_grupos_atualizado = upsert_historico_grupos(df_novas_grupos, competencia, municipio, equipes_slugs_processadas)
        gravar_historico_grupos(historico_grupos_atualizado)

    if not df_novas_condicoes.empty:
        historico_condicoes_atualizado = upsert_historico_condicoes(df_novas_condicoes, competencia, municipio, equipes_slugs_processadas)
        gravar_historico_condicoes(historico_condicoes_atualizado)

    tem_microarea = 'micro_area' in df_ativos.columns
    if tem_microarea:
        historico_ma_atualizado = upsert_historico_microarea(
            df_novas_ma, competencia, municipio, equipes_slugs_processadas,
        )
        gravar_historico_microarea(historico_ma_atualizado)
    else:
        avisos.append(f'ALERTA: exportação sem coluna de microárea — data/historico_microarea.csv não foi atualizado para {municipio} | {competencia}.')

    # ── Listagens por equipe ─────────────────────────────────────────────
    pasta_base = caminhos_cfg.get('pasta_base_drive') or ''
    if not pasta_base:
        pasta_base = str(RAIZ / 'saida' / '_teste_local')
        avisos.append(
            'ALERTA: config/caminhos.yaml não tem pasta_base_drive configurada — as listagens desta '
            f'rodada foram salvas em {pasta_base} (local, fora do Drive real).'
        )

    inc_por_id = {i['id']: i for i in inconsistencias_cfg}
    arquivos_gerados = []
    for slug, indicadores in listagens.items():
        eq_cfg = next(e for e in equipes_cfg if e['slug'] == slug)
        pasta_equipe = Path(pasta_base) / municipio / eq_cfg['nome_exibicao'] / competencia
        pasta_equipe.mkdir(parents=True, exist_ok=True)
        for indicador_id, df_listagem in indicadores.items():
            if df_listagem.empty:
                continue
            nome_inc = inc_por_id[indicador_id]['nome']
            caminho_saida = pasta_equipe / f'{_sanitizar_nome_arquivo(nome_inc)}.xlsx'
            grupo_col = 'Grupo' if indicador_id == 'possivel_duplicata' else ''
            _save_listing_xlsx(df_listagem, caminho_saida,
                                titulo=f'{nome_inc} — {eq_cfg["nome_exibicao"]} — {competencia}',
                                group_col=grupo_col)
            arquivos_gerados.append(caminho_saida)

    # ── Resumo ────────────────────────────────────────────────────────
    print('\n── Resumo ──────────────────────────────────────────────')
    comp_anterior = competencia_anterior(competencia)
    for slug in sorted(equipes_slugs_processadas):
        eq_cfg = next(e for e in equipes_cfg if e['slug'] == slug)
        atual = score(historico_atualizado, municipio, competencia, slug, inconsistencias_cfg)
        anterior = score(historico_atualizado, municipio, comp_anterior, slug, inconsistencias_cfg)

        if atual['score'] is None:
            continue
        linha = f'{eq_cfg["nome_exibicao"]:<20} score={atual["score"]:>6.2f}'
        if anterior['score'] is not None:
            variacao = atual['score'] - anterior['score']
            sinal = '+' if variacao >= 0 else ''
            linha += f'  ({sinal}{variacao:.2f} pp vs {comp_anterior})'
        if atual['dado_incompleto']:
            linha += f'  [dado incompleto: {", ".join(atual["indicadores_ausentes"])}]'
        if atual['score'] < metas['criticidade']['atencao']:
            linha += '  ⚠ CRÍTICA'
        elif atual['score'] < metas['meta_score']:
            linha += '  ⚠ Atenção'
        print(linha)

    if avisos:
        print('\nAlertas:')
        for a in avisos:
            print(f'  - {a}')

    print(f'\n{len(arquivos_gerados)} planilhas de listagem geradas em {pasta_base}.')
    print(f'✓ {HISTORICO_CSV}')
    print(f'✓ {HISTORICO_MICROAREA_CSV}')
    print(f'✓ {HISTORICO_ESTADO_CSV}')
    if not df_novas_faixa_etaria.empty:
        print(f'✓ {HISTORICO_FAIXA_ETARIA_CSV}')
    if not df_novas_grupos.empty:
        print(f'✓ {HISTORICO_GRUPOS_CSV}')
    if not df_novas_condicoes.empty:
        print(f'✓ {HISTORICO_CONDICOES_CSV}')

    return {'avisos': avisos, 'arquivos_gerados': arquivos_gerados}


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--arquivo', help='Caminho de uma exportação bruta.')
    ap.add_argument('--municipio', help='Nome do município (obrigatório com --arquivo).')
    ap.add_argument('--lote', help='Pasta com uma subpasta por município, cada uma com um .xlsx.')
    ap.add_argument('--competencia', required=True, help='AAAA-MM')
    args = ap.parse_args()

    if bool(args.arquivo) == bool(args.lote):
        ap.error('Use --arquivo + --municipio OU --lote (não os dois, nem nenhum).')

    if args.arquivo:
        if not args.municipio:
            ap.error('--arquivo exige --municipio.')
        processar_arquivo(args.arquivo, args.municipio, args.competencia)
    else:
        pasta = Path(args.lote)
        subpastas = [p for p in pasta.iterdir() if p.is_dir()]
        if not subpastas:
            print(f'Nenhuma subpasta de município encontrada em {pasta}.')
            sys.exit(1)
        for sub in subpastas:
            arquivos = list(sub.glob('*.xlsx'))
            if not arquivos:
                print(f'(nenhum .xlsx em {sub}, pulando)')
                continue
            for arq in arquivos:
                processar_arquivo(str(arq), sub.name, args.competencia)

    # Regera o site (site/data.json + as páginas HTML) UMA VEZ no final
    # (mesmo em --lote, com todos os municípios já upsertados) e roda a
    # verificação de privacidade antes de considerar a carga publicável.
    from pipeline.gerar_site import gerar as gerar_site
    print('\nRegenerando site/...')
    ok, violacoes = gerar_site()
    if not ok:
        print(f'✗ VERIFICAÇÃO DE PRIVACIDADE FALHOU ({len(violacoes)} problema(s)) — NÃO publique o site:')
        for v in violacoes:
            print(f'  - {v}')
        sys.exit(1)
    print('✓ site/ regenerado e aprovado na verificação de privacidade.')


if __name__ == '__main__':
    main()
