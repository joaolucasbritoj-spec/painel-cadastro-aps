"""
SISTEMA DE RELATÓRIO MENSAL - CADASTRO INDIVIDUAL (APS)
Uso: python gerar_relatorio.py "caminho/para/arquivo.xlsx"
     python gerar_relatorio.py  (abre seletor de arquivo)

Saídas:
  - Relatório TXT (console + arquivo)
  - Relatório DOCX (modelo institucional Prima Qualitá Saúde)
  - Planilhas Excel formatadas de irregulares por tipo (cabeçalhos coloridos)
"""

import pandas as pd
import sys
import os
import re
import io
import json
import subprocess
import shutil

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
from datetime import datetime
from pathlib import Path


# ── Mapeamento de colunas ────────────────────────────────────
RENAME = {
    'EQUIPE': 'equipe',
    'MICRO ÁREA': 'micro_area',
    'FUNCIONÁRIO': 'funcionario',
    'NOME DO PACIENTE': 'nome',
    'CPF DO PACIENTE': 'cpf',
    'CNS DO PACIENTE': 'cns',
    'DATA DE NASCIMENTO': 'nascimento',
    'SEXO DO PACIENTE': 'sexo',
    'IDADE DO PACIENTE': 'idade',
    'RAÇA/COR': 'raca',
    'TELEFONE CELULAR': 'cel',
    'TELEFONE RESIDENCIAL': 'tel_res',
    'LOGRADOURO': 'logradouro',
    'BAIRRO': 'bairro',
    'CEP': 'cep',
    'ORIGEM DA FICHA': 'origem',
    'DATA DE EDICAO': 'dt_edicao',
    'DATA DE CRIACAO': 'dt_criacao',
    'SAIDA DE CADASTRO DO CIDADAO': 'saida',
    'CONDICOES DE SAUDE': 'cond_saude',
    'Recebe algum auxílio do governo?': 'auxilio',
    'Qual benefício recebe?': 'beneficio',
    'SITUAÇÃO NO MERCADO DE TRABALHO': 'trabalho',
    'VINCULADO À FAMÍLIA': 'vinc_familia',
    'ESCOLARIDADE': 'escolaridade',
}

# ── Meses em português ──────────────────────────────────────────────────────
MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro',
}

# ── Ciclos de Vida (Ministério da Saúde / PNAB) ──────────────────────────────
CICLOS_VIDA = [
    ('Criança (0–9 anos)',         0,   9),
    ('Adolescente (10–19 anos)',  10,  19),
    ('Adulto Jovem (20–39 anos)', 20,  39),
    ('Adulto (40–59 anos)',       40,  59),
    ('Idoso (60–79 anos)',        60,  79),
    ('Longevo (≥80 anos)',        80, 999),
]


def _save_listing_xlsx(df: pd.DataFrame, path: Path, titulo: str = '', group_col: str = '', obs: str = '') -> None:
    """Salva DataFrame como planilha Excel com cabeçalhos coloridos e colunas já separadas."""
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        df.to_csv(path.with_suffix('.csv'), index=False, encoding='utf-8-sig')
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # type: ignore[arg-type]
    ws = wb.create_sheet('Listagem')
    ws.sheet_view.showGridLines = False

    fill_tit = PatternFill('solid', fgColor='1F497D')
    fill_hdr = PatternFill('solid', fgColor='2E74B5')
    fill_alt  = PatternFill('solid', fgColor='F2F2F2')
    fill_grp2 = PatternFill('solid', fgColor='E8F4FF')
    ft_tit  = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
    ft_hdr  = Font(name='Calibri', bold=True, size=9,  color='FFFFFF')
    ft_corp = Font(name='Calibri', size=9)
    brd = Border(
        left=Side('thin', color='D0D0D0'), right=Side('thin', color='D0D0D0'),
        top=Side('thin', color='D0D0D0'), bottom=Side('thin', color='D0D0D0'),
    )
    AC = Alignment(horizontal='center', vertical='center', wrap_text=True)
    AL = Alignment(horizontal='left',   vertical='center')

    ncols = max(len(df.columns), 1)
    start_row = 1
    if titulo:
        c = ws.cell(row=1, column=1, value=titulo)
        c.font = ft_tit; c.fill = fill_tit; c.alignment = AC; c.border = brd
        ws.row_dimensions[1].height = 22
        if ncols > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        start_row = 2

    if obs:
        c = ws.cell(row=start_row, column=1, value=obs)
        c.font = Font(name='Calibri', italic=True, size=9, color='FFFFFF')
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        c.border = brd
        ws.row_dimensions[start_row].height = 30
        if ncols > 1:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=ncols)
        start_row += 1

    LABEL_MAP = {
        'equipe': 'Equipe', 'micro_area': 'Micro Área', 'funcionario': 'ACS/Funcionário',
        'nome': 'Nome do Paciente', 'cpf': 'CPF', 'cns': 'CNS',
        'nascimento': 'Data Nascimento', 'idade': 'Idade', 'cel': 'Celular',
        'tel_res': 'Tel. Residencial', 'dt_edicao': 'Última Edição',
        'dt_criacao': 'Data de Criação', 'saida': 'Saída de Cadastro',
        'condicoes_registradas': 'Condições Registradas',
        'alerta_obito_mudanca': 'Alerta Óbito/Mudança',
    }
    hdrs = [LABEL_MAP.get(col, col.replace('_', ' ').title()) for col in df.columns]
    ws.row_dimensions[start_row].height = 18
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=start_row, column=ci, value=h)
        c.font = ft_hdr; c.fill = fill_hdr; c.alignment = AC; c.border = brd

    _grp_toggle = False
    _prev_grp: object = object()
    for ri, (_, row_data) in enumerate(df.iterrows(), start_row + 1):
        ws.row_dimensions[ri].height = 14
        if group_col and group_col in df.columns:
            cur_grp = row_data.get(group_col)
            if cur_grp != _prev_grp:
                _grp_toggle = not _grp_toggle
                _prev_grp = cur_grp
            bg = fill_grp2 if _grp_toggle else None
        else:
            bg = fill_alt if (ri - start_row) % 2 == 0 else None
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=('' if pd.isna(val) else val))  # type: ignore[arg-type]
            cell.font = ft_corp; cell.alignment = AL; cell.border = brd
            if bg:
                cell.fill = bg

    for ci, col in enumerate(df.columns, 1):
        vals = [str(hdrs[ci - 1])] + [str(v) for v in df[col].fillna('')]
        width = min(max((len(v) for v in vals), default=8) + 2, 50)
        ws.column_dimensions[get_column_letter(ci)].width = width

    ws.freeze_panes = f'A{start_row + 1}'
    wb.save(path)


def is_blank(s: pd.Series) -> pd.Series:
    stripped = s.str.strip()
    return s.isna() | (stripped == '') | (stripped.str.lower() == 'nan')


def is_blank_cpf(s: pd.Series) -> pd.Series:
    """Detecta CPF ausente, incluindo placeholders como '0', '0.0', '00000000000'."""
    stripped = s.str.strip()
    digits_only = stripped.str.replace(r'[\.\-\s]', '', regex=True)
    return (
        is_blank(s)
        | (digits_only == '0')
        | (digits_only == '')
        | digits_only.str.fullmatch(r'0+', na=False)
    )


def parse_age(s) -> int | None:
    if pd.isna(s):
        return None
    m = re.match(r'(\d+)\s+ano', str(s))
    return int(m.group(1)) if m else None


def classificar_risco(cond_str, idade_num=None) -> str:
    """Estratifica o risco individual com base nas condições registradas.

    Critérios (baseados em PNAB / diretrizes clínicas de APS):
    - Alto Risco: HAS+DM combinados, doença cardiovascular, renal crônica,
                  neoplasia, acamado, ou ≥2 condições crônicas simultâneas.
    - Médio Risco: condição crônica isolada, gestante, uso de álcool/tabaco,
                   obesidade/sobrepeso, domiciliado, ou longevo (≥80 anos).
    - Baixo Risco: sem condições crônicas registradas.
    """
    cond = '' if pd.isna(cond_str) else str(cond_str).upper()
    tem_has      = 'HIPERTENSO' in cond
    tem_dm       = 'DIABETES'   in cond
    tem_card     = bool(re.search(r'CARD[IÍ]ACA|CARDIOVASC', cond))
    tem_renal    = 'RENAL'      in cond
    tem_cancer   = bool(re.search(r'C[AÂ]NCER', cond))
    tem_mental   = 'MENTAL'     in cond
    tem_resp     = 'RESPIRAT'   in cond
    tem_acamado  = 'ACAMADO'    in cond
    tem_obeso    = bool(re.search(r'OBES[OA]|OBESIDADE|SOBREPESO', cond))
    tem_domicil  = bool(re.search(r'DOMICILI[AO]', cond))
    n_cronicas = sum([tem_has, tem_dm, tem_card, tem_renal, tem_cancer, tem_mental, tem_resp])
    # Acamado eleva para alto risco — exige cuidado domiciliar intensivo
    if (tem_has and tem_dm) or tem_card or tem_renal or tem_cancer or tem_acamado or n_cronicas >= 2:
        return 'Alto Risco'
    if n_cronicas == 1 or 'GESTANTE' in cond:
        return 'Médio Risco'
    if bool(re.search(r'[AÁ]LCOOL', cond)) or 'FUMANTE' in cond:
        return 'Médio Risco'
    # Obesidade e domiciliado elevam para médio risco
    if tem_obeso or tem_domicil:
        return 'Médio Risco'
    if isinstance(idade_num, (int, float)) and not pd.isna(idade_num) and idade_num >= 80:
        return 'Médio Risco'
    return 'Baixo Risco'


def detect_municipio(filepath: str) -> str:
    name = Path(filepath).stem
    parts = name.split(' - ')
    return parts[-1].strip().title() if len(parts) > 1 else name


def load_data(filepath: str) -> pd.DataFrame:
    import openpyxl.reader.excel as _oxl
    _orig_stylesheet = _oxl.apply_stylesheet

    def _safe_stylesheet(archive, wb):
        try:
            _orig_stylesheet(archive, wb)
        except ValueError:
            pass  # xlsx sometimes has invalid color hex values in stylesheet

    _oxl.apply_stylesheet = _safe_stylesheet
    try:
        df = pd.read_excel(filepath, dtype=str)
    finally:
        _oxl.apply_stylesheet = _orig_stylesheet
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns={k: v for k, v in RENAME.items() if k in df.columns})
    if 'equipe' in df.columns:
        df['equipe'] = df['equipe'].str.strip().replace('', pd.NA)
    df['idade_num'] = df['idade'].apply(parse_age) if 'idade' in df.columns else None
    if 'dt_edicao' in df.columns:
        df['dt_edicao_p'] = pd.to_datetime(df['dt_edicao'], dayfirst=True, errors='coerce')
    return df


def fmt_n(n: int) -> str:
    """Formata número com separador de milhar no padrão BR (ponto)."""
    return f"{n:,}".replace(',', '.')


def fmt_pct(n: int, total: int) -> str:
    """Formata percentual no padrão BR."""
    return f"{n/total*100:.1f}%".replace('.', ',')


def fmt_pct_int(n: int, total: int) -> str:
    """Formata percentual inteiro."""
    return f"{n/total*100:.0f}%"


# ══════════════════════════════════════════════════════════════════════════════
#  EXTRAÇÃO DE DADOS ESTRUTURADOS
# ══════════════════════════════════════════════════════════════════════════════
def extrair_dados(df: pd.DataFrame, municipio: str, filepath: str, hoje: datetime) -> dict:
    """Extrai todos os dados do DataFrame em um dicionário estruturado.

    A partir da seção "Visão geral", todas as demais análises consideram
    apenas os cadastros ativos (exclui óbito e mudança de território) — quem
    saiu do território não deve compor nenhuma estatística do relatório. A
    única exceção é a checagem de possíveis duplicatas, que usa
    ``df_completo`` (todos os registros, inclusive inativos), pois um par
    duplicado pode ter uma das entradas já marcada como óbito/mudança.
    """
    total = len(df)
    mes_ref_pt = f"{MESES_PT[hoje.month]}/{hoje.year}"

    dados = {
        'municipio': municipio,
        'uf': 'RJ',
        'competencia': mes_ref_pt,
        'data_geracao': hoje.strftime('%d/%m/%Y'),
        'hora_geracao': hoje.strftime('%H:%M'),
        'fonte': Path(filepath).name,
        'total': total,
    }

    # ── Visão geral ──────────────────────────────────────────────────────
    df_completo = df  # conjunto completo, preservado só para a checagem de duplicatas
    if 'saida' in df.columns:
        com_saida = ~is_blank(df['saida'])
        n_ativos = (~com_saida).sum()
        n_saida = com_saida.sum()
        n_obito = df['saida'].str.contains('BITO', na=False).sum()
        n_mudanca = df['saida'].str.contains('MUDAN', na=False).sum()
        n_om = df['saida'].str.contains(r'BITO|MUDAN', na=False, regex=True).sum()
        n_outras = int(n_saida - n_om)
        dados['visao_geral'] = {
            'ativos': int(n_ativos),
            'ativos_pct': fmt_pct(n_ativos, total),
            'saida': int(n_saida),
            'saida_pct': fmt_pct(n_saida, total),
            'obitos': int(n_obito),
            'mudanca': int(n_mudanca),
            'outras_saidas': n_outras,
        }
        # Qualquer saída de cadastro (óbito, mudança de território, recusa
        # etc.) tira a pessoa de todas as demais análises — não há sentido
        # em cobrar atualização de quem não está mais na área. A única
        # exceção é a checagem de duplicatas, que usa df_completo (todos os
        # registros) porque o par duplicado pode ter uma entrada já marcada
        # com saída e outra ainda ativa numa microárea.
        df = df[is_blank(df['saida'])].copy()
        total = len(df)

    # "Fora de área": a MICRO ÁREA aponta que o cadastro não pertence a
    # nenhuma área de cobertura real (não é registrado como saída de
    # cadastro, mas ninguém é responsável por atualizar esse cadastro) —
    # mesma exceção de duplicatas: usa df_completo, preservado acima.
    if 'micro_area' in df.columns:
        mask_fora = df['micro_area'].str.contains('FORA', case=False, na=False)
        n_fora_area = int(mask_fora.sum())
        if 'visao_geral' in dados:
            dados['visao_geral']['fora_area'] = n_fora_area
        df = df[~mask_fora].copy()
        total = len(df)

    # ── Sexo ─────────────────────────────────────────────────────────────
    if 'sexo' in df.columns:
        dist_sexo = []
        for v, n in df['sexo'].value_counts(dropna=False).items():
            label = str(v).strip()
            if label.lower() == 'nan' or pd.isna(v):
                label = 'Sem informação'
            else:
                label = label.title()
            dist_sexo.append({'label': label, 'n': int(n), 'pct': fmt_pct(n, total)})
        dados['sexo'] = dist_sexo

    # ── Raça/Cor ─────────────────────────────────────────────────────────
    if 'raca' in df.columns:
        dist_raca = []
        for v, n in df['raca'].value_counts(dropna=False).items():
            label = str(v).strip()
            if label.lower() == 'nan' or pd.isna(v):
                label = 'Sem informação'
            else:
                label = label.title()
            dist_raca.append({'label': label, 'n': int(n), 'pct': fmt_pct(n, total)})
        dados['raca'] = dist_raca

    # ── Equipes ──────────────────────────────────────────────────────────
    if 'equipe' in df.columns:
        equipes = []
        t_sc = t_sn = t_se = t_vf = t_de = 0
        for eq, g in sorted(df.groupby('equipe', dropna=False), key=lambda x: (pd.isna(x[0]), str(x[0]))):
            n = len(g)
            sc = int(is_blank_cpf(g['cpf']).sum()) if 'cpf' in g else 0
            sn = int(is_blank(g['cns']).sum()) if 'cns' in g else 0
            se = int(is_blank(g['logradouro']).sum()) if 'logradouro' in g else 0
            vf = int((g['vinc_familia'].str.strip() == 'NÃO').sum()) if 'vinc_familia' in g else 0
            de = int((g['dt_edicao_p'] < pd.Timestamp('2025-04-27')).sum()) if 'dt_edicao_p' in g else 0
            t_sc += sc; t_sn += sn; t_se += se; t_vf += vf; t_de += de
            equipes.append({
                'nome': 'Sem equipe' if pd.isna(eq) else str(eq).title(),
                'total': n,
                's_cpf': sc, 's_cpf_pct': fmt_pct_int(sc, n),
                's_cns': sn, 's_cns_pct': fmt_pct_int(sn, n),
                's_end': se, 's_end_pct': fmt_pct_int(se, n),
                's_vf': vf, 's_vf_pct': fmt_pct_int(vf, n),
                'desatual': de, 'desatual_pct': fmt_pct_int(de, n),
            })
        # Linha de totais
        equipes.append({
            'nome': 'TOTAL',
            'total': total,
            's_cpf': t_sc, 's_cpf_pct': fmt_pct_int(t_sc, total),
            's_cns': t_sn, 's_cns_pct': fmt_pct_int(t_sn, total),
            's_end': t_se, 's_end_pct': fmt_pct_int(t_se, total),
            's_vf': t_vf, 's_vf_pct': fmt_pct_int(t_vf, total),
            'desatual': t_de, 'desatual_pct': fmt_pct_int(t_de, total),
        })
        dados['equipes'] = equipes

    # ── Grupos prioritários ──────────────────────────────────────────────
    grupos = {}
    if 'idade_num' in df.columns:
        idosos_n = int((df['idade_num'] >= 60).sum())
        criancas_n = int((df['idade_num'] <= 4).sum())
        criancas_02_n = int(((df['idade_num'] >= 0) & (df['idade_num'] <= 2)).sum())
        criancas_05_n = int(((df['idade_num'] >= 0) & (df['idade_num'] <= 5)).sum())
        grupos['idosos'] = {'n': idosos_n, 'pct': fmt_pct(idosos_n, total)}
        grupos['criancas'] = {'n': criancas_n, 'pct': fmt_pct(criancas_n, total)}
        grupos['criancas_02'] = {'n': criancas_02_n, 'pct': fmt_pct(criancas_02_n, total)}
        grupos['criancas_05'] = {'n': criancas_05_n, 'pct': fmt_pct(criancas_05_n, total)}

    if 'auxilio' in df.columns:
        _nao = {'não', 'nao', 'n', 'não possui', 'nao possui'}
        _sem = is_blank(df['auxilio'])
        _tem = ~_sem & ~df['auxilio'].str.strip().str.lower().isin(_nao)
        n_com = int(_tem.sum())
        n_nao_inf = int((df['auxilio'].str.strip().str.upper() == 'NÃO INFORMADO').sum())

        dist_beneficio = []
        if 'beneficio' in df.columns:
            for v, n in df['beneficio'].value_counts(dropna=True).items():
                label = str(v).strip()
                if label.lower() == 'nan':
                    continue
                dist_beneficio.append({'label': label.title(), 'n': int(n), 'pct': fmt_pct(n, total)})

        grupos['auxilio'] = {
            'n_com': n_com,
            'pct_com': fmt_pct(n_com, total),
            'n_nao_inf': n_nao_inf,
            'pct_nao_inf': fmt_pct(n_nao_inf, total),
            'beneficios': dist_beneficio,
        }

    if 'cond_saude' in df.columns:
        gestantes_n = int(df['cond_saude'].str.contains('GESTANTE', na=False).sum())
        grupos['gestantes'] = {'n': gestantes_n}

        # Gestantes por equipe
        mask_gest = df['cond_saude'].str.contains('GESTANTE', na=False)
        if 'equipe' in df.columns:
            g_eq = df[mask_gest]['equipe'].value_counts(dropna=False)
            grupos['gestantes_equipe'] = [
                {'equipe': 'Sem equipe' if pd.isna(eq) else str(eq).title(), 'n': int(n)} for eq, n in g_eq.items()
            ]

        # Listagem nominal de gestantes (para aba Excel do relatório)
        _cols_gest = [c for c in ['equipe', 'micro_area', 'funcionario', 'nome', 'cpf', 'cns',
                                   'nascimento', 'idade', 'cel', 'tel_res', 'cond_saude']
                      if c in df.columns]
        _ord_gest = [c for c in ['equipe', 'micro_area', 'nome'] if c in df.columns] or ['nome']
        dados['gestantes_listing'] = (
            df[mask_gest][_cols_gest]
            .sort_values(_ord_gest)
            .fillna('')
            .astype(str)
            .to_dict('records')
        )

        # Gestantes de alto risco
        _pat_ar = r'HIPERTENSO|DIABETES|RENAL|CARD[IÍ]AC|CARDIOVASC|C[AÂ]NCER|AVC|DERRAME|INFARTO|IAM'
        mask_gest_ar = mask_gest & df['cond_saude'].str.contains(_pat_ar, na=False)
        n_gest_ar = int(mask_gest_ar.sum())
        grupos['gestantes_alto_risco'] = {
            'n': n_gest_ar,
            'pct_gestantes': fmt_pct(n_gest_ar, gestantes_n) if gestantes_n else '0,0%',
        }
        if 'equipe' in df.columns and n_gest_ar > 0:
            gar_eq = df[mask_gest_ar]['equipe'].value_counts(dropna=False)
            grupos['gestantes_alto_risco']['por_equipe'] = [
                {'equipe': 'Sem equipe' if pd.isna(eq) else str(eq).title(), 'n': int(n)} for eq, n in gar_eq.items()
            ]

        # Puérperas
        mask_puerp = df['cond_saude'].str.contains(r'PU[EÉ]RPERA', na=False)
        n_puerp = int(mask_puerp.sum())
        grupos['puerperas'] = {'n': n_puerp}
        if 'equipe' in df.columns and n_puerp > 0:
            p_eq = df[mask_puerp]['equipe'].value_counts(dropna=False)
            grupos['puerperas_equipe'] = [
                {'equipe': 'Sem equipe' if pd.isna(eq) else str(eq).title(), 'n': int(n)} for eq, n in p_eq.items()
            ]

        # Condições de saúde
        conds_map = {
            'Hipertensão (HAS)':                     'HIPERTENSO',
            'Diabetes (DM)':                          'DIABETES',
            'HAS + DM':                               r'HIPERTENSO.*DIABETES|DIABETES.*HIPERTENSO',
            'Usuário de álcool':                      r'ALCOOL|ÁLCOOL',
            'Uso de drogas':                           r'DROGA|SUBSTANC',
            'Fumante':                                 'FUMANTE',
            'Doença respiratória':                     'RESPIRAT',
            'Doença cardíaca':                         r'CARD.ACA|CARDIACA',
            'Doença renal':                             'RENAL',
            'Doença mental':                           'MENTAL',
            'AVC / Derrame':                            r'AVC|DERRAME|ACIDENTE VASCULAR',
            'Infarto':                                  r'INFARTO|IAM',
            'Câncer':                                   r'CANCER|CÂNCER',
            'Hanseníase':                                r'HANSEN',
            'Tuberculose':                               r'TUBERCULOS[EI]',
            'Doença falciforme':                         'FALCIFORME',
            'Internação nos últimos 12 meses':           r'INTERNA[CÇ][AÃ]O',
            'Usa plantas medicinais':                    r'PLANTA MEDICINAL|FITOTERAP',
            'Práticas integrativas e complementares':    r'PRATICA INTEGR|PRÁTICA INTEGR|PICS',
            'Acamado':                                   'ACAMADO',
            'Domiciliado':                               r'DOMICILI[AO]',
        }
        conds = []
        for label, pat in conds_map.items():
            n = int(df['cond_saude'].str.contains(pat, na=False).sum())
            conds.append({'label': label, 'n': n, 'pct': fmt_pct(n, total)})
        grupos['condicoes'] = conds

        # Nutrição infantil (crianças 0–5 anos)
        if 'idade_num' in df.columns:
            mask_c05 = (df['idade_num'] >= 0) & (df['idade_num'] <= 5)
            n_c05 = int(mask_c05.sum())
            mask_bp = df['cond_saude'].str.contains('BAIXO PESO', na=False)
            mask_ob = df['cond_saude'].str.contains(r'OBES[OA]|OBESIDADE', na=False)
            mask_sp = df['cond_saude'].str.contains('SOBREPESO', na=False)
            mask_qualquer_peso = mask_bp | mask_ob | mask_sp
            n_bp = int((mask_c05 & mask_bp).sum())
            n_ob = int((mask_c05 & mask_ob).sum())
            n_sp = int((mask_c05 & mask_sp).sum())
            n_sem = int((mask_c05 & ~mask_qualquer_peso).sum())
            grupos['criancas_nutricao'] = {
                'baixo_peso':   {'n': n_bp, 'pct': fmt_pct(n_bp,  n_c05) if n_c05 else '0,0%'},
                'obesas':       {'n': n_ob, 'pct': fmt_pct(n_ob,  n_c05) if n_c05 else '0,0%'},
                'sobrepeso':    {'n': n_sp, 'pct': fmt_pct(n_sp,  n_c05) if n_c05 else '0,0%'},
                'sem_info':     {'n': n_sem,'pct': fmt_pct(n_sem, n_c05) if n_c05 else '0,0%'},
            }

        # Acamados e domiciliados — grupo especial com necessidade de visita domiciliar
        mask_acam = df['cond_saude'].str.contains('ACAMADO', na=False)
        mask_dom  = df['cond_saude'].str.contains(r'DOMICILI[AO]', na=False)
        n_acam = int(mask_acam.sum())
        n_dom  = int(mask_dom.sum())
        if n_acam > 0 or n_dom > 0:
            aten_dom = {'acamados': n_acam, 'domiciliados': n_dom,
                        'acamados_pct': fmt_pct(n_acam, total),
                        'domiciliados_pct': fmt_pct(n_dom, total)}
            if 'equipe' in df.columns:
                acam_eq = df[mask_acam | mask_dom].groupby('equipe', dropna=False).size()
                aten_dom['por_equipe'] = [
                    {'equipe': 'Sem equipe' if pd.isna(eq) else str(eq).title(), 'n': int(n)} for eq, n in acam_eq.items()
                ]
            if 'micro_area' in df.columns:
                acam_ma = df[mask_acam | mask_dom].groupby('micro_area').size()
                aten_dom['por_microarea'] = [
                    {'microarea': str(ma).strip(), 'n': int(n)} for ma, n in acam_ma.items()
                ]
            grupos['atencao_domiciliar'] = aten_dom

    dados['grupos'] = grupos

    # ── Ciclos de vida (apenas cadastros ativos, sem óbitos/mudanças) ──────────
    if 'idade_num' in df.columns:
        _n_ativos_cv = len(df)
        sem_idade = int(df['idade_num'].isna().sum())
        df_ci = df.dropna(subset=['idade_num'])
        total_ci = len(df_ci)
        ciclos = []
        for label, ini, fim in CICLOS_VIDA:
            mask = (df_ci['idade_num'] >= ini) & (df_ci['idade_num'] <= fim)
            n = int(mask.sum())
            por_sexo = {}
            if 'sexo' in df_ci.columns:
                for sv, g in df_ci[mask].groupby('sexo'):
                    sl = str(sv).strip().title()
                    por_sexo[sl] = len(g)
            ciclos.append({
                'label': label, 'n': n,
                'pct': fmt_pct(n, _n_ativos_cv) if _n_ativos_cv > 0 else '0,0%',
                'pct_validos': fmt_pct(n, total_ci) if total_ci > 0 else '0,0%',
                'por_sexo': por_sexo,
            })
        dados['ciclos_vida'] = {
            'ciclos': ciclos,
            'sem_idade': sem_idade,
            'sem_idade_pct': fmt_pct(sem_idade, _n_ativos_cv),
        }

    # ── Estratificação de risco ──────────────────────────────────────────────
    if 'cond_saude' in df.columns:
        df = df.copy()
        df['_risco'] = df.apply(
            lambda r: classificar_risco(r['cond_saude'], r.get('idade_num')), axis=1
        )
        risco_items = []
        for nivel in ['Alto Risco', 'Médio Risco', 'Baixo Risco']:
            mask = df['_risco'] == nivel
            n = int(mask.sum())
            por_equipe = {}
            if 'equipe' in df.columns:
                for eq, g in df[mask].groupby('equipe', dropna=False):
                    por_equipe['Sem equipe' if pd.isna(eq) else str(eq).title()] = len(g)
            risco_items.append({
                'nivel': nivel, 'n': n, 'pct': fmt_pct(n, total),
                'por_equipe': por_equipe,
            })
        dados['estratificacao_risco'] = risco_items

    # ── Saúde da mulher ──────────────────────────────────────────────────────
    if 'sexo' in df.columns and 'idade_num' in df.columns:
        mask_fem = df['sexo'].str.strip().str.upper().isin(['FEMININO', 'F'])
        df_fem = df[mask_fem]
        n_fem = len(df_fem)
        sm = {'total': n_fem, 'pct': fmt_pct(n_fem, total)}
        mask_reprod = (df_fem['idade_num'] >= 10) & (df_fem['idade_num'] <= 49)
        n_reprod = int(mask_reprod.sum())
        sm['faixa_reprodutiva'] = {'n': n_reprod,
                                   'pct': fmt_pct(n_reprod, n_fem) if n_fem else '0,0%'}
        n_idosa = int((df_fem['idade_num'] >= 60).sum())
        sm['idosas'] = {'n': n_idosa, 'pct': fmt_pct(n_idosa, n_fem) if n_fem else '0,0%'}
        n_crianca_f = int(((df_fem['idade_num'] >= 0) & (df_fem['idade_num'] <= 9)).sum())
        sm['criancas_f'] = {'n': n_crianca_f, 'pct': fmt_pct(n_crianca_f, n_fem) if n_fem else '0,0%'}

        # Citopatológico (25–64 anos) e Mamografia (50–69 anos)
        n_cito = int(((df_fem['idade_num'] >= 25) & (df_fem['idade_num'] <= 64)).sum())
        n_mamo = int(((df_fem['idade_num'] >= 50) & (df_fem['idade_num'] <= 69)).sum())
        sm['cito']      = {'n': n_cito, 'pct': fmt_pct(n_cito, n_fem) if n_fem else '0,0%'}
        sm['mamografia']= {'n': n_mamo, 'pct': fmt_pct(n_mamo, n_fem) if n_fem else '0,0%'}

        if 'cond_saude' in df.columns:
            n_gest = int(df_fem['cond_saude'].str.contains('GESTANTE', na=False).sum())
            taxa_obs = n_gest / total * 100 if total else 0
            sm['gestantes'] = {
                'n': n_gest,
                'taxa_observada_pct': fmt_pct(n_gest, total),
                'taxa_esperada_nota': 'Estimativa: ~1,5–2% da pop. total (DATASUS/SINASC)',
                'alerta': taxa_obs < 1.0,
            }
        dados['saude_mulher'] = sm

    # ── Perfil completo por equipe ───────────────────────────────────────────
    if 'equipe' in df.columns:
        _conds_perfil = {
            'Hipertensão (HAS)':     'HIPERTENSO',
            'Diabetes (DM)':         'DIABETES',
            'HAS + DM':              r'HIPERTENSO.*DIABETES|DIABETES.*HIPERTENSO',
            'Usuário de álcool':     r'ALCOOL|ÁLCOOL',
            'Fumante':               'FUMANTE',
            'Doença respiratória':   'RESPIRAT',
            'Doença cardíaca':       r'CARD.ACA|CARDIACA',
            'Doença renal':          'RENAL',
            'Doença mental':         'MENTAL',
            'Câncer':                r'CANCER|CÂNCER',
            'AVC / Derrame':         r'AVC|DERRAME|ACIDENTE VASCULAR',
            'Infarto':               r'INFARTO|IAM',
            'Tuberculose':           r'TUBERCULOS[EI]',
            'Hanseníase':            r'HANSEN|HANSENI',
            'Drogas Ilícitas':       r'DROGA|SUBSTANC',
            'Plantas Medicinais':    r'PLANTA MEDICINAL|FITOTERAP',
            'Práticas Integrativas': r'PRATICA INTEGR|PRÁTICA INTEGR|PICS',
        }
        _risco_por_eq = {}
        for ri in dados.get('estratificacao_risco', []):
            for eq_n, n in ri.get('por_equipe', {}).items():
                _risco_por_eq.setdefault(eq_n, {})[ri['nivel']] = n

        equipe_perfil = {}
        for eq, g in df.groupby('equipe', dropna=False):
            eq_name = 'Sem equipe' if pd.isna(eq) else str(eq).title()
            perfil = {}

            if 'sexo' in g.columns:
                perfil['sexo'] = {
                    (str(sv).strip().title() if not pd.isna(sv) else 'Sem informação'): int(cnt)
                    for sv, cnt in g['sexo'].value_counts(dropna=False).items()
                }

            if 'raca' in g.columns:
                perfil['raca'] = {
                    (str(rv).strip().title() if not pd.isna(rv) else 'Sem informação'): int(cnt)
                    for rv, cnt in g['raca'].value_counts(dropna=False).items()
                }

            if 'idade_num' in g.columns:
                g_at = g  # já filtrado (df é apenas ativos a partir daqui)
                ciclos_eq = []
                for label, ini, fim in CICLOS_VIDA:
                    mask = (g_at['idade_num'] >= ini) & (g_at['idade_num'] <= fim)
                    item = {'label': label, 'n': int(mask.sum())}
                    if 'sexo' in g_at.columns:
                        for sv, gv in g_at[mask].groupby('sexo'):
                            item[str(sv).strip().title()] = len(gv)
                    ciclos_eq.append(item)
                perfil['ciclos']   = ciclos_eq
                perfil['idosos']   = int((g_at['idade_num'] >= 60).sum())
                perfil['longevos'] = int((g_at['idade_num'] >= 80).sum())
                perfil['criancas'] = int((g_at['idade_num'] <= 4).sum())

            if 'sexo' in g.columns and 'idade_num' in g.columns:
                mask_fem = g['sexo'].str.strip().str.upper().isin(['FEMININO', 'F'])
                gf = g[mask_fem]
                perfil['mulheres']        = len(gf)
                perfil['mulheres_reprod'] = int(((gf['idade_num'] >= 10) & (gf['idade_num'] <= 49)).sum())
                perfil['mulheres_idosas'] = int((gf['idade_num'] >= 60).sum())

            if 'cond_saude' in g.columns:
                perfil['condicoes'] = [
                    {'label': lbl, 'n': int(g['cond_saude'].str.contains(pat, na=False).sum())}
                    for lbl, pat in _conds_perfil.items()
                ]
                perfil['gestantes'] = int(g['cond_saude'].str.contains('GESTANTE', na=False).sum())

            perfil['risco'] = _risco_por_eq.get(eq_name, {})
            equipe_perfil[eq_name] = perfil

        dados['equipe_perfil'] = equipe_perfil

    # ── Análise por Microárea ────────────────────────────────────────────────
    if 'micro_area' in df.columns:
        microareas = []
        for ma, g in df.groupby('micro_area', dropna=False):
            if pd.isna(ma) or str(ma).strip() == '':
                ma_label = 'Sem microárea'
            else:
                ma_label = str(ma).strip()
            n = len(g)
            sc = int(is_blank_cpf(g['cpf']).sum())     if 'cpf'       in g.columns else 0
            sn = int(is_blank(g['cns']).sum())     if 'cns'       in g.columns else 0
            se = int(is_blank(g['logradouro']).sum()) if 'logradouro' in g.columns else 0
            de = int((g['dt_edicao_p'] < pd.Timestamp('2025-04-27')).sum()) if 'dt_edicao_p' in g.columns else 0
            acs = ''
            if 'funcionario' in g.columns:
                vc = g['funcionario'].value_counts()
                acs = str(vc.index[0]).title() if len(vc) > 0 else ''
            equipe_ma = ''
            if 'equipe' in g.columns:
                vc2 = g['equipe'].value_counts()
                equipe_ma = str(vc2.index[0]).title() if len(vc2) > 0 else ''
            microareas.append({
                'microarea': ma_label, 'equipe': equipe_ma, 'acs': acs,
                'total': n,
                's_cpf': sc, 's_cpf_pct': fmt_pct_int(sc, n),
                's_cns': sn, 's_cns_pct': fmt_pct_int(sn, n),
                's_end': se, 's_end_pct': fmt_pct_int(se, n),
                'desatual': de, 'desatual_pct': fmt_pct_int(de, n),
                'score': sc + sn + se,
            })
        dados['microareas'] = sorted(microareas, key=lambda x: x['score'], reverse=True)

    # ── Análise por ACS ──────────────────────────────────────────────────────
    if 'funcionario' in df.columns:
        acs_list = []
        for acs, g in df.groupby('funcionario', dropna=True):
            if pd.isna(acs) or str(acs).strip() == '':
                continue
            n = len(g)
            sc = int(is_blank_cpf(g['cpf']).sum())  if 'cpf' in g.columns else 0
            sn = int(is_blank(g['cns']).sum())  if 'cns' in g.columns else 0
            de = int((g['dt_edicao_p'] < pd.Timestamp('2025-04-27')).sum()) if 'dt_edicao_p' in g.columns else 0
            equipe_a = ''
            if 'equipe' in g.columns:
                vc = g['equipe'].value_counts()
                equipe_a = str(vc.index[0]).title() if len(vc) else ''
            ma_a = ''
            if 'micro_area' in g.columns:
                vc2 = g['micro_area'].value_counts()
                ma_a = str(vc2.index[0]).strip() if len(vc2) else ''
            acs_list.append({
                'nome': str(acs).title(), 'equipe': equipe_a, 'microarea': ma_a,
                'total': n,
                's_cpf': sc, 's_cpf_pct': fmt_pct_int(sc, n),
                's_cns': sn, 's_cns_pct': fmt_pct_int(sn, n),
                'desatual': de, 'desatual_pct': fmt_pct_int(de, n),
            })
        dados['acs'] = sorted(acs_list, key=lambda x: x['total'], reverse=True)

    # ── Análise temporal de cadastros ────────────────────────────────────────
    if 'dt_criacao' in df.columns:
        df['dt_criacao_p'] = pd.to_datetime(df['dt_criacao'], dayfirst=True, errors='coerce')
        por_ano = df['dt_criacao_p'].dt.year.value_counts().sort_index()
        temp = {
            'por_ano': [{'ano': int(a), 'n': int(n)} for a, n in por_ano.items()
                        if not pd.isna(a)],
        }
        if 'dt_edicao_p' in df.columns:
            mesma_data = (df['dt_criacao_p'].dt.date == df['dt_edicao_p'].dt.date)
            n_nunca = int(mesma_data.sum())
            temp['nunca_editados'] = n_nunca
            temp['nunca_editados_pct'] = fmt_pct(n_nunca, total)
            dois_anos = hoje - pd.DateOffset(years=2)
            antigos = (df['dt_criacao_p'] < dois_anos) & (df['dt_edicao_p'] < dois_anos)
            n_antigos = int(antigos.sum())
            temp['antigos_sem_edicao'] = n_antigos
            temp['antigos_sem_edicao_pct'] = fmt_pct(n_antigos, total)
        dados['temporal'] = temp

    # ── Determinantes sociais ────────────────────────────────────────────────
    det = {}
    if 'escolaridade' in df.columns:
        dist_esc = []
        for v, n in df['escolaridade'].value_counts(dropna=False).items():
            label = 'Sem informação' if (pd.isna(v) or str(v).strip().lower() in ('nan', '')) else str(v).strip().title()
            dist_esc.append({'label': label, 'n': int(n), 'pct': fmt_pct(n, total)})
        det['escolaridade'] = dist_esc
        n_analfa = int(df['escolaridade'].str.upper().str.contains('ANALFA', na=False).sum())
        det['analfabetos'] = n_analfa
        det['analfabetos_pct'] = fmt_pct(n_analfa, total)
    if 'trabalho' in df.columns:
        dist_trab = []
        for v, n in df['trabalho'].value_counts(dropna=False).items():
            label = 'Sem informação' if (pd.isna(v) or str(v).strip().lower() in ('nan', '')) else str(v).strip().title()
            dist_trab.append({'label': label, 'n': int(n), 'pct': fmt_pct(n, total)})
        det['trabalho'] = dist_trab
        n_desemp = int(df['trabalho'].str.upper().str.contains('DESEMPREGA|DESOCUPA', na=False).sum())
        det['desempregados'] = n_desemp
        det['desempregados_pct'] = fmt_pct(n_desemp, total)
    if 'bairro' in df.columns:
        sem_bairro = int(is_blank(df['bairro']).sum())
        dist_bairro = []
        for v, n in df['bairro'].value_counts(dropna=True).head(20).items():
            label = str(v).strip().title()
            if label.lower() in ('nan', ''):
                continue
            dist_bairro.append({'label': label, 'n': int(n), 'pct': fmt_pct(n, total)})
        det['bairro'] = dist_bairro
        det['sem_bairro'] = sem_bairro
        det['sem_bairro_pct'] = fmt_pct(sem_bairro, total)
    if det:
        dados['determinantes_sociais'] = det

    # ── Inconsistências avançadas ────────────────────────────────────────────
    # Duplicidade de CPF/CNS é a exceção que usa df_completo (inclui
    # óbito/mudança de território), pois o par duplicado pode ter uma das
    # entradas já marcada como inativa. As demais checagens usam só ativos.
    incons = {}
    if 'cpf' in df_completo.columns:
        df_com_cpf = df_completo[~is_blank_cpf(df_completo['cpf'])].copy()
        cpf_dup = df_com_cpf[df_com_cpf.duplicated(subset=['cpf'], keep=False)]
        incons['cpf_duplicado'] = int(len(cpf_dup))
        incons['cpf_dup_pct'] = fmt_pct(len(cpf_dup), len(df_completo))
        if len(cpf_dup) > 0:
            dup_det = []
            for cpf_val, g in list(cpf_dup.groupby('cpf'))[:20]:
                entry = {'cpf': str(cpf_val), 'ocorrencias': len(g)}
                if 'nome' in g.columns:
                    entry['nomes'] = [str(nm).title() for nm in g['nome'].tolist()[:3]]
                if 'equipe' in g.columns:
                    entry['equipes'] = [str(e).title() for e in g['equipe'].tolist()[:3]]
                dup_det.append(entry)
            incons['cpf_dup_detalhes'] = dup_det
    if 'cns' in df_completo.columns:
        df_com_cns = df_completo[~is_blank(df_completo['cns'])].copy()
        cns_dup = df_com_cns[df_com_cns.duplicated(subset=['cns'], keep=False)]
        incons['cns_duplicado'] = int(len(cns_dup))
        incons['cns_dup_pct'] = fmt_pct(len(cns_dup), len(df_completo))
    if 'idade_num' in df.columns:
        idade_imp = df[df['idade_num'].notna() & (df['idade_num'] > 120)]
        incons['idade_improvavel'] = int(len(idade_imp))
        incons['idade_improvavel_pct'] = fmt_pct(len(idade_imp), total)
    if incons:
        dados['inconsistencias'] = incons

    # ── Qualidade dos dados ──────────────────────────────────────────────
    checks = []
    if 'cpf' in df.columns:
        n = int(is_blank_cpf(df['cpf']).sum())
        checks.append({'label': 'Sem CPF', 'n': n, 'pct': fmt_pct(n, total)})
    if 'cns' in df.columns:
        n = int(is_blank(df['cns']).sum())
        checks.append({'label': 'Sem CNS', 'n': n, 'pct': fmt_pct(n, total)})
    if 'cpf' in df.columns and 'cns' in df.columns:
        n = int((is_blank_cpf(df['cpf']) & is_blank(df['cns'])).sum())
        checks.append({'label': 'Sem CPF e sem CNS', 'n': n, 'pct': fmt_pct(n, total)})
    if 'nascimento' in df.columns:
        n = int(is_blank(df['nascimento']).sum())
        checks.append({'label': 'Sem data de nascimento', 'n': n, 'pct': fmt_pct(n, total)})
    if 'cel' in df.columns and 'tel_res' in df.columns:
        n = int((is_blank(df['cel']) & is_blank(df['tel_res'])).sum())
        checks.append({'label': 'Sem nenhum telefone', 'n': n, 'pct': fmt_pct(n, total)})
    if 'vinc_familia' in df.columns:
        n = int((df['vinc_familia'].str.strip() == 'NÃO').sum())
        checks.append({'label': 'Não vinculado à família', 'n': n, 'pct': fmt_pct(n, total)})
    esc_col = [c for c in df.columns if 'scolaridade' in c.lower() or 'SCOLAR' in c]
    if esc_col:
        n = int(is_blank(df[esc_col[0]]).sum())
        checks.append({'label': 'Sem escolaridade preenchida', 'n': n, 'pct': fmt_pct(n, total)})
    if 'dt_edicao_p' in df.columns:
        n = int((df['dt_edicao_p'] < pd.Timestamp('2025-04-27')).sum())
        checks.append({'label': 'Cadastros desatualizados (últ. edição < abr/2025)', 'n': n, 'pct': fmt_pct(n, total)})
        n2 = int((df['dt_edicao_p'] >= pd.Timestamp('2026-01-01')).sum())
        checks.append({'label': 'Editados em 2026', 'n': n2, 'pct': fmt_pct(n2, total)})
    dados['qualidade'] = checks

    # ── Origem das fichas ────────────────────────────────────────────────
    if 'origem' in df.columns:
        origens = []
        for v, n in df['origem'].value_counts(dropna=False).items():
            label = str(v).strip()
            if label.lower() == 'nan' or pd.isna(v):
                label = 'Sem informação'
            else:
                label = label.title()
            origens.append({'label': label, 'n': int(n), 'pct': fmt_pct(n, total)})
        dados['origens'] = origens

    # ── Alertas ──────────────────────────────────────────────────────────
    alertas = []
    if 'cpf' in df.columns:
        n = int(is_blank_cpf(df['cpf']).sum())
        if n / total > 0.25:
            alertas.append({'nivel': 'CRÍTICO', 'msg': f"CPF ausente em {fmt_n(n)} registros ({fmt_pct_int(n, total)}). Compromete integração com RNDS, BPC e Bolsa Família."})
        elif n / total > 0.10:
            alertas.append({'nivel': 'ATENÇÃO', 'msg': f"CPF ausente em {fmt_n(n)} registros ({fmt_pct_int(n, total)}). Regularizar com busca ativa nas microáreas."})

    if 'cns' in df.columns:
        n = int(is_blank(df['cns']).sum())
        if n / total > 0.30:
            alertas.append({'nivel': 'CRÍTICO', 'msg': f"CNS ausente em {fmt_n(n)} registros ({fmt_pct_int(n, total)}). Impede vinculação de atendimentos."})

    if 'vinc_familia' in df.columns:
        n = int((df['vinc_familia'].str.strip() == 'NÃO').sum())
        if n / total > 0.20:
            alertas.append({'nivel': 'ATENÇÃO', 'msg': f"{fmt_n(n)} cidadãos não vinculados à família ({fmt_pct_int(n, total)}). Verificar cadastros domiciliares correspondentes."})

    if 'dt_edicao_p' in df.columns:
        n = int((df['dt_edicao_p'] < pd.Timestamp('2025-04-27')).sum())
        if n / total > 0.30:
            alertas.append({'nivel': 'ATENÇÃO', 'msg': f"{fmt_n(n)} cadastros sem edição após abril/2025 ({fmt_pct_int(n, total)}). Planejar atualização cadastral."})

    if 'equipe' in df.columns and 'cpf' in df.columns:
        for eq, g in df.groupby('equipe', dropna=False):
            pct_s_cpf = is_blank_cpf(g['cpf']).sum() / len(g)
            eq_label = 'Sem equipe' if pd.isna(eq) else str(eq).title()
            if pct_s_cpf > 0.40:
                alertas.append({'nivel': 'ATENÇÃO', 'msg': f"Equipe {eq_label}: {pct_s_cpf*100:.0f}% sem CPF. Necessita atenção específica."})

    if 'cond_saude' in df.columns:
        n_sem = int(is_blank(df['cond_saude']).sum())
        if n_sem / total > 0.50:
            alertas.append({'nivel': 'INFO', 'msg': f"{fmt_n(n_sem)} registros ({fmt_pct_int(n_sem, total)}) sem condição de saúde preenchida. Campo importante para gestão de crônicos e atenção domiciliar."})

        # Acamados e domiciliados — sempre geram alerta informativo para visita domiciliar
        n_acam = int(df['cond_saude'].str.contains('ACAMADO', na=False).sum())
        n_dom  = int(df['cond_saude'].str.contains(r'DOMICILI[AO]', na=False).sum())
        if n_acam > 0:
            alertas.append({'nivel': 'INFO', 'msg': f"{fmt_n(n_acam)} paciente(s) acamado(s) identificado(s). Exige(m) programa de Atenção Domiciliar (AD) ativo — verificar cadência de visitas e plano de cuidados."})
        if n_dom > 0:
            alertas.append({'nivel': 'INFO', 'msg': f"{fmt_n(n_dom)} paciente(s) domiciliado(s) (restrição de mobilidade). Garantir acesso a consultas domiciliares e insumos pela equipe de AD."})

        # Obesidade — alerta quando prevalência elevada
        n_obeso = int(df['cond_saude'].str.contains(r'OBES[OA]|OBESIDADE|SOBREPESO', na=False).sum())
        if n_obeso / total > 0.10:
            alertas.append({'nivel': 'ATENÇÃO', 'msg': f"{fmt_n(n_obeso)} registros com obesidade/sobrepeso ({fmt_pct_int(n_obeso, total)}). Planejar ações de promoção alimentar, atividade física e rastreamento metabólico."})

    dados['alertas'] = alertas

    # ── Recomendações (dinâmicas) ────────────────────────────────────────
    recs = []
    rec_n = 0

    # CPF/CNS
    if 'cpf' in df.columns or 'cns' in df.columns:
        n_cpf = int(is_blank_cpf(df['cpf']).sum()) if 'cpf' in df.columns else 0
        n_cns = int(is_blank(df['cns']).sum()) if 'cns' in df.columns else 0
        if n_cpf > 0 or n_cns > 0:
            rec_n += 1
            recs.append({
                'num': str(rec_n),
                'titulo': 'Regularização de CPF/CNS',
                'descricao': f"Gerar lista nominal por equipe e microárea (ver planilhas Excel anexas). Meta: ≥90% de CPF e CNS preenchidos até o próximo mês.",
            })

    # Vínculo familiar
    if 'vinc_familia' in df.columns:
        n_vf = int((df['vinc_familia'].str.strip() == 'NÃO').sum())
        if n_vf > 0:
            rec_n += 1
            recs.append({
                'num': str(rec_n),
                'titulo': 'Vínculo Familiar',
                'descricao': 'Regularizar cidadãos não vinculados à família — impede completude do núcleo domiciliar.',
            })

    # Atualização cadastral
    if 'dt_edicao_p' in df.columns:
        n_des = int((df['dt_edicao_p'] < pd.Timestamp('2025-04-27')).sum())
        if n_des > 0:
            rec_n += 1
            recs.append({
                'num': str(rec_n),
                'titulo': 'Atualização Cadastral',
                'descricao': 'Priorizar visitas a cadastros sem edição desde antes de abril/2025.',
            })

    # Escolaridade
    if esc_col:
        n_esc = int(is_blank(df[esc_col[0]]).sum())
        if n_esc > 0:
            pct_esc = fmt_pct(n_esc, total)
            rec_n += 1
            recs.append({
                'num': str(rec_n),
                'titulo': 'Escolaridade',
                'descricao': f"{pct_esc} sem escolaridade preenchida — campo importante para indicadores educacionais.",
            })

    # Gestantes
    if 'cond_saude' in df.columns:
        gestantes_n = int(df['cond_saude'].str.contains('GESTANTE', na=False).sum())
        if gestantes_n > 0:
            rec_n += 1
            recs.append({
                'num': str(rec_n),
                'titulo': 'Gestantes',
                'descricao': 'Confirmar vinculação ao pré-natal e acompanhamento nas equipes com maior concentração.',
            })

    dados['recomendacoes'] = recs

    return dados


# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO TXT (mantém formato original)
# ══════════════════════════════════════════════════════════════════════════════
def sec(title: str, width: int = 72) -> str:
    return f"\n{'═' * width}\n  {title}\n{'═' * width}"


def subsec(title: str) -> str:
    return f"\n  ── {title}"


def gerar_txt(dados: dict) -> str:
    """Gera o relatório TXT a partir dos dados estruturados."""
    total = dados['total']
    linhas = []
    add = linhas.append

    add("=" * 72)
    add(f"  RELATÓRIO MENSAL DE CADASTRO INDIVIDUAL")
    add(f"  Município: {dados['municipio'].upper()}")
    add(f"  Competência: {dados['competencia']}")
    add(f"  Gerado em: {dados['data_geracao']} às {dados['hora_geracao']}")
    add("=" * 72)

    # 1. Visão Geral
    add(sec("1. VISÃO GERAL"))
    add(f"\n  Total de registros no arquivo : {total:>8,}")

    if 'visao_geral' in dados:
        vg = dados['visao_geral']
        add(f"  Registros ATIVOS (sem saída)  : {vg['ativos']:>8,} ({vg['ativos_pct']})")
        add(f"  Com saída de cadastro         : {vg['saida']:>8,} ({vg['saida_pct']})")
        add(f"    └─ Óbitos registrados       : {vg['obitos']:>8,}")
        add(f"    └─ Mudança de território    : {vg['mudanca']:>8,}")
        add(f"    └─ Outras saídas            : {vg['outras_saidas']:>8,}")
        if 'fora_area' in vg:
            add(f"  Fora de área (micro-área)     : {vg['fora_area']:>8,}")

    if 'sexo' in dados:
        add(subsec("Distribuição por sexo"))
        for item in dados['sexo']:
            add(f"    {item['label']:<25}: {item['n']:>7,}  ({item['pct']})")

    if 'raca' in dados:
        add(subsec("Raça/Cor"))
        for item in dados['raca']:
            add(f"    {item['label']:<30}: {item['n']:>7,}  ({item['pct']})")

    # 2. Equipes
    if 'equipes' in dados:
        add(sec("2. DISTRIBUIÇÃO E COMPLETUDE POR EQUIPE"))
        add(f"\n  {'EQUIPE':<14} {'TOTAL':>7}  {'S/CPF':>7} {'%':>5}  {'S/CNS':>7} {'%':>5}  {'S/END':>7} {'%':>5}  {'S/VF':>7} {'%':>5}  {'DESATUAL':>8} {'%':>5}")
        add(f"  {'-'*13} {'-'*7}  {'-'*7} {'-'*5}  {'-'*7} {'-'*5}  {'-'*7} {'-'*5}  {'-'*7} {'-'*5}  {'-'*8} {'-'*5}")
        for eq in dados['equipes']:
            add(f"  {eq['nome']:<14} {eq['total']:>7,}  {eq['s_cpf']:>7,} {eq['s_cpf_pct']:>5}  {eq['s_cns']:>7,} {eq['s_cns_pct']:>5}  {eq['s_end']:>7,} {eq['s_end_pct']:>5}  {eq['s_vf']:>7,} {eq['s_vf_pct']:>5}  {eq['desatual']:>8,} {eq['desatual_pct']:>5}")
        add("\n  Legenda: S/CPF=sem CPF | S/CNS=sem CNS | S/END=sem endereço | S/VF=não vinculado à família | DESATUAL=sem edição após abr/2025")

    # 3. Grupos
    add(sec("3. GRUPOS POPULACIONAIS PRIORITÁRIOS"))
    gp = dados.get('grupos', {})
    if 'idosos' in gp:
        add(f"\n  Idosos (≥60 anos)    : {gp['idosos']['n']:>7,}  ({gp['idosos']['pct']})")
    if 'criancas' in gp:
        add(f"  Crianças (0-4 anos)  : {gp['criancas']['n']:>7,}  ({gp['criancas']['pct']})")
    if 'gestantes' in gp:
        add(f"  Gestantes            : {gp['gestantes']['n']:>7,}")

    if 'auxilio' in gp:
        aux = gp['auxilio']
        add(f"\n  Recebem algum auxílio: {aux['n_com']:>7,}  ({aux['pct_com']})")
        add(f"  Não informado        : {aux['n_nao_inf']:>7,}  ({aux['pct_nao_inf']})")
        if aux.get('beneficios'):
            add(subsec("Benefícios recebidos"))
            for b in aux['beneficios']:
                add(f"    {b['label']:<50}: {b['n']:>6,}  ({b['pct']})")

    if 'condicoes' in gp:
        add(subsec("Condições de saúde registradas"))
        for c in gp['condicoes']:
            add(f"    {c['label']:<30}: {c['n']:>6,}  ({c['pct']})")

    if 'gestantes_equipe' in gp:
        add(subsec("Gestantes por equipe"))
        for g in gp['gestantes_equipe']:
            add(f"    {g['equipe']:<20}: {g['n']:>4,}")

    if 'atencao_domiciliar' in gp:
        ad = gp['atencao_domiciliar']
        add(subsec("Atenção Domiciliar (Acamados e Domiciliados)"))
        add(f"    Acamados     : {ad['acamados']:>4,}  ({ad['acamados_pct']})")
        add(f"    Domiciliados : {ad['domiciliados']:>4,}  ({ad['domiciliados_pct']})")
        if ad.get('por_equipe'):
            add(f"    Por equipe:")
            for eq in ad['por_equipe']:
                add(f"      {eq['equipe']:<20}: {eq['n']:>3,}")
        if ad.get('por_microarea'):
            add(f"    Por microárea:")
            for ma in ad['por_microarea']:
                add(f"      Microárea {ma['microarea']:<10}: {ma['n']:>3,}")

    # 4. Qualidade
    add(sec("4. QUALIDADE DOS DADOS"))
    add("")
    for q in dados.get('qualidade', []):
        pct_num = q['n'] / total * 100
        bar = '█' * int(pct_num / 5) + '░' * (20 - int(pct_num / 5))
        add(f"  {q['label']:<45}: {q['n']:>7,}  {q['pct']:>7}  [{bar}]")

    # 5. Origem
    if 'origens' in dados:
        add(sec("5. ORIGEM DAS FICHAS"))
        add("")
        for o in dados['origens']:
            add(f"  {o['label']:<30}: {o['n']:>7,}  ({o['pct']})")

    # 6. Alertas
    add(sec("6. ALERTAS PARA A COORDENAÇÃO"))
    add("")
    for a in dados.get('alertas', []):
        add(f"  [{a['nivel']}] {a['msg']}")
    if not dados.get('alertas'):
        add("  Nenhum alerta crítico identificado.")

    # 7. Recomendações
    add(sec("7. RECOMENDAÇÕES PRIORITÁRIAS"))
    add("")
    for r in dados.get('recomendacoes', []):
        add(f"  {r['num']}. {r['titulo'].upper()}")
        add(f"     {r['descricao']}")
        add("")

    # 8. Saúde da Mulher
    sm = dados.get('saude_mulher', {})
    if sm:
        add(sec("8. SAÚDE DA MULHER"))
        add(f"\n  Registros femininos          : {sm['total']:>8,}  ({sm['pct']})")
        fr = sm.get('faixa_reprodutiva', {})
        if fr:
            add(f"  Faixa reprodutiva (10–49 a.) : {fr['n']:>8,}  ({fr['pct']} das mulheres)")
        if sm.get('idosas'):
            add(f"  Idosas (≥60 anos)            : {sm['idosas']['n']:>8,}  ({sm['idosas']['pct']} das mulheres)")
        if sm.get('gestantes'):
            g_sm = sm['gestantes']
            add(f"  Gestantes identificadas      : {g_sm['n']:>8,}  (taxa obs.: {g_sm['taxa_observada_pct']})")
            if g_sm.get('alerta'):
                add(f"  ⚠  Taxa de gestantes abaixo do esperado — revisar cobertura do pré-natal")

    # 9. Ciclos de Vida
    if 'ciclos_vida' in dados:
        add(sec("9. CICLOS DE VIDA (MS / PNAB)"))
        cv = dados['ciclos_vida']
        add(f"\n  {'CICLO DE VIDA':<32} {'N':>8} {'% TOTAL':>9} {'% C/IDADE':>10}")
        add(f"  {'-'*31} {'-'*8}  {'-'*9} {'-'*10}")
        for c in cv['ciclos']:
            add(f"  {c['label']:<32} {c['n']:>8,}  {c['pct']:>9} {c['pct_validos']:>10}")
        add(f"\n  Sem idade informada          : {cv['sem_idade']:>8,}  ({cv['sem_idade_pct']})")

    # 10. Estratificação de Risco
    if 'estratificacao_risco' in dados:
        add(sec("10. ESTRATIFICAÇÃO DE RISCO (APS)"))
        add("")
        for r in dados['estratificacao_risco']:
            try:
                bar_v = int(float(r['pct'].replace('%', '').replace(',', '.')) / 5)
            except Exception:
                bar_v = 0
            bar = '█' * bar_v + '░' * (20 - bar_v)
            add(f"  {r['nivel']:<22}: {r['n']:>7,}  {r['pct']:>7}  [{bar}]")
        add(f"\n  Critérios:")
        add(f"    ALTO  — HAS+DM, doença cardiovascular, renal crônica, neoplasia ou ≥2 crônicas")
        add(f"    MÉDIO — 1 condição crônica, gestante, álcool/tabaco, longevo (≥80 anos)")
        add(f"    BAIXO — sem condições crônicas registradas")

    # 11. Por Microárea
    if 'microareas' in dados:
        add(sec("11. ANÁLISE POR MICROÁREA"))
        add(f"\n  {'MICROÁREA':<18} {'EQUIPE':<14} {'ACS':<20} {'TOTAL':>6}  {'S/CPF':>6} {'%':>5}  {'S/CNS':>6} {'%':>5}  {'S/END':>6} {'%':>5}  {'DESAT.':>6} {'%':>5}")
        add(f"  {'-'*17} {'-'*13} {'-'*19} {'-'*6}  {'-'*6} {'-'*5}  {'-'*6} {'-'*5}  {'-'*6} {'-'*5}  {'-'*6} {'-'*5}")
        for ma in dados['microareas']:
            add(f"  {ma['microarea']:<18} {ma['equipe']:<14} {ma['acs']:<20} {ma['total']:>6,}  "
                f"{ma['s_cpf']:>6,} {ma['s_cpf_pct']:>5}  {ma['s_cns']:>6,} {ma['s_cns_pct']:>5}  "
                f"{ma['s_end']:>6,} {ma['s_end_pct']:>5}  {ma['desatual']:>6,} {ma['desatual_pct']:>5}")

    # 12. Por ACS
    if 'acs' in dados:
        add(sec("12. ANÁLISE POR AGENTE COMUNITÁRIO DE SAÚDE (ACS)"))
        add(f"\n  {'ACS':<25} {'EQUIPE':<14} {'TOTAL':>6}  {'S/CPF':>6} {'%':>5}  {'S/CNS':>6} {'%':>5}  {'DESAT.':>6} {'%':>5}")
        add(f"  {'-'*24} {'-'*13} {'-'*6}  {'-'*6} {'-'*5}  {'-'*6} {'-'*5}  {'-'*6} {'-'*5}")
        for a in dados['acs']:
            add(f"  {a['nome']:<25} {a['equipe']:<14} {a['total']:>6,}  "
                f"{a['s_cpf']:>6,} {a['s_cpf_pct']:>5}  {a['s_cns']:>6,} {a['s_cns_pct']:>5}  "
                f"{a['desatual']:>6,} {a['desatual_pct']:>5}")

    # 13. Análise Temporal
    if 'temporal' in dados:
        add(sec("13. ANÁLISE TEMPORAL DE CADASTROS"))
        temp = dados['temporal']
        if temp.get('por_ano'):
            add(subsec("Cadastros por ano de criação"))
            for item in temp['por_ano']:
                bar_v = min(int(item['n'] / total * 100 / 5), 20)
                bar = '█' * bar_v + '░' * (20 - bar_v)
                add(f"    {item['ano']}  {item['n']:>7,}  [{bar}]")
        if 'nunca_editados' in temp:
            add(f"\n  Nunca editados (criação = edição) : {temp['nunca_editados']:>7,}  ({temp['nunca_editados_pct']})")
        if 'antigos_sem_edicao' in temp:
            add(f"  Criados há +2 anos sem edição     : {temp['antigos_sem_edicao']:>7,}  ({temp['antigos_sem_edicao_pct']})")

    # 14. Determinantes Sociais
    det = dados.get('determinantes_sociais', {})
    if det:
        add(sec("14. DETERMINANTES SOCIAIS DA SAÚDE"))
        if det.get('escolaridade'):
            add(subsec("Escolaridade"))
            for e in det['escolaridade']:
                add(f"    {e['label']:<45}: {e['n']:>7,}  ({e['pct']})")
            if det.get('analfabetos'):
                add(f"\n  Analfabetos: {det['analfabetos']:,}  ({det['analfabetos_pct']})")
        if det.get('trabalho'):
            add(subsec("Situação no mercado de trabalho"))
            for t in det['trabalho']:
                add(f"    {t['label']:<45}: {t['n']:>7,}  ({t['pct']})")
        if det.get('bairro'):
            add(subsec("Top 20 bairros por nº de cadastros"))
            for b in det['bairro']:
                add(f"    {b['label']:<35}: {b['n']:>7,}  ({b['pct']})")

    # 15. Inconsistências Avançadas
    incons = dados.get('inconsistencias', {})
    if incons:
        add(sec("15. INCONSISTÊNCIAS AVANÇADAS"))
        add("")
        if 'cpf_duplicado' in incons:
            add(f"  CPF duplicado na base         : {incons['cpf_duplicado']:>7,}  ({incons['cpf_dup_pct']})")
        if 'cns_duplicado' in incons:
            add(f"  CNS duplicado na base         : {incons['cns_duplicado']:>7,}  ({incons['cns_dup_pct']})")
        if 'idade_improvavel' in incons:
            add(f"  Idades improváveis (>120 anos): {incons['idade_improvavel']:>7,}  ({incons['idade_improvavel_pct']})")
        if incons.get('cpf_dup_detalhes'):
            add(subsec("CPFs duplicados — primeiros casos"))
            for d in incons['cpf_dup_detalhes'][:10]:
                nomes = ' / '.join(d.get('nomes', []))
                add(f"    CPF {d['cpf']}: {d['ocorrencias']}x  — {nomes}")

    add("=" * 72)
    add(f"  Relatório gerado automaticamente em {dados['data_geracao']} {dados['hora_geracao']}")
    add(f"  Fonte: {dados['fonte']}")
    add("=" * 72)

    return '\n'.join(linhas)


# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO EXCEL ANALÍTICO (openpyxl)
# ══════════════════════════════════════════════════════════════════════════════
def gerar_excel_relatorio(dados: dict, out_path: Path) -> bool:
    """Gera planilha Excel com múltiplas abas analíticas."""
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  ⚠  openpyxl não encontrado. Execute: pip install openpyxl")
        return False

    # ── Estilos ──────────────────────────────────────────────────────────────
    C_AZE = '1F497D'; C_AZM = '2E74B5'; C_AZC = 'BDD7EE'
    C_CIN = 'F2F2F2'; C_VDE = 'E2EFDA'; C_AMR = 'FFF2CC'
    C_VML = 'FFE7E7'; C_BCO = 'FFFFFF'

    f_h1   = PatternFill('solid', fgColor=C_AZE)
    f_h2   = PatternFill('solid', fgColor=C_AZM)
    f_alt  = PatternFill('solid', fgColor=C_CIN)
    f_tot  = PatternFill('solid', fgColor=C_AZC)
    f_alto = PatternFill('solid', fgColor=C_VML)
    f_med  = PatternFill('solid', fgColor=C_AMR)
    f_baixo= PatternFill('solid', fgColor=C_VDE)

    ft_tit  = Font(name='Calibri', bold=True, size=13, color='FFFFFF')
    ft_sub  = Font(name='Calibri', italic=True, size=9, color='FFFFFF')
    ft_h    = Font(name='Calibri', bold=True, size=9, color='FFFFFF')
    ft_corp = Font(name='Calibri', size=9)
    ft_tot  = Font(name='Calibri', bold=True, size=10, color=C_AZE)
    ft_lab  = Font(name='Calibri', bold=True, size=9, color=C_AZE)

    brd = Border(
        left=Side('thin', color='D0D0D0'), right=Side('thin', color='D0D0D0'),
        top=Side('thin', color='D0D0D0'), bottom=Side('thin', color='D0D0D0'),
    )
    AC = Alignment(horizontal='center', vertical='center', wrap_text=True)
    AL = Alignment(horizontal='left',   vertical='center')
    AR = Alignment(horizontal='right',  vertical='center')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    total = dados['total']
    mun   = dados['municipio'].upper()

    def _title(ws, txt, ncols, row=1, fill=None, font=None):
        ws.row_dimensions[row].height = 26
        c = ws.cell(row=row, column=1, value=txt)
        c.font = font or ft_tit; c.fill = fill or f_h1
        c.alignment = AC; c.border = brd
        if ncols > 1:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)

    def _header(ws, hdrs, row, fill=None):
        ws.row_dimensions[row].height = 18
        for ci, h in enumerate(hdrs, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.font = ft_h; c.fill = fill or f_h2; c.alignment = AC; c.border = brd

    def _row(ws, vals, row, fill=None, font=None, aligns=None):
        ws.row_dimensions[row].height = 15
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=ci, value=v)
            c.font = font or ft_corp
            c.alignment = (aligns[ci - 1] if aligns else AL)
            c.border = brd
            if fill:
                c.fill = fill

    def _widths(ws, ws_list):
        for i, w in enumerate(ws_list, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # ── Aba 1: Resumo Geral ──────────────────────────────────────────────────
    ws = wb.create_sheet('Resumo Geral')
    ws.sheet_view.showGridLines = False
    _title(ws, f'CADASTRO INDIVIDUAL · {mun} · {dados["competencia"]}', 4)
    _title(ws, f'Prima Qualitá Saúde · {dados["data_geracao"]} {dados["hora_geracao"]} · Fonte: {dados["fonte"]}',
           4, row=2, fill=f_h2, font=ft_sub)
    r = 4
    _header(ws, ['INDICADOR', 'VALOR', 'PERCENTUAL', 'STATUS'], r); r += 1

    def _res(label, val, pct='', status='', f_st=None):
        nonlocal r
        bg = f_alt if r % 2 == 0 else None
        _row(ws, [label, val, pct, status], r, fill=bg, aligns=[AL, AR, AC, AC])
        if f_st:
            ws.cell(row=r, column=4).fill = f_st
        r += 1

    _res('Total de Registros', total, '100%')
    vg = dados.get('visao_geral', {})
    if vg:
        _res('  Registros Ativos (sem saída)', vg['ativos'], vg['ativos_pct'])
        _res('  Com Saída de Cadastro', vg['saida'], vg['saida_pct'])
        _res('    ↳ Óbitos', vg['obitos'])
        _res('    ↳ Mudança de Território', vg['mudanca'])
        _res('    ↳ Outras Saídas (recusa, erro de cadastro etc.)', vg['outras_saidas'])
        if 'fora_area' in vg:
            _res('  Fora de Área (micro-área)', vg['fora_area'])

    if 'ciclos_vida' in dados:
        for c_item in dados['ciclos_vida']['ciclos']:
            _res(f"  {c_item['label']}", c_item['n'], c_item['pct'])

    for q in dados.get('qualidade', []):
        pn = q['n'] / total * 100 if total else 0
        if pn > 25:
            st, fs = '⚠  CRÍTICO', f_alto
        elif pn > 10:
            st, fs = '⚠  ATENÇÃO', f_med
        else:
            st, fs = '✓ OK', f_baixo
        _res(q['label'], q['n'], q['pct'], st, fs)

    _widths(ws, [48, 12, 12, 12]); ws.freeze_panes = 'A5'

    # ── Aba 2: Ciclos de Vida ────────────────────────────────────────────────
    if 'ciclos_vida' in dados:
        ws2 = wb.create_sheet('Ciclos de Vida')
        ws2.sheet_view.showGridLines = False
        cv = dados['ciclos_vida']
        _title(ws2, f'CICLOS DE VIDA · {mun}', 5)
        _title(ws2, 'Distribuição etária segundo faixas do Ministério da Saúde (PNAB)',
               5, row=2, fill=f_h2, font=ft_sub)
        r2 = 4
        _header(ws2, ['CICLO DE VIDA', 'N', '% DO TOTAL', 'N MASC.', 'N FEM.'], r2); r2 += 1
        fills_cv = [
            PatternFill('solid', fgColor='E3F2FD'), PatternFill('solid', fgColor='FFF3E0'),
            PatternFill('solid', fgColor='E8F5E9'), PatternFill('solid', fgColor='F3E5F5'),
            PatternFill('solid', fgColor='FBE9E7'), PatternFill('solid', fgColor='FCE4EC'),
        ]
        for fi, ci in enumerate(cv['ciclos']):
            sx = ci.get('por_sexo', {})
            masc = sx.get('Masculino', sx.get('M', ''))
            fem  = sx.get('Feminino',  sx.get('F', ''))
            _row(ws2, [ci['label'], ci['n'], ci['pct'], masc, fem], r2,
                 fill=fills_cv[fi % len(fills_cv)], aligns=[AL, AR, AC, AR, AR])
            r2 += 1
        tot_cv = sum(ci['n'] for ci in cv['ciclos'])
        _row(ws2, ['TOTAL COM IDADE INFORMADA', tot_cv, '', '', ''], r2,
             fill=f_tot, font=ft_tot, aligns=[AL, AR, AC, AR, AR]); r2 += 1
        _row(ws2, ['Sem idade informada', cv['sem_idade'], cv['sem_idade_pct'], '', ''], r2,
             aligns=[AL, AR, AC, AR, AR])
        _widths(ws2, [32, 10, 12, 10, 10]); ws2.freeze_panes = 'A5'

    # ── Aba 3: Estratificação de Risco ───────────────────────────────────────
    if 'estratificacao_risco' in dados:
        ws3 = wb.create_sheet('Estratificação de Risco')
        ws3.sheet_view.showGridLines = False
        er = dados['estratificacao_risco']
        _title(ws3, f'ESTRATIFICAÇÃO DE RISCO · {mun}', 3)
        _title(ws3, 'Classificação baseada nas condições de saúde registradas',
               3, row=2, fill=f_h2, font=ft_sub)
        r3 = 4
        _header(ws3, ['NÍVEL DE RISCO', 'N', '% TOTAL'], r3); r3 += 1
        fills_er = {'Alto Risco': f_alto, 'Médio Risco': f_med, 'Baixo Risco': f_baixo}
        for ri in er:
            _row(ws3, [ri['nivel'], ri['n'], ri['pct']], r3,
                 fill=fills_er.get(ri['nivel']), aligns=[AL, AR, AC]); r3 += 1
        _row(ws3, ['TOTAL', total, '100%'], r3, fill=f_tot, font=ft_tot, aligns=[AL, AR, AC]); r3 += 2

        # Critérios
        c = ws3.cell(row=r3, column=1, value='CRITÉRIOS DE CLASSIFICAÇÃO')
        c.font = ft_lab; ws3.merge_cells(f'A{r3}:C{r3}'); r3 += 1
        for nivel, crit in [
            ('ALTO RISCO',  'HAS + DM concomitantes, doença cardiovascular, renal crônica, neoplasia, ou ≥2 condições crônicas'),
            ('MÉDIO RISCO', 'Condição crônica isolada (HAS ou DM), gestante, usuário de álcool/tabaco, longevo ≥80 anos'),
            ('BAIXO RISCO', 'Sem condições crônicas registradas'),
        ]:
            _row(ws3, [nivel, crit, ''], r3, aligns=[AL, AL, AL])
            ws3.merge_cells(f'B{r3}:C{r3}'); r3 += 1

        # Por equipe
        equipes = set()
        for ri in er:
            equipes.update(ri.get('por_equipe', {}).keys())
        equipes = sorted(equipes)
        if equipes:
            r3 += 1
            c = ws3.cell(row=r3, column=1, value='ESTRATIFICAÇÃO POR EQUIPE')
            c.font = ft_lab; r3 += 1
            _header(ws3, ['EQUIPE', 'ALTO RISCO', 'MÉDIO RISCO', 'BAIXO RISCO'], r3); r3 += 1
            for fi, eq in enumerate(equipes):
                vals = [eq] + [ri.get('por_equipe', {}).get(eq, 0) for ri in er]
                _row(ws3, vals, r3, fill=f_alt if fi % 2 == 0 else None,
                     aligns=[AL, AR, AR, AR]); r3 += 1
        _widths(ws3, [35, 12, 55]); ws3.freeze_panes = 'A5'

    # ── Aba 4: Completude por Equipe ─────────────────────────────────────────
    if 'equipes' in dados:
        ws4 = wb.create_sheet('Completude por Equipe')
        ws4.sheet_view.showGridLines = False
        _title(ws4, f'COMPLETUDE CADASTRAL POR EQUIPE · {mun}', 14)
        _title(ws4, 'Indicadores de qualidade do cadastro individual',
               14, row=2, fill=f_h2, font=ft_sub)
        r4 = 4
        _header(ws4, ['EQUIPE', 'TOTAL', 'S/CPF', '%', 'S/CNS', '%', 'S/END', '%', 'S/VF', '%',
                      'DESATUAL.', '%', 'DUPLIC.', '%'], r4); r4 += 1
        for fi, eq in enumerate(dados['equipes']):
            is_tot = 'TOTAL' in eq['nome'].upper()
            fill = f_tot if is_tot else (f_alt if fi % 2 == 0 else None)
            font = ft_tot if is_tot else ft_corp
            _row(ws4, [eq['nome'], eq['total'],
                       eq['s_cpf'], eq['s_cpf_pct'], eq['s_cns'], eq['s_cns_pct'],
                       eq['s_end'], eq['s_end_pct'], eq['s_vf'],  eq['s_vf_pct'],
                       eq['desatual'], eq['desatual_pct'],
                       eq.get('dup', 0), eq.get('dup_pct', '0%')],
                 r4, fill=fill, font=font,
                 aligns=[AL, AR, AR, AC, AR, AC, AR, AC, AR, AC, AR, AC, AR, AC]); r4 += 1
        r4 += 1
        c = ws4.cell(row=r4, column=1,
                     value='Legenda: S/CPF = sem CPF | S/CNS = sem CNS | S/END = sem endereço | '
                           'S/VF = não vinculado à família | DESATUAL. = sem edição após abr/2025 | '
                           'DUPLIC. = registros em possíveis duplicatas (inclui óbitos/mudança de território)')
        c.font = Font(name='Calibri', italic=True, size=8, color='606060')
        ws4.merge_cells(f'A{r4}:N{r4}')
        _widths(ws4, [20, 8, 8, 6, 8, 6, 8, 6, 8, 6, 10, 6, 8, 6]); ws4.freeze_panes = 'A5'

    # ── Aba 5: Condições de Saúde ────────────────────────────────────────────
    gp = dados.get('grupos', {})
    if gp.get('condicoes') or gp.get('gestantes') or gp.get('auxilio'):
        ws5 = wb.create_sheet('Condições de Saúde')
        ws5.sheet_view.showGridLines = False
        _title(ws5, f'CONDIÇÕES DE SAÚDE · {mun}', 3)
        r5 = 3
        if gp.get('condicoes'):
            r5 += 1
            ws5.cell(row=r5, column=1, value='CONDIÇÕES CRÔNICAS REGISTRADAS').font = ft_lab; r5 += 1
            _header(ws5, ['CONDIÇÃO DE SAÚDE', 'N', '% DO TOTAL'], r5); r5 += 1
            for fi, cd in enumerate(gp['condicoes']):
                _row(ws5, [cd['label'], cd['n'], cd['pct']], r5,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r5 += 1
        if gp.get('gestantes'):
            r5 += 1
            ws5.cell(row=r5, column=1, value='GESTANTES').font = ft_lab; r5 += 1
            _row(ws5, ['Total de Gestantes', gp['gestantes']['n'], ''], r5, aligns=[AL, AR, AC]); r5 += 1
            if gp.get('gestantes_equipe'):
                _header(ws5, ['EQUIPE', 'Nº GESTANTES', ''], r5); r5 += 1
                for fi, g in enumerate(gp['gestantes_equipe']):
                    _row(ws5, [g['equipe'], g['n'], ''], r5,
                         fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r5 += 1
        if gp.get('auxilio'):
            aux = gp['auxilio']
            r5 += 1
            ws5.cell(row=r5, column=1, value='AUXÍLIOS E BENEFÍCIOS SOCIAIS').font = ft_lab; r5 += 1
            _row(ws5, ['Recebem algum auxílio', aux['n_com'], aux['pct_com']], r5, aligns=[AL, AR, AC]); r5 += 1
            _row(ws5, ['Não informado', aux['n_nao_inf'], aux['pct_nao_inf']], r5, aligns=[AL, AR, AC]); r5 += 1
            if aux.get('beneficios'):
                _header(ws5, ['BENEFÍCIO', 'N', '% DO TOTAL'], r5); r5 += 1
                for fi, b in enumerate(aux['beneficios']):
                    _row(ws5, [b['label'], b['n'], b['pct']], r5,
                         fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r5 += 1
        _widths(ws5, [42, 12, 12]); ws5.freeze_panes = 'A3'

    # ── Aba 6: Qualidade dos Dados ───────────────────────────────────────────
    if dados.get('qualidade'):
        ws6 = wb.create_sheet('Qualidade dos Dados')
        ws6.sheet_view.showGridLines = False
        _title(ws6, f'QUALIDADE DOS DADOS CADASTRAIS · {mun}', 4)
        _title(ws6, 'Indicadores de completude cadastral — Atenção Primária à Saúde',
               4, row=2, fill=f_h2, font=ft_sub)
        r6 = 4
        _header(ws6, ['INDICADOR', 'QTD.', '%', 'STATUS'], r6); r6 += 1
        for fi, q in enumerate(dados['qualidade']):
            pn = q['n'] / total * 100 if total else 0
            if pn > 25:
                st, fs = '⚠  CRÍTICO', f_alto
            elif pn > 10:
                st, fs = '⚠  ATENÇÃO', f_med
            else:
                st, fs = '✓ OK', f_baixo
            bg = f_alt if fi % 2 == 0 else None
            _row(ws6, [q['label'], q['n'], q['pct'], st], r6, fill=bg,
                 aligns=[AL, AR, AC, AC])
            ws6.cell(row=r6, column=4).fill = fs
            r6 += 1
        _widths(ws6, [48, 12, 10, 12]); ws6.freeze_panes = 'A5'

    # ── Aba 7: Alertas ───────────────────────────────────────────────────────
    if dados.get('alertas'):
        ws7 = wb.create_sheet('Alertas e Recomendações')
        ws7.sheet_view.showGridLines = False
        _title(ws7, f'ALERTAS E RECOMENDAÇÕES · {mun}', 3)
        r7 = 3
        _header(ws7, ['NÍVEL', 'ALERTA', ''], r7); r7 += 1
        fills_al = {'CRÍTICO': f_alto, 'ATENÇÃO': f_med, 'INFO': PatternFill('solid', fgColor='E3F2FD')}
        for al in dados['alertas']:
            fg = fills_al.get(al['nivel'], None)
            _row(ws7, [al['nivel'], al['msg'], ''], r7, aligns=[AC, AL, AL])
            if fg:
                for ci in range(1, 4):
                    ws7.cell(row=r7, column=ci).fill = fg
            ws7.merge_cells(f'B{r7}:C{r7}'); r7 += 1
        if dados.get('recomendacoes'):
            r7 += 1
            ws7.cell(row=r7, column=1, value='RECOMENDAÇÕES PRIORITÁRIAS').font = ft_lab; r7 += 1
            _header(ws7, ['Nº', 'TÍTULO', 'DESCRIÇÃO'], r7); r7 += 1
            for rec in dados['recomendacoes']:
                _row(ws7, [rec['num'], rec['titulo'], rec['descricao']], r7,
                     aligns=[AC, AL, AL]); r7 += 1
        _widths(ws7, [10, 30, 60]); ws7.freeze_panes = 'A4'

    # ── Aba 7b: Atenção Domiciliar ───────────────────────────────────────────
    ad = dados.get('grupos', {}).get('atencao_domiciliar', {})
    if ad:
        ws_ad = wb.create_sheet('Atenção Domiciliar')
        ws_ad.sheet_view.showGridLines = False
        _title(ws_ad, f'ATENÇÃO DOMICILIAR (ACAMADOS E DOMICILIADOS) · {mun}', 3)
        _title(ws_ad, 'Pacientes com necessidade de cuidado domiciliar — exigem visita periódica da equipe',
               3, row=2, fill=f_h2, font=ft_sub)
        r_ad = 4
        _header(ws_ad, ['CATEGORIA', 'N', '%'], r_ad); r_ad += 1
        _row(ws_ad, ['Acamados', ad['acamados'], ad['acamados_pct']], r_ad,
             fill=f_alto, aligns=[AL, AR, AC]); r_ad += 1
        _row(ws_ad, ['Domiciliados (restrição de mobilidade)', ad['domiciliados'], ad['domiciliados_pct']], r_ad,
             fill=f_med, aligns=[AL, AR, AC]); r_ad += 1
        _row(ws_ad, ['Total em atenção domiciliar', ad['acamados'] + ad['domiciliados'],
                     fmt_pct(ad['acamados'] + ad['domiciliados'], total)],
             r_ad, fill=f_tot, font=ft_tot, aligns=[AL, AR, AC]); r_ad += 2
        if ad.get('por_equipe'):
            ws_ad.cell(row=r_ad, column=1, value='DISTRIBUIÇÃO POR EQUIPE').font = ft_lab; r_ad += 1
            _header(ws_ad, ['EQUIPE', 'N (ACAM + DOM)', ''], r_ad); r_ad += 1
            for fi, eq in enumerate(ad['por_equipe']):
                _row(ws_ad, [eq['equipe'], eq['n'], ''], r_ad,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r_ad += 1
        if ad.get('por_microarea'):
            r_ad += 1
            ws_ad.cell(row=r_ad, column=1, value='DISTRIBUIÇÃO POR MICROÁREA').font = ft_lab; r_ad += 1
            _header(ws_ad, ['MICROÁREA', 'N (ACAM + DOM)', ''], r_ad); r_ad += 1
            for fi, ma in enumerate(ad['por_microarea']):
                _row(ws_ad, [ma['microarea'], ma['n'], ''], r_ad,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r_ad += 1
        _widths(ws_ad, [42, 16, 12])

    # ── Aba 8: Saúde da Mulher ───────────────────────────────────────────────
    sm = dados.get('saude_mulher', {})
    if sm:
        ws8 = wb.create_sheet('Saúde da Mulher')
        ws8.sheet_view.showGridLines = False
        _title(ws8, f'SAÚDE DA MULHER · {mun}', 3)
        _title(ws8, 'Distribuição de registros femininos por faixa de vida', 3, row=2, fill=f_h2, font=ft_sub)
        r8 = 4
        _header(ws8, ['INDICADOR', 'N', '%'], r8); r8 += 1
        linhas8 = [
            ('Total de registros femininos', sm['total'], sm['pct']),
            ('Faixa reprodutiva (10–49 anos)', sm.get('faixa_reprodutiva', {}).get('n', ''),
             sm.get('faixa_reprodutiva', {}).get('pct', '')),
            ('Idosas (≥60 anos)', sm.get('idosas', {}).get('n', ''),
             sm.get('idosas', {}).get('pct', '')),
        ]
        if sm.get('gestantes'):
            linhas8.append(('Gestantes identificadas', sm['gestantes']['n'],
                            sm['gestantes']['taxa_observada_pct']))
        for fi, (label, val, pct) in enumerate(linhas8):
            _row(ws8, [label, val, pct], r8, fill=f_alt if fi % 2 == 0 else None,
                 aligns=[AL, AR, AC]); r8 += 1

        # Rastreamento — Citopatológico e Mamografia
        r8 += 1
        ws8.cell(row=r8, column=1, value='RASTREAMENTO ONCOLÓGICO').font = ft_lab; r8 += 1
        _header(ws8, ['INDICADOR', 'N ELEGÍVEIS', '%'], r8); r8 += 1
        gp8 = dados.get('grupos', {})
        for label8, key8 in [
            ('Mulheres 25–64 anos (citopatológico)', 'cito'),
            ('Mulheres 50–69 anos (mamografia)',      'mamografia'),
        ]:
            v8 = sm.get(key8, {})
            _row(ws8, [label8, v8.get('n', ''), v8.get('pct', '')], r8,
                 fill=f_alt if r8 % 2 == 0 else None, aligns=[AL, AR, AC]); r8 += 1

        # Gestantes e puérperas
        r8 += 1
        ws8.cell(row=r8, column=1, value='GESTANTES E PUÉRPERAS').font = ft_lab; r8 += 1
        _header(ws8, ['INDICADOR', 'N', '%'], r8); r8 += 1
        gest8 = gp8.get('gestantes', {})
        gar8  = gp8.get('gestantes_alto_risco', {})
        puerp8= gp8.get('puerperas', {})
        _row(ws8, ['Gestantes', gest8.get('n', 0), ''], r8,
             fill=f_alt if r8 % 2 == 0 else None, aligns=[AL, AR, AC]); r8 += 1
        _row(ws8, ['  ↳ Gestantes de alto risco', gar8.get('n', 0),
                   gar8.get('pct_gestantes', '')], r8, fill=f_alto, aligns=[AL, AR, AC]); r8 += 1
        _row(ws8, ['Puérperas', puerp8.get('n', 0), ''], r8,
             fill=f_alt if r8 % 2 == 0 else None, aligns=[AL, AR, AC]); r8 += 1

        _widths(ws8, [44, 14, 12]); ws8.freeze_panes = 'A5'

    # ── Aba Gestantes ────────────────────────────────────────────────────────
    if dados.get('gestantes_listing'):
        _GEST_LABELS = {
            'equipe': 'Equipe', 'micro_area': 'Micro Área', 'funcionario': 'ACS/Funcionário',
            'nome': 'Nome do Paciente', 'cpf': 'CPF', 'cns': 'CNS',
            'nascimento': 'Dt. Nascimento', 'idade': 'Idade',
            'cel': 'Celular', 'tel_res': 'Tel. Residencial',
            'cond_saude': 'Condições de Saúde',
        }
        _GEST_WIDTHS = {
            'equipe': 18, 'micro_area': 12, 'funcionario': 24,
            'nome': 32, 'cpf': 14, 'cns': 16, 'nascimento': 13,
            'idade': 8, 'cel': 14, 'tel_res': 14, 'cond_saude': 44,
        }
        gest_list = dados['gestantes_listing']
        ws_g = wb.create_sheet('Gestantes')
        ws_g.sheet_view.showGridLines = False
        g_cols = list(gest_list[0].keys()) if gest_list else []
        ncg = len(g_cols)
        _title(ws_g, f'LISTAGEM DE GESTANTES · {mun}', ncg)
        _title(ws_g, f'Total identificado: {len(gest_list)} gestante(s) → Fonte: {dados["fonte"]}',
               ncg, row=2, fill=f_h2, font=ft_sub)
        rg = 4
        _header(ws_g, [_GEST_LABELS.get(c, c.replace('_', ' ').title()) for c in g_cols], rg); rg += 1
        _pat_ar_g = r'HIPERTENSO|DIABETES|RENAL|CARD[IÍ]AC|CARDIOVASC|C[AÂ]NCER|AVC|DERRAME|INFARTO|IAM'
        for fi, rec in enumerate(gest_list):
            vals = [rec.get(c, '') for c in g_cols]
            cond_val = rec.get('cond_saude', '')
            is_ar = bool(re.search(_pat_ar_g, str(cond_val).upper()))
            row_fill = f_alto if is_ar else (f_alt if fi % 2 == 0 else None)
            _row(ws_g, vals, rg, fill=row_fill); rg += 1
        for ci, col in enumerate(g_cols, 1):
            ws_g.column_dimensions[get_column_letter(ci)].width = _GEST_WIDTHS.get(col, 15)
        ws_g.freeze_panes = 'A5'

    # ── Aba 9: Por Microárea ─────────────────────────────────────────────────
    if dados.get('microareas'):
        ws9 = wb.create_sheet('Por Microárea')
        ws9.sheet_view.showGridLines = False
        _title(ws9, f'ANÁLISE POR MICROÁREA · {mun}', 10)
        _title(ws9, 'Completude cadastral por microárea — ordenado por maior pendência', 10, row=2, fill=f_h2, font=ft_sub)
        r9 = 4
        _header(ws9, ['MICROÁREA', 'EQUIPE', 'ACS', 'TOTAL', 'S/CPF', '%', 'S/CNS', '%', 'S/END', '%'], r9); r9 += 1
        for fi, ma in enumerate(dados['microareas']):
            _row(ws9, [ma['microarea'], ma['equipe'], ma['acs'], ma['total'],
                       ma['s_cpf'], ma['s_cpf_pct'], ma['s_cns'], ma['s_cns_pct'],
                       ma['s_end'], ma['s_end_pct']],
                 r9, fill=f_alt if fi % 2 == 0 else None,
                 aligns=[AL, AL, AL, AR, AR, AC, AR, AC, AR, AC]); r9 += 1
        _widths(ws9, [18, 16, 24, 8, 8, 6, 8, 6, 8, 6]); ws9.freeze_panes = 'A5'

    # ── Aba 10: Por ACS ──────────────────────────────────────────────────────
    if dados.get('acs'):
        ws10 = wb.create_sheet('Por ACS')
        ws10.sheet_view.showGridLines = False
        _title(ws10, f'ANÁLISE POR ACS · {mun}', 8)
        _title(ws10, 'Carga de cadastros e completude por Agente Comunitário de Saúde', 8, row=2, fill=f_h2, font=ft_sub)
        r10 = 4
        _header(ws10, ['ACS', 'EQUIPE', 'MICROÁREA', 'TOTAL', 'S/CPF', '%', 'S/CNS', '%'], r10); r10 += 1
        for fi, a in enumerate(dados['acs']):
            _row(ws10, [a['nome'], a['equipe'], a.get('microarea', ''), a['total'],
                        a['s_cpf'], a['s_cpf_pct'], a['s_cns'], a['s_cns_pct']],
                 r10, fill=f_alt if fi % 2 == 0 else None,
                 aligns=[AL, AL, AL, AR, AR, AC, AR, AC]); r10 += 1
        _widths(ws10, [28, 18, 12, 8, 8, 6, 8, 6]); ws10.freeze_panes = 'A5'

    # ── Aba 11: Temporal ─────────────────────────────────────────────────────
    temp = dados.get('temporal', {})
    if temp:
        ws11 = wb.create_sheet('Análise Temporal')
        ws11.sheet_view.showGridLines = False
        _title(ws11, f'ANÁLISE TEMPORAL DE CADASTROS · {mun}', 3)
        r11 = 3
        if temp.get('por_ano'):
            ws11.cell(row=r11, column=1, value='CADASTROS POR ANO DE CRIAÇÃO').font = ft_lab; r11 += 1
            _header(ws11, ['ANO', 'N', 'BARRA VISUAL'], r11); r11 += 1
            t_total = sum(x['n'] for x in temp['por_ano']) or 1
            for fi, item in enumerate(temp['por_ano']):
                bar_v = min(int(item['n'] / t_total * 40), 40)
                _row(ws11, [item['ano'], item['n'], '█' * bar_v], r11,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AC, AR, AL]); r11 += 1
        r11 += 1
        ws11.cell(row=r11, column=1, value='INDICADORES DE ATUALIZAÇÃO').font = ft_lab; r11 += 1
        _header(ws11, ['INDICADOR', 'N', '%'], r11); r11 += 1
        for label, key_n, key_pct in [
            ('Nunca editados desde a criação', 'nunca_editados', 'nunca_editados_pct'),
            ('Criados há +2 anos sem atualização', 'antigos_sem_edicao', 'antigos_sem_edicao_pct'),
        ]:
            if key_n in temp:
                _row(ws11, [label, temp[key_n], temp[key_pct]], r11, aligns=[AL, AR, AC]); r11 += 1
        _widths(ws11, [40, 12, 45])

    # ── Aba 12: Determinantes Sociais ────────────────────────────────────────
    det = dados.get('determinantes_sociais', {})
    if det:
        ws12 = wb.create_sheet('Determinantes Sociais')
        ws12.sheet_view.showGridLines = False
        _title(ws12, f'DETERMINANTES SOCIAIS · {mun}', 3)
        r12 = 3
        if det.get('escolaridade'):
            ws12.cell(row=r12, column=1, value='ESCOLARIDADE').font = ft_lab; r12 += 1
            _header(ws12, ['NÍVEL DE ESCOLARIDADE', 'N', '%'], r12); r12 += 1
            for fi, e in enumerate(det['escolaridade']):
                _row(ws12, [e['label'], e['n'], e['pct']], r12,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r12 += 1
            r12 += 1
        if det.get('trabalho'):
            ws12.cell(row=r12, column=1, value='SITUAÇÃO NO MERCADO DE TRABALHO').font = ft_lab; r12 += 1
            _header(ws12, ['SITUAÇÃO', 'N', '%'], r12); r12 += 1
            for fi, t in enumerate(det['trabalho']):
                _row(ws12, [t['label'], t['n'], t['pct']], r12,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r12 += 1
            r12 += 1
        if det.get('bairro'):
            ws12.cell(row=r12, column=1, value='BAIRROS (TOP 20)').font = ft_lab; r12 += 1
            _header(ws12, ['BAIRRO', 'N', '%'], r12); r12 += 1
            for fi, b in enumerate(det['bairro']):
                _row(ws12, [b['label'], b['n'], b['pct']], r12,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AR, AC]); r12 += 1
        _widths(ws12, [45, 12, 10])

    # ── Aba 13: Inconsistências Avançadas ────────────────────────────────────
    incons = dados.get('inconsistencias', {})
    if incons:
        ws13 = wb.create_sheet('Inconsistências')
        ws13.sheet_view.showGridLines = False
        _title(ws13, f'INCONSISTÊNCIAS AVANÇADAS · {mun}', 3)
        r13 = 3
        _header(ws13, ['TIPO DE INCONSISTÊNCIA', 'N', '%'], r13); r13 += 1
        linhas13 = [
            ('CPF duplicado na base', 'cpf_duplicado', 'cpf_dup_pct'),
            ('CNS duplicado na base', 'cns_duplicado', 'cns_dup_pct'),
            ('Idades improváveis (>120 anos)', 'idade_improvavel', 'idade_improvavel_pct'),
        ]
        for label, key_n, key_pct in linhas13:
            if key_n in incons:
                pn = incons[key_n] / dados['total'] * 100 if dados['total'] else 0
                fs = f_alto if pn > 5 else (f_med if pn > 1 else None)
                _row(ws13, [label, incons[key_n], incons[key_pct]], r13,
                     fill=fs, aligns=[AL, AR, AC]); r13 += 1
        if incons.get('cpf_dup_detalhes'):
            r13 += 1
            ws13.cell(row=r13, column=1, value='DETALHAMENTO — CPFs DUPLICADOS (até 20 casos)').font = ft_lab; r13 += 1
            _header(ws13, ['CPF', 'Nº OCORRÊNCIAS', 'NOMES ENCONTRADOS'], r13); r13 += 1
            for fi, d in enumerate(incons['cpf_dup_detalhes']):
                nomes = ' / '.join(d.get('nomes', []))
                _row(ws13, [d['cpf'], d['ocorrencias'], nomes], r13,
                     fill=f_alt if fi % 2 == 0 else None, aligns=[AL, AC, AL]); r13 += 1
        _widths(ws13, [40, 12, 50])

    # ── Aba 14: Saúde da Criança ─────────────────────────────────────────────
    gpc = dados.get('grupos', {})
    if gpc.get('criancas_02') or gpc.get('criancas_05') or gpc.get('criancas_nutricao'):
        ws14 = wb.create_sheet('Saúde da Criança')
        ws14.sheet_view.showGridLines = False
        _title(ws14, f'SAÚDE DA CRIANÇA · {mun}', 3)
        _title(ws14, 'Distribuição de crianças cadastradas por faixa e estado nutricional (0–5 anos)',
               3, row=2, fill=f_h2, font=ft_sub)
        r14 = 4

        _header(ws14, ['INDICADOR', 'N', '% DO TOTAL'], r14); r14 += 1
        linhas14 = [
            ('Crianças 0–2 anos',  gpc.get('criancas_02', {}).get('n', ''), gpc.get('criancas_02', {}).get('pct', '')),
            ('Crianças 0–5 anos',  gpc.get('criancas_05', {}).get('n', ''), gpc.get('criancas_05', {}).get('pct', '')),
            ('Crianças 0–4 anos (ciclo PNAB)', gpc.get('criancas', {}).get('n', ''), gpc.get('criancas', {}).get('pct', '')),
        ]
        for fi14, (lbl14, v14, p14) in enumerate(linhas14):
            _row(ws14, [lbl14, v14, p14], r14,
                 fill=f_alt if fi14 % 2 == 0 else None, aligns=[AL, AR, AC]); r14 += 1

        nut = gpc.get('criancas_nutricao')
        if nut:
            r14 += 1
            ws14.cell(row=r14, column=1,
                      value='ESTADO NUTRICIONAL — CRIANÇAS 0–5 ANOS (registradas)').font = ft_lab
            r14 += 1
            _header(ws14, ['CATEGORIA', 'N', '% DAS CRIANÇAS 0–5'], r14); r14 += 1
            nut_linhas = [
                ('Baixo peso',                     nut['baixo_peso']),
                ('Sobrepeso',                       nut['sobrepeso']),
                ('Obesidade',                       nut['obesas']),
                ('Sem informação de peso', nut['sem_info']),
            ]
            fills_nut = [f_alto, f_med, f_alto, PatternFill('solid', fgColor='E3F2FD')]
            for fi14, (lbl14, d14) in enumerate(nut_linhas):
                _row(ws14, [lbl14, d14['n'], d14['pct']], r14,
                     fill=fills_nut[fi14], aligns=[AL, AR, AC]); r14 += 1
            r14 += 1
            nota = ws14.cell(row=r14, column=1,
                value='Nota: "Sem informação" = crianças sem nenhum marcador nutricional registrado no cadastro individual. '
                      'Não significa necessariamente ausência de dado clínico.')
            nota.font = Font(name='Calibri', italic=True, size=8, color='606060')
            ws14.merge_cells(f'A{r14}:C{r14}')
        _widths(ws14, [52, 12, 20]); ws14.freeze_panes = 'A5'

    wb.save(out_path)
    return True


# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO DOCX (via Node.js + docx-js)
# ══════════════════════════════════════════════════════════════════════════════
def gerar_docx(dados: dict, out_path: Path, header_img_path: Path):
    """Gera o DOCX chamando o script Node.js com os dados em JSON."""
    script_dir = out_path.parent
    js_path = Path(__file__).parent / 'gerar_docx.js'

    if not js_path.exists():
        print(f"  ⚠  Script de geração DOCX não encontrado: {js_path}")
        print(f"    Coloque o arquivo 'gerar_docx.js' na mesma pasta do gerar_relatorio.py")
        return False

    # Verificar se Node.js está disponível
    try:
        subprocess.run(['node', '--version'], capture_output=True, check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("  ⚠  Node.js não encontrado. Instale o Node.js para gerar o DOCX.")
        print("    https://nodejs.org/")
        return False

    # Verificar se o pacote docx está instalado
    try:
        subprocess.run(['node', '-e', 'require("docx")'], capture_output=True, check=True)
    except subprocess.CalledProcessError:
        print("  ⚙ Instalando pacote 'docx' para Node.js...")
        try:
            subprocess.run(['npm', 'install', '-g', 'docx'], capture_output=True, check=True)
        except (FileNotFoundError, subprocess.CalledProcessError):
            print("  ⚠  npm não encontrado. Instale o Node.js com npm para gerar o DOCX.")
            print("    https://nodejs.org/")
            return False

    # Salvar JSON temporário
    json_path = script_dir / '_temp_dados_relatorio.json'
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    try:
        result = subprocess.run(
            ['node', str(js_path), str(json_path), str(out_path), str(header_img_path)],
            capture_output=True, text=True, encoding='utf-8'
        )
        if result.returncode == 0:
            return True
        else:
            print(f"  ⚠  Erro na geração do DOCX: {result.stderr}")
            return False
    finally:
        # Limpar JSON temporário
        if json_path.exists():
            json_path.unlink()


# ══════════════════════════════════════════════════════════════════════════════
#  DETECÇÃO DE DUPLICATAS
# ══════════════════════════════════════════════════════════════════════════════
def _normalizar_nome(nome) -> str:
    """Remove acentos e pontuação, deixa maiúsculas e espaço simples para comparação fuzzy."""
    import unicodedata as _uc
    if pd.isna(nome) or not str(nome).strip():
        return ''
    s = _uc.normalize('NFKD', str(nome))
    s = ''.join(c for c in s if not _uc.combining(c))
    s = re.sub(r'[^A-Za-z\s]', '', s).upper()
    return re.sub(r'\s+', ' ', s).strip()


def _tem_saida_obito_mudanca(s) -> bool:
    """True se o valor de 'saida' indica óbito ou mudança de território."""
    if pd.isna(s):
        return False
    su = str(s).upper()
    return ('BITO' in su) or ('MUDAN' in su)


def detectar_duplicatas(df: pd.DataFrame) -> pd.DataFrame:
    """Detecta prováveis registros duplicados por:
      - CPF idêntico (excluindo placeholders)
      - CNS idêntico
      - Nome muito similar (≥80%) + mesma data de nascimento
    Retorna DataFrame com grupos numerados, motivo da suspeita e datas de criação/edição.
    Datas com >100 registros iguais são ignoradas no fuzzy (provavelmente placeholder).

    Dentro de cada grupo de duplicatas, sinaliza óbito/mudança de território apenas
    quando o cadastro SEM atualização mais recente traz essa informação e o cadastro
    mais recente do grupo não a traz — nesse caso é preciso conferir manualmente
    (o registro mais recente pode não estar refletindo o óbito/mudança).
    """
    from difflib import SequenceMatcher

    cols_out = [c for c in [
        'equipe', 'micro_area', 'funcionario', 'nome', 'cpf', 'cns',
        'nascimento', 'sexo', 'saida', 'dt_criacao', 'dt_edicao',
    ] if c in df.columns]

    df_w = df[cols_out].copy().reset_index(drop=True)
    n = len(df_w)

    parent = list(range(n))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(x: int, y: int) -> None:
        parent[find(x)] = find(y)

    motivo_map: dict = {}

    def add_motivo(i: int, j: int, mot: str) -> None:
        key = (min(i, j), max(i, j))
        motivo_map.setdefault(key, set()).add(mot)
        union(i, j)

    # 1. CPF idêntico
    if 'cpf' in df_w.columns:
        for _, g in df_w[~is_blank_cpf(df_w['cpf'])].groupby('cpf'):
            idxs = g.index.tolist()
            for k in range(1, len(idxs)):
                add_motivo(idxs[0], idxs[k], 'CPF idêntico')

    # 2. CNS idêntico
    if 'cns' in df_w.columns:
        for _, g in df_w[~is_blank(df_w['cns'])].groupby('cns'):
            idxs = g.index.tolist()
            for k in range(1, len(idxs)):
                add_motivo(idxs[0], idxs[k], 'CNS idêntico')

    # 3. Nome fuzzy + mesma data de nascimento
    if 'nome' in df_w.columns and 'nascimento' in df_w.columns:
        nomes_norm = df_w['nome'].apply(_normalizar_nome)
        for _, g in df_w.groupby('nascimento', dropna=True):
            idxs = g.index.tolist()
            if len(idxs) < 2 or len(idxs) > 100:
                continue  # pula datas placeholder (ex: 01/01/1900 com centenas de registros)
            for a in range(len(idxs)):
                for b in range(a + 1, len(idxs)):
                    na, nb = nomes_norm[idxs[a]], nomes_norm[idxs[b]]
                    if not na or not nb:
                        continue
                    ratio = SequenceMatcher(None, na, nb).ratio()
                    if ratio >= 0.80:
                        add_motivo(idxs[a], idxs[b],
                                   f'Nome similar ({ratio:.0%}) + mesma Dt. Nascimento')

    # Agrupar por union-find
    groups: dict = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)

    dup_groups = [(root, idxs) for root, idxs in groups.items() if len(idxs) > 1]
    if not dup_groups:
        return pd.DataFrame()

    rows = []
    for grupo_id, (_, idxs) in enumerate(dup_groups, 1):
        motivos: set = set()
        for a in range(len(idxs)):
            for b in range(a + 1, len(idxs)):
                motivos.update(motivo_map.get(
                    (min(idxs[a], idxs[b]), max(idxs[a], idxs[b])), set()))
        mot_str = ' | '.join(sorted(motivos))

        # Óbito/mudança de território: só alerta se o cadastro sem edição mais
        # recente do grupo tem a informação e o mais recente não tem.
        alerta_por_idx: dict = {}
        if 'saida' in df_w.columns and 'dt_edicao' in df_w.columns:
            dt_parsed = {i: pd.to_datetime(df_w.at[i, 'dt_edicao'], dayfirst=True, errors='coerce')
                         for i in idxs}
            validos = [i for i in idxs if pd.notna(dt_parsed[i])]
            if validos:
                mais_recente = max(validos, key=lambda i: dt_parsed[i])
                saida_recente = df_w.at[mais_recente, 'saida']
                if not _tem_saida_obito_mudanca(saida_recente):
                    for i in idxs:
                        if i == mais_recente:
                            continue
                        saida_i = df_w.at[i, 'saida']
                        if _tem_saida_obito_mudanca(saida_i):
                            tipo = 'ÓBITO' if 'BITO' in str(saida_i).upper() else 'MUDANÇA DE TERRITÓRIO'
                            alerta_por_idx[i] = (
                                f"Cadastro de {df_w.at[i, 'dt_edicao']} consta como {tipo}, mas o "
                                f"cadastro mais recente do grupo ({df_w.at[mais_recente, 'dt_edicao']}) "
                                f"não traz essa informação — conferir antes de excluir/mesclar."
                            )

        for i in idxs:
            row: dict = {'Grupo': grupo_id, 'Motivo da Suspeita': mot_str}
            row.update({col: df_w.at[i, col] for col in cols_out})
            row['alerta_obito_mudanca'] = alerta_por_idx.get(i, '')
            rows.append(row)

    result = pd.DataFrame(rows)
    sort_cols = ['Grupo'] + [c for c in ['equipe', 'nome'] if c in result.columns]
    return result.sort_values(sort_cols).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
#  ORQUESTRADOR PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def gerar_relatorio(filepath: str):
    municipio = detect_municipio(filepath)
    hoje = datetime.now()
    out_dir = Path(filepath).parent
    base_name = f"Relatório - {municipio} - {hoje.strftime('%m.%Y')}"
    suf = f"{municipio} - {hoje.strftime('%m.%Y')}"

    print(f"\nCarregando dados... {filepath}")
    df = load_data(filepath)

    # ── 1. Extrair dados estruturados ─────────────────────────────────────
    dados = extrair_dados(df, municipio, filepath, hoje)

    # ── Detectar possíveis duplicatas (usa df completo, incluindo inativos) ──
    # Feito antes da geração do Excel para que a contagem por equipe entre
    # na aba "Completude por Equipe".
    print("\nDetectando possíveis duplicatas...")
    df_dup = detectar_duplicatas(df)
    dup_por_equipe: dict = {}
    if not df_dup.empty:
        if 'equipe' in df_dup.columns:
            for eq, g in df_dup.groupby('equipe', dropna=False):
                dup_por_equipe['Sem equipe' if pd.isna(eq) else str(eq).title()] = len(g)
        p_dup = out_dir / f"Possíveis Duplicatas - {suf}.xlsx"
        _save_listing_xlsx(df_dup, p_dup,
                           titulo=f'Possíveis Duplicatas · {municipio} · {hoje.strftime("%m/%Y")}',
                           group_col='Grupo')
        n_grupos = int(df_dup['Grupo'].nunique())
        print(f"✓ {len(df_dup)} registros em {n_grupos} grupos suspeitos → {p_dup.name}")
    else:
        print("  Nenhuma duplicata detectada.")

    if 'equipes' in dados:
        for eq_item in dados['equipes']:
            n_dup = len(df_dup) if eq_item['nome'] == 'TOTAL' else dup_por_equipe.get(eq_item['nome'], 0)
            eq_item['dup'] = n_dup
            eq_item['dup_pct'] = fmt_pct_int(n_dup, eq_item['total']) if eq_item['total'] else '0%'

    # ── 2. Gerar e salvar TXT ─────────────────────────────────────────────
    txt = gerar_txt(dados)
    relatorio_path = out_dir / f"{base_name}.txt"
    with open(relatorio_path, 'w', encoding='utf-8') as f:
        f.write(txt)
    print(txt)
    print(f"\n✓ Relatório TXT salvo: {relatorio_path}")

    # ── 2b. Salvar JSON estruturado (usado pelo gerador Word) ─────────────
    json_path = out_dir / f"{base_name}.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    print(f"✓ Dados JSON salvos:   {json_path}")

    # ── 2c. Gerar planilha Excel analítica ────────────────────────────────
    excel_path = out_dir / f"{base_name}.xlsx"
    print("\nGerando planilha Excel...")
    if gerar_excel_relatorio(dados, excel_path):
        print(f"✓ Planilha Excel salva: {excel_path}")
    else:
        print("  ⚠  Planilha Excel não foi gerada.")

    # ── 3. Gerar DOCX ────────────────────────────────────────────────────
    docx_path = out_dir / f"{base_name}.docx"
    header_img = Path(__file__).parent / 'assets' / 'header_prima.png'

    print("\nGerando relatório DOCX...")
    if gerar_docx(dados, docx_path, header_img):
        print(f"✓ Relatório DOCX salvo: {docx_path}")
    else:
        print("  ⚠  Relatório DOCX não foi gerado. Apenas o TXT está disponível.")

    # Excluir qualquer saída de cadastro (óbito, mudança de território, fora
    # de área, recusa etc.) das listagens nominais — quem não está mais na
    # área não precisa aparecer em listas de correção cadastral.
    if 'saida' in df.columns:
        df = df[is_blank(df['saida'])].copy()
    if 'micro_area' in df.columns:
        df = df[~df['micro_area'].str.contains('FORA', case=False, na=False)].copy()

    # ── 4. Gerar CSVs de irregulares ─────────────────────────────────────
    total = len(df)
    print("\nGerando planilhas Excel de irregulares...")
    csvs = {}

    if 'cpf' in df.columns:
        s = df[is_blank_cpf(df['cpf'])][['equipe','micro_area','funcionario','nome','cns']].sort_values(['equipe','micro_area','nome'])
        p = out_dir / f"Irregulares sem CPF - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Sem CPF'] = (len(s), p)

    if 'cns' in df.columns:
        s = df[is_blank(df['cns'])][['equipe','micro_area','funcionario','nome','cpf']].sort_values(['equipe','micro_area','nome'])
        p = out_dir / f"Irregulares sem CNS - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Sem CNS'] = (len(s), p)

    if 'cpf' in df.columns and 'cns' in df.columns:
        s = df[is_blank_cpf(df['cpf']) & is_blank(df['cns'])][['equipe','micro_area','funcionario','nome']].sort_values(['equipe','micro_area','nome'])
        p = out_dir / f"Irregulares sem CPF e CNS - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Sem CPF e CNS'] = (len(s), p)

    if 'vinc_familia' in df.columns:
        s = df[df['vinc_familia'].str.strip() == 'NÃO'][['equipe','micro_area','funcionario','nome','cpf']].sort_values(['equipe','micro_area','nome'])
        p = out_dir / f"Irregulares sem Vínculo Familiar - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Não vinculado à família'] = (len(s), p)

    if 'dt_edicao_p' in df.columns:
        s = df[df['dt_edicao_p'] < pd.Timestamp('2025-04-27')][['equipe','micro_area','funcionario','nome','cpf','dt_edicao']].sort_values(['equipe','dt_edicao'])
        p = out_dir / f"Cadastros Desatualizados - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Desatualizados'] = (len(s), p)

    if 'cel' in df.columns and 'tel_res' in df.columns:
        s = df[is_blank(df['cel']) & is_blank(df['tel_res'])][['equipe','micro_area','funcionario','nome','cpf']].sort_values(['equipe','micro_area','nome'])
        p = out_dir / f"Irregulares sem Telefone - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Sem Telefone'] = (len(s), p)

    if 'escolaridade' in df.columns:
        s = df[is_blank(df['escolaridade'])][['equipe','micro_area','funcionario','nome','cpf']].sort_values(['equipe','micro_area','nome'])
        p = out_dir / f"Irregulares sem Escolaridade - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Sem Escolaridade'] = (len(s), p)

    if 'bairro' in df.columns:
        s = df[is_blank(df['bairro'])][['equipe','micro_area','funcionario','nome','cpf']].sort_values(['equipe','micro_area','nome'])
        p = out_dir / f"Irregulares sem Bairro - {suf}.xlsx"
        _save_listing_xlsx(s, p)
        csvs['Sem Bairro'] = (len(s), p)

    # ── 5. CSVs de grupos prioritários ───────────────────────────────────
    _cols_nom = [c for c in ['equipe', 'micro_area', 'funcionario', 'nome', 'cpf', 'cns',
                              'nascimento', 'idade', 'cel', 'tel_res'] if c in df.columns]
    print("\nGerando listagens de grupos prioritários...")
    csvs2 = {}

    if 'cond_saude' in df.columns:
        # Gestantes
        m = df['cond_saude'].str.contains('GESTANTE', na=False)
        if m.any():
            _cols_gest_lst = [c for c in _cols_nom + ['cond_saude'] if c in df.columns]
            s = df[m][_cols_gest_lst].sort_values(['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome'])
            p = out_dir / f"Listagem de Gestantes - {suf}.xlsx"
            _save_listing_xlsx(s, p, titulo=f'Gestantes — {municipio} — {hoje.strftime("%m/%Y")}')
            csvs2['Gestantes'] = (len(s), p)

        # Gestantes de alto risco
        _pat_ar = r'HIPERTENSO|DIABETES|RENAL|CARD[IÍ]AC|CARDIOVASC|C[AÂ]NCER|AVC|DERRAME|INFARTO|IAM'
        m_ar = df['cond_saude'].str.contains('GESTANTE', na=False) & df['cond_saude'].str.contains(_pat_ar, na=False)
        if m_ar.any():
            s = df[m_ar][_cols_nom + ['cond_saude']].sort_values(['equipe', 'nome'] if 'equipe' in _cols_nom else ['nome'])
            p = out_dir / f"Gestantes de Alto Risco - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2['Gestantes alto risco'] = (len(s), p)

        # Puérperas
        m_pu = df['cond_saude'].str.contains(r'PU[EÉ]RPERA', na=False)
        if m_pu.any():
            s = df[m_pu][_cols_nom].sort_values(['equipe', 'nome'] if 'equipe' in _cols_nom else ['nome'])
            p = out_dir / f"Listagem de Puérperas - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2['Puérperas'] = (len(s), p)

        # Acamados e domiciliados (Atenção Domiciliar)
        m_ad = (df['cond_saude'].str.contains('ACAMADO', na=False)
                | df['cond_saude'].str.contains(r'DOMICILI[AO]', na=False))
        if m_ad.any():
            _cols_ad = [c for c in _cols_nom + ['cond_saude'] if c in df.columns]
            s = df[m_ad][_cols_ad].sort_values(['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome'])
            p = out_dir / f"Listagem de Acamados e Domiciliados - {suf}.xlsx"
            _save_listing_xlsx(s, p, titulo=f'Acamados e Domiciliados — {municipio} — {hoje.strftime("%m/%Y")}')
            csvs2['Acamados/Domiciliados'] = (len(s), p)

    # Citopatológico (mulheres 25–64 anos)
    if 'sexo' in df.columns and 'idade_num' in df.columns:
        m_fem = df['sexo'].str.strip().str.upper().isin(['FEMININO', 'F'])
        m_cito = m_fem & (df['idade_num'] >= 25) & (df['idade_num'] <= 64)
        if m_cito.any():
            s = df[m_cito][_cols_nom].sort_values(['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome'])
            p = out_dir / f"Citopatológico 25 a 64 anos - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2['Citopatológico (25–64 anos)'] = (len(s), p)

        # Mamografia (mulheres 50–69 anos)
        m_mamo = m_fem & (df['idade_num'] >= 50) & (df['idade_num'] <= 69)
        if m_mamo.any():
            s = df[m_mamo][_cols_nom].sort_values(['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome'])
            p = out_dir / f"Mamografia 50 a 69 anos - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2['Mamografia (50–69 anos)'] = (len(s), p)

    # Crianças 0–2 anos
    if 'idade_num' in df.columns:
        m_02 = (df['idade_num'] >= 0) & (df['idade_num'] <= 2)
        if m_02.any():
            s = df[m_02][_cols_nom].sort_values(['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome'])
            p = out_dir / f"Listagem de Crianças de 0 a 2 anos - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2['Crianças 0–2 anos'] = (len(s), p)

        # Crianças 0–5 anos
        m_05 = (df['idade_num'] >= 0) & (df['idade_num'] <= 5)
        if m_05.any():
            s = df[m_05][_cols_nom].sort_values(['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome'])
            p = out_dir / f"Listagem de Crianças de 0 a 5 anos - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2['Crianças 0–5 anos'] = (len(s), p)

        # Crianças 0–5 com estado nutricional
        if 'cond_saude' in df.columns:
            m_bp  = m_05 & df['cond_saude'].str.contains('BAIXO PESO', na=False)
            m_ob  = m_05 & df['cond_saude'].str.contains(r'OBES[OA]|OBESIDADE', na=False)
            m_sp  = m_05 & df['cond_saude'].str.contains('SOBREPESO', na=False)
            m_sem = m_05 & ~(df['cond_saude'].str.contains(r'BAIXO PESO|OBES[OA]|OBESIDADE|SOBREPESO', na=False))
            _ord = ['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome']
            _obs_sem_peso = (
                'Observação: a situação de peso (abaixo do peso, acima do peso ou peso adequado) é uma '
                'informação registrada no sistema Atenção Básica, na aba "Situação de Peso" da ficha da '
                'criança. Os cidadãos abaixo não têm esse campo preenchido no sistema.'
            )
            for lbl_c, mask_c, fname_c, titulo_c, obs_c in [
                ('Crianças com Baixo Peso',    m_bp,  f"Crianças com Baixo Peso - {suf}.xlsx", '', ''),
                ('Crianças com Obesidade',     m_ob,  f"Crianças com Obesidade - {suf}.xlsx", '', ''),
                ('Crianças sem Info de Peso',  m_sem, f"Crianças sem Info de Peso - {suf}.xlsx",
                 f'Crianças sem Info de Peso · {municipio} · {hoje.strftime("%m/%Y")}', _obs_sem_peso),
            ]:
                if mask_c.any():
                    s = df[mask_c][_cols_nom].sort_values(_ord)
                    p = out_dir / fname_c
                    _save_listing_xlsx(s, p, titulo=titulo_c, obs=obs_c)
                    csvs2[lbl_c] = (len(s), p)

    # Condições crônicas — um CSV por condição, com coluna mostrando todas as condições do paciente
    if 'cond_saude' in df.columns:
        _conds_csv = [
            ('HAS',             'HIPERTENSO',           r'HIPERTENSO',          'has'),
            ('Diabetes',        'DIABETES',              r'DIABETES',            'diabetes'),
            ('Obesidade',       'OBES/OBESIDADE',        r'OBES[OA]|OBESIDADE',  'obesidade'),
            ('Doença Cardíaca', 'CARDÍACA/CARDIOVASC',  r'CARD[IÍ]AC|CARDIOVASC','doenca_cardiaca'),
            ('Doença Renal',    'RENAL',                 r'RENAL',               'doenca_renal'),
            ('Câncer',          'CANCER',                r'C[AÂ]NCER',           'cancer'),
            ('Doença Mental',   'MENTAL',                r'MENTAL',              'doenca_mental'),
            ('AVC / Derrame',   'AVC|DERRAME',           r'AVC|DERRAME',         'avc_derrame'),
            ('Tuberculose',     'TUBERCULOSE',           r'TUBERCULOS[EI]',      'tuberculose'),
            ('Hanseníase',      'HANSEN',                r'HANSEN',              'hanseniase'),
            ('Fumante',         'FUMANTE',               r'FUMANTE',             'fumante'),
            ('Álcool',          'ALCOOL',                r'[AÁ]LCOOL',           'alcool'),
        ]

        # Pré-computa rótulo com todas as condições de cada paciente (ex: "HAS | Diabetes | Obesidade")
        def _label_conds(cond_str):
            if pd.isna(cond_str):
                return ''
            c = str(cond_str).upper()
            return ' | '.join(lbl for lbl, _, pat, _ in _conds_csv if re.search(pat, c))

        serie_label = df['cond_saude'].apply(_label_conds)
        _ord_cr = ['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome']

        for label_c, _, pat_c, fname_c in _conds_csv:
            mask_c = df['cond_saude'].str.contains(pat_c, na=False)
            if not mask_c.any():
                continue
            s = df[mask_c][_cols_nom].copy()
            s['condicoes_registradas'] = serie_label[mask_c]
            s = s.sort_values(_ord_cr)
            label_c_safe = re.sub(r'[\\/:*?"<>|]', '-', label_c)
            p = out_dir / f"Crônicos {label_c_safe} - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2[f'Crônicas — {label_c}'] = (len(s), p)

    # Idosos ≥60 anos com estratificação de risco
    if 'idade_num' in df.columns:
        m_id = df['idade_num'] >= 60
        if m_id.any():
            s = df[m_id][_cols_nom].copy()
            if 'cond_saude' in df.columns:
                s['condicoes_registradas'] = df[m_id]['cond_saude'].apply(_label_conds)
            s['estratificacao_risco'] = df[m_id].apply(
                lambda r: classificar_risco(r.get('cond_saude'), r.get('idade_num')), axis=1
            )
            _ord_id = ['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome']
            s = s.sort_values(_ord_id)
            p = out_dir / f"Listagem de Idosos 60 anos ou mais - {suf}.xlsx"
            _save_listing_xlsx(s, p)
            csvs2['Idosos (≥60 anos)'] = (len(s), p)

    # Beneficiários de auxílio do governo
    if 'auxilio' in df.columns:
        _nao_aux = {'não', 'nao', 'n', 'não possui', 'nao possui'}
        m_aux = ~is_blank(df['auxilio']) & ~df['auxilio'].str.strip().str.lower().isin(_nao_aux)
        _ord_aux = ['equipe', 'micro_area', 'nome'] if 'micro_area' in _cols_nom else ['equipe', 'nome']

        if m_aux.any():
            _cols_aux = [c for c in _cols_nom + ['beneficio'] if c in df.columns]
            s = df[m_aux][_cols_aux].sort_values(_ord_aux)
            p = out_dir / f"Listagem de Beneficiários de Auxílio do Governo - {suf}.xlsx"
            _save_listing_xlsx(s, p, titulo=f'Beneficiários de Auxílio do Governo — {municipio} — {hoje.strftime("%m/%Y")}')
            csvs2['Beneficiários de Auxílio do Governo'] = (len(s), p)

        # Uma planilha por tipo de benefício recebido
        if 'beneficio' in df.columns:
            m_ben = ~is_blank(df['beneficio'])
            for tipo in sorted(df.loc[m_ben, 'beneficio'].str.strip().unique()):
                m_tipo = m_ben & (df['beneficio'].str.strip() == tipo)
                s = df[m_tipo][_cols_nom].sort_values(_ord_aux)
                tipo_label = tipo.title()
                tipo_safe = re.sub(r'[\\/:*?"<>|]', '-', tipo_label)
                p = out_dir / f"Auxílio - {tipo_safe} - {suf}.xlsx"
                _save_listing_xlsx(s, p)
                csvs2[f'Auxílio — {tipo_label}'] = (len(s), p)

    print("\n  Irregulares:")
    for label, (n, p) in csvs.items():
        print(f"  ✓ {label:<35}: {n:,} registros → {p.name}")
    print("\n  Grupos prioritários:")
    for label, (n, p) in csvs2.items():
        print(f"  ✓ {label:<35}: {n:,} registros → {p.name}")

    print("\nPronto!")


# ── ENTRY POINT ───────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            filepath = filedialog.askopenfilename(
                title="Selecione o arquivo de Cadastro Individual",
                filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
            )
            if not filepath:
                print("Nenhum arquivo selecionado. Encerrando.")
                sys.exit(0)
        except Exception:
            print("Uso: python gerar_relatorio.py \"caminho/para/arquivo.xlsx\"")
            sys.exit(1)

    if not os.path.exists(filepath):
        print(f"Arquivo não encontrado: {filepath}")
        sys.exit(1)

    gerar_relatorio(filepath)
