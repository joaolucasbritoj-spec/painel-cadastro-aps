"""Ponte para as funções de gerar_relatorio.py.

Por instrução explícita, a lógica desse script NÃO é reescrita aqui — só
importamos e reaproveitamos o que já existe (load_data, is_blank,
is_blank_cpf, detectar_duplicatas, _save_listing_xlsx, detect_municipio).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# gerar_relatorio.py troca sys.stdout/sys.stderr por um TextIOWrapper
# próprio assim que é importado (para acentuação no console do Windows).
# Isso dá dois problemas para quem só importa funções dele (nosso caso):
#   1. Empacotar o objeto de captura do pytest nesse wrapper deixa o pytest
#      sem conseguir fechá-lo sozinho depois ("I/O operation on closed
#      file", nem um teste chega a rodar) — por isso trocamos sys.stdout/
#      stderr pelos streams reais do processo (sys.__stdout__/__stderr__)
#      só durante o import, para o script empacotar esses, não o do pytest.
#   2. Um TextIOWrapper fecha o buffer que envolve quando é descartado pelo
#      coletor de lixo — como restauramos sys.stdout/stderr logo em
#      seguida, os wrappers criados pelo script ficariam sem nenhuma
#      referência e seriam descartados, fechando sys.__stdout__/__stderr__
#      DE VERDADE (o processo inteiro perde a saída). Por isso guardamos
#      uma referência viva para eles em _wrappers_gerar_relatorio, só para
#      nunca deixarem de existir — nunca os usamos de fato.
# Não reescrevemos a lógica do script para resolver isso: o contorno mora
# só aqui, na ponte de import.
_stdout_original, _stderr_original = sys.stdout, sys.stderr
sys.stdout, sys.stderr = sys.__stdout__, sys.__stderr__
from gerar_relatorio import (  # noqa: E402  (import depende do sys.path acima)
    detect_municipio,
    detectar_duplicatas,  # noqa: F401  (reexportado para quem importar de pipeline.legado)
    is_blank,
    is_blank_cpf,
    load_data,
    _save_listing_xlsx,
)
_wrappers_gerar_relatorio = (sys.stdout, sys.stderr)
sys.stdout, sys.stderr = _stdout_original, _stderr_original

import pandas as pd  # noqa: E402


def carregar_ativos(caminho: str) -> pd.DataFrame:
    """Carrega uma exportação bruta e devolve só os cadastros ativos e
    dentro de área — o mesmo filtro que gerar_relatorio.py aplica a partir
    da seção "Visão geral" (exclui óbito, mudança de território e qualquer
    outra saída de cadastro, além de "FORA DE ÁREA" em micro_area).

    Não usar para a checagem de duplicatas: detectar_duplicatas() precisa
    do DataFrame completo (com inativos), porque um par duplicado pode ter
    uma das entradas já marcada como saída — para isso, use load_data()
    diretamente.
    """
    df = load_data(caminho)
    if 'saida' in df.columns:
        df = df[is_blank(df['saida'])].copy()
    if 'micro_area' in df.columns:
        df = df[~df['micro_area'].str.contains('FORA', case=False, na=False)].copy()
    return df
