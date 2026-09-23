"""Gera as páginas estáticas de site/ a partir de compilar_site.compilar().

Cada rota vira um ARQUIVO HTML de verdade (não é uma SPA com roteador
client-side): o mesmo esqueleto HTML/CSS/JS é usado em todas, mudando só o
JSON embutido (window.__DADOS__/window.__ROTA__) e o caminho relativo até
estilo.css/app.js. Isso é o que garante:
  - URL estável em qualquer hospedagem estática (GitHub Pages, Netlify);
  - abrir local com duplo-clique, sem servidor: os dados vêm embutidos no
    próprio HTML, não por fetch (fetch de arquivo local via file:// é
    bloqueado pelo navegador).

Roda pipeline/privacidade.py no final sobre site/ inteiro (HTML incluído)
antes de considerar a geração aprovada.
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.compilar_site import compilar  # noqa: E402
from pipeline.privacidade import verificar_pastas  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
SITE_DIR = RAIZ / 'site'

MOLDE_HTML = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{titulo}</title>
<link rel="stylesheet" href="{prefixo}estilo.css">
</head>
<body>
<div id="app"></div>
<script type="application/json" id="dados-pagina">{dados_json}</script>
<script type="application/json" id="rota-pagina">{rota_json}</script>
<script src="{prefixo}app.js"></script>
</body>
</html>
"""


def _json_para_script(obj) -> str:
    """json.dumps, mas escapando "</" — sem isso, um texto livre de config
    (ex.: a orientação de algum indicador) que por acaso contivesse
    "</script>" fecharia a tag prematuramente e quebraria a página."""
    return json.dumps(obj, ensure_ascii=False).replace('</', '<\\/')


def _escrever_pagina(pasta: Path, titulo: str, dados: dict, rota: dict, prefixo: str) -> None:
    pasta.mkdir(parents=True, exist_ok=True)
    rota_completa = {**rota, 'prefixo': prefixo}
    html = MOLDE_HTML.format(
        titulo=titulo,
        prefixo=prefixo,
        dados_json=_json_para_script(dados),
        rota_json=_json_para_script(rota_completa),
    )
    (pasta / 'index.html').write_text(html, encoding='utf-8')


def _copiar_assets() -> None:
    """site/ precisa da sua PRÓPRIA cópia de assets/ porque é essa pasta
    (não a raiz do repo) que vai para o GitHub Pages/Netlify."""
    destino = SITE_DIR / 'assets'
    destino.mkdir(parents=True, exist_ok=True)

    origem_logo = RAIZ / 'assets' / 'logo_prima.png'
    if origem_logo.exists():
        shutil.copy2(origem_logo, destino / 'logo_prima.png')

    origem_brasoes = RAIZ / 'assets' / 'brasoes'
    if origem_brasoes.exists():
        destino_brasoes = destino / 'brasoes'
        destino_brasoes.mkdir(exist_ok=True)
        for arq in origem_brasoes.iterdir():
            if arq.is_file():
                shutil.copy2(arq, destino_brasoes / arq.name)


def gerar() -> tuple[bool, list]:
    dados = compilar()
    _copiar_assets()

    _escrever_pagina(SITE_DIR, 'Painel de Cadastro — Visão da rede', dados, {'tipo': 'rede'}, '')
    _escrever_pagina(SITE_DIR / 'ranking', 'Painel de Cadastro — Ranking', dados, {'tipo': 'ranking'}, '../')
    _escrever_pagina(SITE_DIR / 'metodologia', 'Painel de Cadastro — Metodologia', dados, {'tipo': 'metodologia'}, '../')

    for municipio in dados['municipios']:
        _escrever_pagina(
            SITE_DIR / municipio['slug'],
            f"Painel de Cadastro — {municipio['municipio']}",
            dados, {'tipo': 'municipio', 'municipio': municipio['slug']}, '../',
        )
        for equipe in municipio['equipes']:
            _escrever_pagina(
                SITE_DIR / municipio['slug'] / equipe['slug'],
                f"Painel de Cadastro — {equipe['equipe']} — {municipio['municipio']}",
                dados,
                {'tipo': 'equipe', 'municipio': municipio['slug'], 'equipe': equipe['slug']},
                '../../',
            )

    violacoes = verificar_pastas([SITE_DIR, RAIZ / 'data'])
    return (len(violacoes) == 0, violacoes)


if __name__ == '__main__':
    ok, violacoes = gerar()
    if not ok:
        print(f'✗ Site gerado, mas REPROVADO na verificação de privacidade ({len(violacoes)} problema(s)):')
        for v in violacoes:
            print(f'  - {v}')
        raise SystemExit(1)
    print(f'✓ Site gerado em {SITE_DIR} e aprovado na verificação de privacidade.')
