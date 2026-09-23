"""Carregamento dos arquivos de configuração em config/*.yaml.

Nenhuma regra de negócio deve ficar fixa no código do painel — este módulo
só lê e devolve o que está nos YAMLs, sem valores padrão escondidos que
mascarariam uma configuração ausente.
"""
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parent.parent
CONFIG_DIR = RAIZ / 'config'


def _carregar_yaml(nome_arquivo: str):
    caminho = CONFIG_DIR / nome_arquivo
    with open(caminho, encoding='utf-8') as f:
        return yaml.safe_load(f)


def carregar_equipes() -> list[dict]:
    return _carregar_yaml('equipes.yaml')


def carregar_inconsistencias() -> list[dict]:
    return _carregar_yaml('inconsistencias.yaml')


def carregar_grupos() -> list[dict]:
    return _carregar_yaml('grupos.yaml')


def carregar_metas() -> dict:
    return _carregar_yaml('metas.yaml')


def carregar_caminhos() -> dict:
    return _carregar_yaml('caminhos.yaml')


def carregar_condicoes_saude() -> dict:
    return _carregar_yaml('condicoes_saude.yaml')


def buscar_equipe(municipio: str, nome_exportacao: str, equipes: list[dict] | None = None) -> dict | None:
    """Procura uma equipe cadastrada pelo município + nome como vem na
    coluna 'equipe' da exportação (comparação normalizada: sem diferenciar
    maiúsculas/minúsculas nem espaços nas pontas). Devolve None se não achar
    — quem chama decide se isso é um erro fatal ou só um alerta.
    """
    equipes = equipes if equipes is not None else carregar_equipes()
    alvo_mun = municipio.strip().casefold()
    alvo_nome = nome_exportacao.strip().casefold()
    for eq in equipes:
        if (eq['municipio'].strip().casefold() == alvo_mun
                and eq['nome_exportacao'].strip().casefold() == alvo_nome):
            return eq
    return None
