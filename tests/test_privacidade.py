"""Teste da verificação de privacidade — usa só dado FICTÍCIO gerado aqui
mesmo no teste, nunca um export real."""
import csv
import json

from pipeline.privacidade import verificar_arquivo


def test_csv_com_coluna_nome_e_reprovado(tmp_path):
    caminho = tmp_path / 'listagem.csv'
    with open(caminho, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['equipe', 'nome', 'cpf'])
        w.writerow(['Arranchadouro', 'Fulano de Tal', '11122233344'])

    violacoes = verificar_arquivo(caminho)
    tipos = {v.tipo for v in violacoes}
    assert 'coluna nominal' in tipos


def test_csv_agregado_legitimo_e_aprovado(tmp_path):
    """Mesmo esquema de data/historico.csv — não pode disparar falso positivo,
    mesmo com indicador_id valendo "sem_cpf"/"sem_cns" (são VALORES, não
    nomes de coluna)."""
    caminho = tmp_path / 'historico.csv'
    with open(caminho, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['competencia', 'municipio', 'ine', 'equipe', 'indicador_id', 'valor', 'total_cadastrados'])
        w.writerow(['2026-09', 'Santa Maria Madalena', '', 'arranchadouro', 'sem_cpf', '2715', '5664'])
        w.writerow(['2026-09', 'Santa Maria Madalena', '', 'arranchadouro', 'sem_cns', '521', '5664'])

    assert verificar_arquivo(caminho) == []


def test_csv_com_valor_com_cara_de_cpf_e_reprovado(tmp_path):
    """Mesmo com cabeçalho inofensivo, um valor de 11 dígitos deve acender
    o alerta — defesa contra dado nominal vazado num campo mal nomeado."""
    caminho = tmp_path / 'exportacao_suspeita.csv'
    with open(caminho, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['equipe', 'documento'])
        w.writerow(['Triunfo', '123.456.789-00'])

    violacoes = verificar_arquivo(caminho)
    assert any('CPF' in v.tipo for v in violacoes)


def test_json_com_chave_nome_e_reprovado(tmp_path):
    caminho = tmp_path / 'data.json'
    caminho.write_text(json.dumps({'equipes': [{'nome': 'Fulano de Tal'}]}), encoding='utf-8')

    violacoes = verificar_arquivo(caminho)
    assert any(v.tipo == 'chave nominal' for v in violacoes)


def test_json_agregado_legitimo_e_aprovado(tmp_path):
    """Esquema real de site/data.json: id de indicador como VALOR (nunca
    como chave), "equipe"/"municipio"/"titulo" em vez de "nome"."""
    caminho = tmp_path / 'data.json'
    caminho.write_text(json.dumps({
        'municipios': [{
            'municipio': 'Santa Maria Madalena',
            'equipes': [{
                'equipe': 'Arranchadouro',
                'serie': [{
                    'competencia': '2026-09',
                    'indicadores': [
                        {'id': 'sem_cpf', 'valor': 2715, 'ausente': False},
                        {'id': 'sem_cns', 'valor': 521, 'ausente': False},
                    ],
                }],
            }],
        }],
    }, ensure_ascii=False), encoding='utf-8')

    assert verificar_arquivo(caminho) == []


def test_json_com_valor_com_cara_de_cns_e_reprovado(tmp_path):
    caminho = tmp_path / 'data.json'
    caminho.write_text(json.dumps({'observacao': '123456789012345'}), encoding='utf-8')

    violacoes = verificar_arquivo(caminho)
    assert any('CNS' in v.tipo for v in violacoes)


def test_html_com_dado_embutido_nominal_e_reprovado(tmp_path):
    """site/*/index.html embute os dados direto no HTML, em
    <script type="application/json"> (para abrir local sem servidor) — a
    verificação precisa varrer esse bloco também, não só o JSON de origem."""
    caminho = tmp_path / 'index.html'
    caminho.write_text(
        '<html><body>'
        '<script type="application/json" id="dados-pagina">{"equipes": [{"nome": "Fulano de Tal"}]}</script>'
        '<script type="application/json" id="rota-pagina">{"tipo": "rede"}</script>'
        '</body></html>',
        encoding='utf-8',
    )
    violacoes = verificar_arquivo(caminho)
    assert any(v.tipo == 'chave nominal' for v in violacoes)


def test_html_com_dado_embutido_legitimo_e_aprovado(tmp_path):
    caminho = tmp_path / 'index.html'
    caminho.write_text(
        '<html><body>'
        '<script type="application/json" id="dados-pagina">{"municipios": [{"municipio": "Santa Maria Madalena", '
        '"equipes": [{"equipe": "Arranchadouro"}]}]}</script>'
        '<script type="application/json" id="rota-pagina">{"tipo": "rede", "prefixo": ""}</script>'
        '</body></html>',
        encoding='utf-8',
    )
    assert verificar_arquivo(caminho) == []
