"""Verificação de privacidade (regra inegociável do projeto): aborta a
publicação se encontrar qualquer coluna/campo nominal dentro de site/ ou
data/. O painel publicado só pode conter números agregados por equipe e
por microárea — nunca nome, CPF, CNS, data de nascimento ou nome de
funcionário/ACS (o nível mais baixo permitido é a microárea).

Verifica dois jeitos de vazamento:
  1. Nome de coluna/campo suspeito (CSV: cabeçalho; JSON: qualquer chave,
     em qualquer profundidade) — comparação normalizada (sem acento, sem
     maiúsculas, sem separadores), por SUBSTRING: é melhor sinalizar
     demais do que deixar passar um "nome_do_paciente" por não bater
     exatamente com "nome".
  2. Valor com CARA de CPF/CNS (11 ou 15 dígitos, ignorando pontuação),
     mesmo que a chave/coluna pareça inofensiva — defesa extra contra um
     dado nominal vazar num campo mal nomeado.
"""
from __future__ import annotations

import csv
import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

# Os 4 nomes citados explicitamente na regra de privacidade do projeto, mais
# "funcionario" (ACS) por causa da regra "menor nível é a microárea, nunca
# nome de funcionário". NÃO inclui termos como "endereco"/"logradouro":
# eles colidiriam com o próprio id do indicador "sem_endereco", que É
# agregado e PODE aparecer no painel — o dado bruto de endereço nunca chega
# perto de site/ ou data/ para começo de conversa, então bloquear a palavra
# não protegeria nada e só geraria falso positivo.
TERMOS_PROIBIDOS = ['nome', 'cpf', 'cns', 'nascimento', 'funcionario']

_RE_CPF = re.compile(r'^\d{11}$')
_RE_CNS = re.compile(r'^\d{15}$')


def _normalizar(texto: str) -> str:
    s = unicodedata.normalize('NFKD', str(texto))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'[^a-z0-9]', '', s.lower())


def _campo_suspeito(nome_campo: str) -> str | None:
    normalizado = _normalizar(nome_campo)
    for termo in TERMOS_PROIBIDOS:
        if termo in normalizado:
            return termo
    return None


def _valor_com_cara_de_documento(valor) -> str | None:
    if not isinstance(valor, str):
        return None
    digitos = re.sub(r'\D', '', valor)
    if _RE_CPF.fullmatch(digitos):
        return 'CPF'
    if _RE_CNS.fullmatch(digitos):
        return 'CNS'
    return None


@dataclass
class Violacao:
    arquivo: str
    tipo: str
    detalhe: str

    def __str__(self) -> str:
        return f'{self.arquivo}: {self.tipo} — {self.detalhe}'


def _checar_csv(caminho: Path) -> list[Violacao]:
    violacoes = []
    with open(caminho, encoding='utf-8', newline='') as f:
        leitor = csv.reader(f)
        cabecalho = next(leitor, [])
        for campo in cabecalho:
            termo = _campo_suspeito(campo)
            if termo:
                violacoes.append(Violacao(str(caminho), 'coluna nominal', f'coluna "{campo}" (contém "{termo}")'))
        if violacoes:
            return violacoes  # cabeçalho já reprovado, não precisa escanear valor a valor
        for linha in leitor:
            for valor in linha:
                tipo_doc = _valor_com_cara_de_documento(valor)
                if tipo_doc:
                    violacoes.append(Violacao(str(caminho), f'valor com cara de {tipo_doc}', 'valor numérico suspeito numa coluna que deveria ser agregada'))
    return violacoes


def _checar_json_recursivo(obj, caminho: Path, trilha: str = '') -> list[Violacao]:
    violacoes = []
    if isinstance(obj, dict):
        for chave, valor in obj.items():
            termo = _campo_suspeito(str(chave))
            if termo:
                violacoes.append(Violacao(str(caminho), 'chave nominal', f'chave "{trilha}{chave}" (contém "{termo}")'))
            violacoes += _checar_json_recursivo(valor, caminho, f'{trilha}{chave}.')
    elif isinstance(obj, list):
        for item in obj:
            violacoes += _checar_json_recursivo(item, caminho, trilha)
    else:
        tipo_doc = _valor_com_cara_de_documento(obj)
        if tipo_doc:
            violacoes.append(Violacao(str(caminho), f'valor com cara de {tipo_doc}', f'em "{trilha.rstrip(".")}"'))
    return violacoes


_RE_BLOCO_HTML = re.compile(
    r'<script type="application/json" id="(?:dados|rota)-pagina">(.*?)</script>', re.DOTALL,
)


def _checar_html(caminho: Path) -> list[Violacao]:
    """As páginas de site/ embutem os dados direto no HTML, em
    <script type="application/json"> (para abrir local sem servidor, sem
    precisar de fetch) — por isso o HTML também precisa ser varrido, não
    só o JSON de origem. A busca é pela tag inteira (não por chave/valor
    "solto"), porque o fim de um bloco JSON só é inequívoco na fronteira
    real da tag — tentar reconhecer "o JSON acaba aqui" contando chaves
    quebraria se algum texto livre de config contivesse "};" no meio."""
    texto = caminho.read_text(encoding='utf-8')
    violacoes: list[Violacao] = []
    for bloco in _RE_BLOCO_HTML.findall(texto):
        dados = json.loads(bloco.replace('<\\/', '</'))
        violacoes += _checar_json_recursivo(dados, caminho)
    return violacoes


def verificar_arquivo(caminho: Path) -> list[Violacao]:
    if caminho.suffix.lower() == '.csv':
        return _checar_csv(caminho)
    if caminho.suffix.lower() == '.json':
        with open(caminho, encoding='utf-8') as f:
            dados = json.load(f)
        return _checar_json_recursivo(dados, caminho)
    if caminho.suffix.lower() in ('.html', '.htm'):
        return _checar_html(caminho)
    return []


def verificar_pastas(pastas: list[Path]) -> list[Violacao]:
    violacoes: list[Violacao] = []
    for pasta in pastas:
        if not pasta.exists():
            continue
        for caminho in pasta.rglob('*'):
            if caminho.is_file() and caminho.suffix.lower() in ('.csv', '.json', '.html', '.htm'):
                violacoes += verificar_arquivo(caminho)
    return violacoes


if __name__ == '__main__':
    raiz = Path(__file__).resolve().parent.parent
    violacoes = verificar_pastas([raiz / 'site', raiz / 'data'])
    if violacoes:
        print(f'✗ VERIFICAÇÃO DE PRIVACIDADE FALHOU — {len(violacoes)} problema(s):')
        for v in violacoes:
            print(f'  - {v}')
        raise SystemExit(1)
    print('✓ Verificação de privacidade OK — nenhum dado nominal encontrado em site/ ou data/.')
