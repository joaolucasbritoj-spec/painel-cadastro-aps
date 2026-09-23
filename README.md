# Painel de Qualidade do Cadastro — Prima Qualitá Saúde

Painel mensal de acompanhamento da qualidade do Cadastro Individual (Estratégia
Saúde da Família) nos municípios atendidos: Santa Maria Madalena, Saquarema,
Arraial do Cabo e Quissamã. Substitui a planilha `MAD.xlsx` + Power BI por um
pipeline em Python + um site estático, publicado no GitHub Pages/Netlify.

> **Privacidade (LGPD):** o site publicado e o histórico versionado (pasta
> `data/`) contêm **só números agregados por equipe e por microárea** — nunca
> nome, CPF, CNS ou qualquer outro dado nominal. Isso é verificado
> automaticamente a cada carga (`pipeline/privacidade.py`); se a verificação
> falhar, o pipeline avisa e **não** deixa a página nova ser publicada.

---

## 1. Passo a passo mensal (para quem não mexe com código)

Todo mês, depois de exportar o relatório de Cadastro Individual do seu
sistema:

1. **Abra um terminal nesta pasta** (`painel-cadastro-aps`). No Windows, clique
   com o botão direito dentro da pasta e escolha "Abrir no Terminal".

2. **Rode o comando** trocando o caminho do arquivo, o município e o mês
   (sempre no formato `AAAA-MM`):

   ```
   python pipeline/atualizar.py --arquivo "C:\caminho\para\export.xlsx" --municipio "Santa Maria Madalena" --competencia 2026-10
   ```

   Se você tiver os 4 exports do mês numa pasta, organizados em subpastas por
   município (uma pasta chamada exatamente como o município, com o `.xlsx`
   dentro), pode rodar todos de uma vez:

   ```
   python pipeline/atualizar.py --lote "C:\caminho\para\exportacoes-do-mes" --competencia 2026-10
   ```

   > **Por que preciso dizer o município na mão?** O nome do arquivo que o
   > sistema gera não é confiável para isso (às vezes vem o nome da unidade,
   > não do município) — por segurança, o município é sempre informado por
   > você, nunca adivinhado.

3. **Confira o resumo** que aparece no terminal ao final: o score de cada
   equipe, a variação em relação ao mês anterior, e a lista de **Alertas**.
   Alertas não são erro do programa — são coisas para você olhar antes de
   publicar (ex.: uma equipe apareceu no arquivo mas não está cadastrada em
   `config/equipes.yaml`, ou uma condição de saúde apareceu zerada numa
   equipe grande demais para isso ser normal).

4. **As planilhas de correção de cada equipe** já são salvas automaticamente
   na pasta do Google Drive configurada (uma subpasta por equipe e por mês).
   Nada para fazer manualmente aqui.

5. **Publique o site** (veja a seção [Publicação](#4-publicação) abaixo) —
   normalmente é só `git add`, `git commit` e `git push`; quem cuida do
   GitHub Pages/Netlify configurado faz o resto sozinho.

Se o terminal mostrar `✗ VERIFICAÇÃO DE PRIVACIDADE FALHOU`, **não publique**
— chame quem mantém o projeto tecnicamente antes de prosseguir.

---

## 2. Preparando pela primeira vez

1. Instale o [Python 3.11 ou mais recente](https://www.python.org/downloads/).
2. Nesta pasta, instale as dependências:

   ```
   pip install -r requirements.txt
   ```

3. Preencha os arquivos de configuração que ainda estão com campos em
   branco (procure por comentários `PENDENTE` dentro deles):
   - `config/equipes.yaml` — código INE, link da pasta do Drive e link do
     Padlet de cada equipe (e as equipes dos outros 3 municípios, à medida
     que as exportações deles chegarem).
   - `config/caminhos.yaml` — pasta onde ficam as exportações mensais e a
     pasta-base do Google Drive para Desktop onde as planilhas de correção
     são salvas.

   Nenhum desses arquivos entra no cálculo de nada além de organizar
   nomes/links — editar e salvar já é suficiente, não precisa rodar nada
   além do pipeline normal depois.

---

## 3. O que cada arquivo de configuração controla

Nada de regra de negócio fica escondida em código — tudo isso é editável
sem programar, em `config/`:

| Arquivo | O que controla |
|---|---|
| `metas.yaml` | Meta de cada inconsistência (padrão 10%), meta de score (80) e os limites de criticidade. |
| `inconsistencias.yaml` | O catálogo de inconsistências: quais existem, como detectar cada uma, se conta no score, e a orientação de correção mostrada no painel. Para desativar uma, mude `ativo: false` — o histórico já gravado não é apagado. |
| `equipes.yaml` | Município, INE, nome, link do Drive e do Padlet de cada equipe. |
| `grupos.yaml` | Os grupos prioritários (hipertensos, diabéticos, gestantes, crianças, idosos) usados nos indicadores complementares. |
| `condicoes_saude.yaml` | As condições de saúde acompanhadas e o critério de "alerta de plausibilidade" (zero casos numa equipe grande é suspeito). |
| `caminhos.yaml` | Pasta das exportações e pasta-base do Drive. |

Depois de editar qualquer um desses arquivos, rode o pipeline de novo (ou
`python pipeline/gerar_site.py`, se só quiser recompilar o site sem
processar uma exportação nova) para o painel refletir a mudança.

---

## 4. Publicação

O site pronto para publicar fica inteiro dentro da pasta `site/` (gerada
pelo pipeline — não edite os arquivos `.html` dela na mão, eles são
recriados a cada carga).

### GitHub Pages (repositório público ou privado)

Este repositório já vem com `.github/workflows/pages.yml`: a cada
`git push` na branch principal, o GitHub publica automaticamente o conteúdo
de `site/`. Só precisa, uma vez, no GitHub:

1. Criar o repositório e subir este projeto (`git remote add origin ...`,
   `git push -u origin main`).
2. Em **Settings → Pages**, em "Source", escolher **GitHub Actions**.

Repositório privado também funciona (o GitHub permite Pages a partir de
repositório privado em contas pessoais), mas a página publicada em si fica
acessível a quem tiver o link — como o site só tem dado agregado, isso é
seguro por design.

### Netlify (alternativa, útil se preferir manter o repositório fechado)

Já existe um `netlify.toml` configurado com `publish = "site"`. Basta
conectar o repositório em [app.netlify.com](https://app.netlify.com) —
nenhuma configuração de build é necessária (o site já vem pronto, gerado
pelo pipeline antes do commit).

---

## 5. Testes

```
python -m pytest
```

Cobre o cálculo de cada inconsistência, o score legado (fidelidade com os
números que o Power BI já mostrava), o score corrigido e a verificação de
privacidade — todos com dado fictício gerado no próprio teste, nunca com
uma exportação real.

---

## 6. Estrutura do repositório

```
config/       YAMLs de configuração (a única fonte de regra de negócio)
data/         Histórico agregado (CSV, sem dado nominal) — versionado
pipeline/     Scripts Python: migração, carga mensal, geração do site
site/         Site estático gerado (HTML/CSS/JS puro, sem framework)
tests/        Testes automatizados (pytest)
assets/       Logo, brasões dos municípios
gerar_relatorio.py   Script original de relatório (reaproveitado pelo pipeline, não reescrito)
```

`gerar_relatorio.py`, `data/*.csv` e os YAMLs de `config/` são as únicas
partes que valem a pena revisar manualmente entre uma carga e outra — o
resto (`site/`, planilhas de listagem) é sempre gerado de novo.
