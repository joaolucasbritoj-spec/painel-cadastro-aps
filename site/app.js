/* ────────────────────────────────────────────────────────────────────────
   Painel de Qualidade do Cadastro — app.js
   JS puro, sem framework/biblioteca de gráfico. Cada página é gerada como
   um arquivo estático próprio (por pipeline/gerar_site.py), com os dados
   já embutidos em <script type="application/json" id="dados-pagina"> e a
   rota em id="rota-pagina" — não há fetch nem roteador client-side, então
   a página abre normalmente com um duplo-clique local (file://) e também
   funciona em qualquer hospedagem estática (GitHub Pages, Netlify).
   ──────────────────────────────────────────────────────────────────────── */

(function () {
  'use strict';

  // ══════════════════════════════════════════════════════════════════════
  //  FORMATAÇÃO (padrão brasileiro)
  // ══════════════════════════════════════════════════════════════════════

  function formatarNumero(n) {
    return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  function formatarPercentual(valor, total, casas) {
    casas = casas === undefined ? 1 : casas;
    if (!total) return '—';
    var pct = (valor / total) * 100;
    return pct.toFixed(casas).replace('.', ',') + '%';
  }

  function formatarPontoPercentual(valor, casas) {
    casas = casas === undefined ? 2 : casas;
    var sinal = valor > 0 ? '+' : '';
    return sinal + valor.toFixed(casas).replace('.', ',');
  }

  var MESES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

  function formatarCompetencia(aaaaMm) {
    var partes = aaaaMm.split('-');
    var mes = parseInt(partes[1], 10);
    var ano2 = partes[0].slice(2);
    return MESES_ABREV[mes - 1] + '/' + ano2;
  }

  // ══════════════════════════════════════════════════════════════════════
  //  CRITICIDADE (escala única, derivada de metas.criticidade)
  // ══════════════════════════════════════════════════════════════════════

  function classificarCriticidade(score, metas) {
    if (score === null || score === undefined) return null;
    if (score >= metas.criticidade.adequada) return 'adequada';
    if (score >= metas.criticidade.atencao) return 'atencao';
    return 'critica';
  }

  var ROTULO_CRITICIDADE = { adequada: 'Adequada', atencao: 'Atenção', critica: 'Crítica' };

  function elementoSelo(chave, textoExtra) {
    var selo = document.createElement('span');
    if (chave === null) {
      selo.className = 'selo selo--incompleto';
      selo.textContent = textoExtra || 'Sem dado';
      return selo;
    }
    selo.className = 'selo selo--' + chave;
    selo.textContent = textoExtra || ROTULO_CRITICIDADE[chave];
    return selo;
  }

  // ══════════════════════════════════════════════════════════════════════
  //  AGREGAÇÃO — SEMPRE soma dos numeradores / soma dos denominadores,
  //  nunca média de percentuais (regra explícita do projeto). Usada tanto
  //  para agregar equipes de um município quanto municípios da rede.
  // ══════════════════════════════════════════════════════════════════════

  // Extrai, de uma lista de equipes (cada uma com .serie), o item de série
  // de uma competência específica. Equipe sem carga naquela competência
  // simplesmente não entra na soma (ausente != zero).
  function itensDaCompetencia(equipes, competencia) {
    var itens = [];
    for (var i = 0; i < equipes.length; i++) {
      var serie = equipes[i].serie;
      for (var j = 0; j < serie.length; j++) {
        if (serie[j].competencia === competencia) { itens.push(serie[j]); break; }
      }
    }
    return itens;
  }

  // Score agregado (rede ou município): mesma regra de pipeline/score.py,
  // só que somando sobre várias equipes em vez de ler direto do histórico.
  function scoreAgregado(equipes, competencia, indicadoresCfg) {
    var itens = itensDaCompetencia(equipes, competencia);
    if (itens.length === 0) return { score: null, dadoIncompleto: true, totalCadastrados: 0 };

    var idsScore = {};
    indicadoresCfg.forEach(function (i) { if (i.entra_no_score) idsScore[i.id] = true; });

    var totalGeral = 0, somaGeral = 0, algumAusente = false;
    itens.forEach(function (item) {
      totalGeral += item.total_cadastrados;
      item.indicadores.forEach(function (ind) {
        if (!idsScore[ind.id]) return;
        if (ind.ausente) { algumAusente = true; return; }
        somaGeral += ind.valor;
      });
    });

    if (totalGeral === 0) return { score: null, dadoIncompleto: true, totalCadastrados: 0 };
    var score = 100 - (somaGeral / totalGeral) * 100;
    score = Math.max(0, Math.min(100, score));
    return { score: Math.round(score * 100) / 100, dadoIncompleto: algumAusente, totalCadastrados: totalGeral };
  }

  // % Inconsistência Global: MESMO conjunto do score + as inativas
  // (entra_no_score=false), sobre o total de cadastrados.
  function percentualInconsistenciaGlobal(equipes, competencia, indicadoresCfg) {
    var itens = itensDaCompetencia(equipes, competencia);
    if (itens.length === 0) return null;
    var totalGeral = 0, somaGeral = 0;
    var idsValidos = {};
    indicadoresCfg.forEach(function (i) { idsValidos[i.id] = true; });
    itens.forEach(function (item) {
      totalGeral += item.total_cadastrados;
      item.indicadores.forEach(function (ind) {
        if (!idsValidos[ind.id] || ind.ausente) return;
        somaGeral += ind.valor;
      });
    });
    if (totalGeral === 0) return null;
    return (somaGeral / totalGeral) * 100;
  }

  // Detalhe (soma, total, %) de UM indicador, agregado sobre uma lista de
  // equipes — usado nos tooltips, que sempre mostram valor absoluto E percentual.
  function detalheIndicador(equipes, competencia, indicadorId) {
    var itens = itensDaCompetencia(equipes, competencia);
    var total = 0, soma = 0, algumPresente = false;
    itens.forEach(function (item) {
      var ind = item.indicadores.filter(function (i) { return i.id === indicadorId; })[0];
      if (!ind || ind.ausente) return;
      algumPresente = true;
      total += item.total_cadastrados;
      soma += ind.valor;
    });
    if (!algumPresente || total === 0) return null;
    return { soma: soma, total: total, pct: (soma / total) * 100 };
  }

  // Taxa (%) de UM indicador, agregada sobre uma lista de equipes.
  function taxaIndicador(equipes, competencia, indicadorId) {
    var d = detalheIndicador(equipes, competencia, indicadorId);
    return d ? d.pct : null;
  }

  function competenciaAnteriorNaSerie(competencias, competenciaAtual) {
    var idx = competencias.indexOf(competenciaAtual);
    if (idx <= 0) return null;
    return competencias[idx - 1];
  }

  // Alertas de plausibilidade (condições de saúde): sem taxa clínica
  // documentada para a maioria das condições, o critério é "zero casos
  // numa equipe grande o suficiente é estatisticamente improvável" — quase
  // sempre sinal de campo não preenchido, nunca "confirmação de que está
  // tudo bem". A única exceção com taxa mínima de verdade é a que o
  // catálogo trouxer com `taxa_esperada_min_pct` (hoje só gestantes).
  function calcularAlertasPlausibilidade(catalogoCondicoes, itensCondicao, minimoCadastros) {
    var alertas = [];
    (itensCondicao || []).forEach(function (item) {
      var cat = catalogoCondicoes.filter(function (c) { return c.id === item.id; })[0];
      if (!cat || !item.total_cadastrados) return;
      var pct = (item.valor / item.total_cadastrados) * 100;
      if (cat.taxa_esperada_min_pct !== null && cat.taxa_esperada_min_pct !== undefined) {
        if (pct < cat.taxa_esperada_min_pct) {
          alertas.push({
            titulo: cat.titulo,
            mensagem: cat.titulo + ' está em ' + pct.toFixed(2).replace('.', ',') + '%, abaixo do esperado (≥' +
              cat.taxa_esperada_min_pct + '%' + (cat.taxa_esperada_nota ? ' — ' + cat.taxa_esperada_nota : '') + ').',
          });
        }
      } else if (item.valor === 0 && item.total_cadastrados >= minimoCadastros) {
        alertas.push({
          titulo: cat.titulo,
          mensagem: 'Ninguém tem "' + cat.titulo + '" registrado (0 de ' + formatarNumero(item.total_cadastrados) +
            ' cadastros) — confira se esse campo está sendo preenchido no atendimento.',
        });
      }
    });
    return alertas;
  }

  // Posição da equipe dentro do município nesta competência — dense rank
  // (empate divide a mesma posição, sem pular número), 1º = MELHOR score
  // (regra 8 do painel novo — no legado era o contrário).
  function posicaoRanking(equipes, competencia) {
    var comScore = equipes.map(function (eq) {
      var it = eq.serie.filter(function (s) { return s.competencia === competencia; })[0];
      return it && it.score !== null ? { slug: eq.slug, score: it.score } : null;
    }).filter(Boolean);
    var ordenados = comScore.slice().sort(function (a, b) { return b.score - a.score; });
    var posicoes = {}, rank = 0, ultimoScore = null;
    ordenados.forEach(function (item) {
      if (item.score !== ultimoScore) { rank += 1; ultimoScore = item.score; }
      posicoes[item.slug] = rank;
    });
    return { posicoes: posicoes, total: ordenados.length };
  }

  // Regressão linear simples (mínimos quadrados) sobre até os últimos 3
  // pontos DISPONÍVEIS (ausente não conta como ponto) — com menos de 2,
  // não dá para estimar tendência. `pontos` deve trazer o índice ABSOLUTO
  // de cada um (posição real no eixo x do gráfico, não posição dentro do
  // recorte dos últimos 3), para a projeção depois saber onde continuar
  // a linha mesmo com um histórico mais longo que 3 competências.
  function regressaoLinear(pontos) {
    var pts = pontos.filter(function (p) { return p.valor !== null; }).slice(-3);
    if (pts.length < 2) return null;
    var n = pts.length;
    var somaX = 0, somaY = 0, somaXY = 0, somaXX = 0;
    pts.forEach(function (p) {
      somaX += p.i; somaY += p.valor; somaXY += p.i * p.valor; somaXX += p.i * p.i;
    });
    var denom = n * somaXX - somaX * somaX;
    if (denom === 0) return null;
    var inclinacao = (n * somaXY - somaX * somaY) / denom;
    var intercepto = (somaY - inclinacao * somaX) / n;
    var ultimo = pts[n - 1];
    return { inclinacao: inclinacao, intercepto: intercepto, ultimoIndice: ultimo.i, ultimoValor: ultimo.valor };
  }

  // Projeta quando a taxa de um indicador cruzaria a meta, "no ritmo
  // atual" (extrapolando a reta da regressão). Não tenta plotar esse
  // ponto no gráfico se for muito no futuro — só devolve o texto e um
  // trechinho curto de linha tracejada para ilustrar a tendência.
  function projetarMeta(pontosSerie, metaPct, competenciaAtual) {
    if (pontosSerie.filter(function (p) { return p.valor !== null; }).length < 2) {
      return { status: 'dados_insuficientes' };
    }
    var ultimoValor = pontosSerie.filter(function (p) { return p.valor !== null; }).slice(-1)[0].valor;
    if (ultimoValor <= metaPct) return { status: 'meta_atingida' };

    var reg = regressaoLinear(pontosSerie);
    if (!reg) return { status: 'dados_insuficientes' };
    if (reg.inclinacao >= 0) return { status: 'sem_queda' };

    var passosParaMeta = (metaPct - reg.ultimoValor) / reg.inclinacao;
    var mesesParaMeta = Math.ceil(passosParaMeta);
    var partes = competenciaAtual.split('-').map(Number);
    var totalMeses = partes[0] * 12 + (partes[1] - 1) + mesesParaMeta;
    var anoMeta = Math.floor(totalMeses / 12);
    var mesMeta = (totalMeses % 12) + 1;
    var competenciaMeta = anoMeta + '-' + (mesMeta < 10 ? '0' : '') + mesMeta;

    // trechinho curto de projeção (2 passos), só para ilustrar a direção —
    // não tenta alcançar visualmente a meta se ela estiver muito longe.
    var pontoFinal = { i: reg.ultimoIndice + 2, v: Math.max(0, reg.ultimoValor + reg.inclinacao * 2) };
    var pontoInicial = { i: reg.ultimoIndice, v: reg.ultimoValor };

    return {
      status: 'projetado',
      competenciaMeta: competenciaMeta,
      mesesParaMeta: mesesParaMeta,
      segmentoProjecao: [pontoInicial, pontoFinal],
    };
  }

  // ══════════════════════════════════════════════════════════════════════
  //  NAVEGAÇÃO — caminho relativo entre páginas estáticas
  // ══════════════════════════════════════════════════════════════════════

  var PROFUNDIDADE = { rede: 0, municipio: 1, equipe: 2, ranking: 1, metodologia: 1 };

  function linkPara(destino, rotaAtual) {
    var subir = '../'.repeat(PROFUNDIDADE[rotaAtual.tipo] || 0);
    if (destino.tipo === 'rede') return subir || './';
    if (destino.tipo === 'municipio') return subir + destino.municipio + '/';
    if (destino.tipo === 'equipe') return subir + destino.municipio + '/' + destino.equipe + '/';
    if (destino.tipo === 'ranking') return subir + 'ranking/';
    if (destino.tipo === 'metodologia') return subir + 'metodologia/';
    return subir || './';
  }

  // ══════════════════════════════════════════════════════════════════════
  //  SVG — helpers genéricos para desenhar os gráficos à mão
  // ══════════════════════════════════════════════════════════════════════

  var SVG_NS = 'http://www.w3.org/2000/svg';

  function svgEl(tag, attrs) {
    var el = document.createElementNS(SVG_NS, tag);
    for (var chave in attrs) {
      if (attrs[chave] !== undefined && attrs[chave] !== null) el.setAttribute(chave, attrs[chave]);
    }
    return el;
  }

  var tooltipEl = null;
  function mostrarTooltip(evento, html) {
    if (!tooltipEl) {
      tooltipEl = document.createElement('div');
      tooltipEl.className = 'tooltip-grafico';
      document.body.appendChild(tooltipEl);
    }
    tooltipEl.innerHTML = html;
    tooltipEl.style.left = evento.clientX + 'px';
    tooltipEl.style.top = (evento.clientY - 12) + 'px';
    tooltipEl.hidden = false;
  }
  function esconderTooltip() { if (tooltipEl) tooltipEl.hidden = true; }

  function corPorTaxaVsMeta(taxaPct, metaPct) {
    if (taxaPct <= metaPct) return 'adequada';
    if (taxaPct <= metaPct * 2) return 'atencao';
    return 'critica';
  }

  // ── Gráfico de evolução (linha) ─────────────────────────────────────
  function graficoEvolucao(opcoes) {
    var largura = 640, altura = 200, margem = { topo: 16, baixo: 26, esq: 34, dir: 16 };
    var w = largura - margem.esq - margem.dir, h = altura - margem.topo - margem.baixo;

    var svg = svgEl('svg', { class: 'grafico', viewBox: '0 0 ' + largura + ' ' + altura, role: 'img', 'aria-label': opcoes.aria });
    var g = svgEl('g', { transform: 'translate(' + margem.esq + ',' + margem.topo + ')' });
    svg.appendChild(g);

    var valores = opcoes.pontos.map(function (p) { return p.valor; }).filter(function (v) { return v !== null; });
    var maxValor = Math.max(opcoes.meta || 0, valores.length ? Math.max.apply(null, valores) : 0) * 1.15 || 100;
    var minValor = 0;
    var n = opcoes.pontos.length;
    var passoX = n > 1 ? w / (n - 1) : 0;

    function x(i) { return passoX * i; }
    function y(v) { return h - ((v - minValor) / (maxValor - minValor || 1)) * h; }

    // eixo base
    g.appendChild(svgEl('line', { class: 'eixo', x1: 0, y1: h, x2: w, y2: h }));

    // linha de meta
    if (opcoes.meta !== undefined && opcoes.meta !== null) {
      var ym = y(opcoes.meta);
      g.appendChild(svgEl('line', { class: 'meta', x1: 0, y1: ym, x2: w, y2: ym }));
      var rotuloMeta = svgEl('text', { class: 'meta-rotulo', x: w, y: ym - 4, 'text-anchor': 'end' });
      rotuloMeta.textContent = 'meta ' + opcoes.meta;
      g.appendChild(rotuloMeta);
    }

    // linha da série (só entre pontos consecutivos não-nulos)
    var pontosValidos = [];
    opcoes.pontos.forEach(function (p, i) { if (p.valor !== null) pontosValidos.push({ i: i, v: p.valor }); });
    var d = '';
    pontosValidos.forEach(function (p, k) {
      d += (k === 0 ? 'M' : 'L') + x(p.i).toFixed(1) + ',' + y(p.v).toFixed(1) + ' ';
    });
    if (d) g.appendChild(svgEl('path', { class: 'linha-serie stroke-adequada', d: d }));

    // projeção tracejada (opcional)
    if (opcoes.projecao && opcoes.projecao.length > 1) {
      var dp = '';
      opcoes.projecao.forEach(function (p, k) {
        dp += (k === 0 ? 'M' : 'L') + x(p.i).toFixed(1) + ',' + y(p.v).toFixed(1) + ' ';
      });
      g.appendChild(svgEl('path', { class: 'linha-projecao stroke-atencao', d: dp }));
    }

    // pontos + rótulo direto + eixo x
    opcoes.pontos.forEach(function (p, i) {
      var rotuloX = svgEl('text', { class: 'rotulo-eixo', x: x(i), y: h + 18, 'text-anchor': 'middle' });
      rotuloX.textContent = p.rotulo;
      g.appendChild(rotuloX);

      if (p.valor === null) return;
      var cor = opcoes.corPonto ? opcoes.corPonto(p.valor) : (classificarCriticidade(p.valor, opcoes.metas) || 'adequada');
      var ponto = svgEl('circle', { class: 'ponto fill-' + cor, cx: x(i), cy: y(p.valor), r: 4 });
      ponto.addEventListener('mousemove', function (ev) {
        mostrarTooltip(ev, '<strong>' + p.rotulo + '</strong><br>' + opcoes.formatarValor(p.valor) + (p.extra || ''));
      });
      ponto.addEventListener('mouseleave', esconderTooltip);
      g.appendChild(ponto);

      var rotuloValor = svgEl('text', { class: 'rotulo-direto', x: x(i), y: y(p.valor) - 10, 'text-anchor': 'middle' });
      rotuloValor.textContent = opcoes.formatarValor(p.valor);
      g.appendChild(rotuloValor);
    });

    return svg;
  }

  // ── Gráfico de barras com rótulo direto (distribuição por indicador) ──
  function graficoBarrasDiretas(opcoes) {
    var largura = 640;
    var alturaLinha = 30;
    var margem = { esq: 190, dir: 60, topo: 4, baixo: 4 };
    var w = largura - margem.esq - margem.dir;
    var altura = opcoes.itens.length * alturaLinha + margem.topo + margem.baixo;

    var svg = svgEl('svg', { class: 'grafico', viewBox: '0 0 ' + largura + ' ' + altura, role: 'img', 'aria-label': opcoes.aria });
    var maxValor = Math.max(opcoes.meta, opcoes.itens.reduce(function (m, it) { return Math.max(m, it.valor || 0); }, 0)) * 1.2 || 10;

    function x(v) { return (v / maxValor) * w; }

    opcoes.itens.forEach(function (item, i) {
      var cy = margem.topo + i * alturaLinha;
      var rotulo = svgEl('text', {
        x: margem.esq - 10, y: cy + alturaLinha / 2 + 4, 'text-anchor': 'end', class: 'rotulo-eixo',
      });
      rotulo.textContent = item.rotulo + (item.ausente ? ' (dado incompleto)' : '');
      svg.appendChild(rotulo);

      if (item.ausente || item.valor === null) {
        var tracoAusente = svgEl('line', {
          x1: margem.esq, y1: cy + alturaLinha / 2, x2: margem.esq + 14, y2: cy + alturaLinha / 2,
          class: 'eixo',
        });
        svg.appendChild(tracoAusente);
        return;
      }

      var cor = corPorTaxaVsMeta(item.valor, opcoes.meta);
      var barra = svgEl('rect', {
        x: margem.esq, y: cy + 5, width: Math.max(2, x(item.valor)), height: alturaLinha - 12,
        rx: 4, class: 'fill-' + cor + ' ponto',
      });
      barra.addEventListener('mousemove', function (ev) {
        mostrarTooltip(ev, '<strong>' + item.rotulo + '</strong><br>' + formatarNumero(item.n) + ' de ' + formatarNumero(item.total) + ' (' + item.valor.toFixed(1).replace('.', ',') + '%)');
      });
      barra.addEventListener('mouseleave', esconderTooltip);
      svg.appendChild(barra);

      var valorTxt = svgEl('text', {
        x: margem.esq + x(item.valor) + 8, y: cy + alturaLinha / 2 + 4, class: 'rotulo-direto',
      });
      valorTxt.textContent = item.valor.toFixed(1).replace('.', ',') + '%';
      svg.appendChild(valorTxt);
    });

    // linha de meta vertical
    var xm = margem.esq + x(opcoes.meta);
    svg.appendChild(svgEl('line', { class: 'meta', x1: xm, y1: 0, x2: xm, y2: altura }));
    var rotuloMeta = svgEl('text', { class: 'meta-rotulo', x: xm + 4, y: 10 });
    rotuloMeta.textContent = 'meta ' + opcoes.meta + '%';
    svg.appendChild(rotuloMeta);

    return svg;
  }

  // ── Mapa de calor (linhas × colunas genérico) ────────────────────────
  // Rótulo de coluna fica horizontal (sem rotação) — por isso quem chama
  // deve preferir colocar a dimensão de rótulo CURTO nas colunas (equipes)
  // e a de rótulo longo nas linhas (indicadores), não o contrário: texto
  // rotacionado dentro de <svg> é frágil entre navegadores (o corte/
  // sobreposição observado numa versão anterior deste gráfico vinha
  // exatamente disso).
  function graficoMapaCalor(opcoes) {
    var maiorTitulo = opcoes.colunas.reduce(function (m, c) { return Math.max(m, c.titulo.length); }, 4);
    var margem = { esq: 150, topo: 34 };
    var celula = Math.max(60, maiorTitulo * 6.5 + 16), celulaAltura = 30;
    var w = opcoes.colunas.length * celula, h = opcoes.linhas.length * celulaAltura;
    var largura = margem.esq + w + 10, altura = margem.topo + h + 6;

    var svg = svgEl('svg', { class: 'grafico', viewBox: '0 0 ' + largura + ' ' + altura, role: 'img', 'aria-label': opcoes.aria });

    opcoes.colunas.forEach(function (col, c) {
      var t = svgEl('text', {
        x: margem.esq + c * celula + celula / 2, y: margem.topo - 10, 'text-anchor': 'middle', class: 'rotulo-eixo',
      });
      t.textContent = col.titulo;
      svg.appendChild(t);
    });

    opcoes.linhas.forEach(function (linha, r) {
      var t = svgEl('text', { x: margem.esq - 10, y: margem.topo + r * celulaAltura + celulaAltura / 2 + 4, 'text-anchor': 'end', class: 'rotulo-eixo' });
      t.textContent = linha.titulo;
      svg.appendChild(t);

      opcoes.colunas.forEach(function (col, c) {
        var valor = opcoes.obterValor(linha, col);
        var meta = opcoes.obterMeta(linha, col);
        var gx = margem.esq + c * celula, gy = margem.topo + r * celulaAltura;
        var cor = valor === null ? 'neutro' : corPorTaxaVsMeta(valor, meta);
        var rect = svgEl('rect', {
          x: gx, y: gy, width: celula, height: celulaAltura, class: 'celula-mapa-calor fill-' + cor,
        });
        svg.appendChild(rect);
        if (valor !== null) {
          var rotulo = svgEl('text', {
            x: gx + celula / 2, y: gy + celulaAltura / 2, class: 'celula-mapa-calor__rotulo',
          });
          rotulo.textContent = valor.toFixed(0) + '%';
          rotulo.style.fill = '#fff';
          svg.appendChild(rotulo);
        }
        rect.addEventListener('mousemove', function (ev) {
          var txt = valor === null ? 'sem dado' : valor.toFixed(1).replace('.', ',') + '% (meta ' + meta + '%)';
          mostrarTooltip(ev, '<strong>' + linha.titulo + '</strong><br>' + col.titulo + ': ' + txt);
        });
        rect.addEventListener('mouseleave', esconderTooltip);
      });
    });

    return svg;
  }

  // ── Curva de Pareto (concentração de inconsistências por microárea) ──
  function graficoPareto(opcoes) {
    var itens = opcoes.itens.slice().sort(function (a, b) { return b.valor - a.valor; });
    var totalGeral = itens.reduce(function (s, it) { return s + it.valor; }, 0) || 1;

    var largura = Math.max(480, itens.length * 64);
    var altura = 220, margem = { topo: 20, baixo: 60, esq: 40, dir: 40 };
    var w = largura - margem.esq - margem.dir, h = altura - margem.topo - margem.baixo;
    var svg = svgEl('svg', { class: 'grafico', viewBox: '0 0 ' + largura + ' ' + altura, role: 'img', 'aria-label': opcoes.aria });
    var g = svgEl('g', { transform: 'translate(' + margem.esq + ',' + margem.topo + ')' });
    svg.appendChild(g);

    var maxValor = Math.max.apply(null, itens.map(function (it) { return it.valor; })) * 1.15 || 1;
    var passo = itens.length > 1 ? w / itens.length : w;
    var largBarra = Math.min(48, passo * 0.6);

    g.appendChild(svgEl('line', { class: 'eixo', x1: 0, y1: h, x2: w, y2: h }));

    // linha de referência 80% acumulado
    var y80 = h - 0.8 * h;
    g.appendChild(svgEl('line', { class: 'meta', x1: 0, y1: y80, x2: w, y2: y80 }));
    var rot80 = svgEl('text', { class: 'meta-rotulo', x: w, y: y80 - 4, 'text-anchor': 'end' });
    rot80.textContent = '80% acumulado';
    g.appendChild(rot80);

    var acumulado = 0;
    var pontosLinha = [];
    itens.forEach(function (item, i) {
      acumulado += item.valor;
      var pctAcumulado = acumulado / totalGeral;
      var cx = passo * i + passo / 2;

      var barra = svgEl('rect', {
        x: cx - largBarra / 2, y: h - (item.valor / maxValor) * h, width: largBarra, height: (item.valor / maxValor) * h,
        class: 'fill-atencao ponto', rx: 3,
      });
      barra.addEventListener('mousemove', function (ev) {
        mostrarTooltip(ev, '<strong>' + item.rotulo + '</strong><br>' + formatarNumero(item.valor) + ' inconsistências<br>' + (pctAcumulado * 100).toFixed(0) + '% acumulado');
      });
      barra.addEventListener('mouseleave', esconderTooltip);
      g.appendChild(barra);

      var valorTxt = svgEl('text', { class: 'rotulo-direto', x: cx, y: h - (item.valor / maxValor) * h - 6, 'text-anchor': 'middle' });
      valorTxt.textContent = formatarNumero(item.valor);
      g.appendChild(valorTxt);

      var rotuloX = svgEl('text', {
        class: 'rotulo-eixo', x: cx, y: h + 16, 'text-anchor': 'end',
        transform: 'rotate(-40 ' + cx + ' ' + (h + 16) + ')',
      });
      rotuloX.textContent = item.rotulo;
      g.appendChild(rotuloX);

      pontosLinha.push({ x: cx, y: h - pctAcumulado * h });
    });

    var d = pontosLinha.map(function (p, i) { return (i === 0 ? 'M' : 'L') + p.x.toFixed(1) + ',' + p.y.toFixed(1); }).join(' ');
    g.appendChild(svgEl('path', { class: 'linha-serie stroke-critica', d: d }));
    pontosLinha.forEach(function (p) { g.appendChild(svgEl('circle', { class: 'fill-critica', cx: p.x, cy: p.y, r: 3 })); });

    return svg;
  }

  // ── Waffle 10×10 (cadastros limpos) ──────────────────────────────────
  function construirWaffle(pctLimpos) {
    var preenchidos = Math.round(pctLimpos);
    var grade = document.createElement('div');
    grade.className = 'waffle-grade';
    grade.setAttribute('role', 'img');
    grade.setAttribute('aria-label', pctLimpos.toFixed(1).replace('.', ',') + '% dos cadastros sem nenhuma inconsistência ativa');
    for (var i = 0; i < 100; i++) {
      var quad = document.createElement('div');
      quad.className = 'waffle-quad ' + (i < preenchidos ? 'bg-adequada' : 'bg-neutro');
      grade.appendChild(quad);
    }
    return grade;
  }

  // ── Pirâmide etária espelhada (com CPF × sem CPF) ────────────────────
  function graficoPiramide(faixas, aria) {
    var largura = 640, alturaLinha = 32;
    var margem = { esq: 60, dir: 60, topo: 10, baixo: 22, centro: 40 };
    var w = (largura - margem.esq - margem.dir - margem.centro) / 2;
    var altura = faixas.length * alturaLinha + margem.topo + margem.baixo;

    var svg = svgEl('svg', { class: 'grafico', viewBox: '0 0 ' + largura + ' ' + altura, role: 'img', 'aria-label': aria });
    var meioX = margem.esq + w + margem.centro / 2;
    var maxValor = Math.max.apply(null, faixas.map(function (f) { return Math.max(f.comCpf, f.semCpf); })) || 1;

    faixas.forEach(function (f, i) {
      var cy = margem.topo + i * alturaLinha;
      var largComCpf = (f.comCpf / maxValor) * w;
      var largSemCpf = (f.semCpf / maxValor) * w;

      var barraComCpf = svgEl('rect', {
        x: meioX - margem.centro / 2 - largComCpf, y: cy + 4, width: largComCpf, height: alturaLinha - 10,
        class: 'fill-adequada ponto', rx: 3,
      });
      barraComCpf.addEventListener('mousemove', function (ev) {
        mostrarTooltip(ev, '<strong>' + f.faixa + ' anos</strong><br>Com CPF: ' + formatarNumero(f.comCpf));
      });
      barraComCpf.addEventListener('mouseleave', esconderTooltip);
      svg.appendChild(barraComCpf);

      var barraSemCpf = svgEl('rect', {
        x: meioX + margem.centro / 2, y: cy + 4, width: largSemCpf, height: alturaLinha - 10,
        class: 'fill-critica ponto', rx: 3,
      });
      barraSemCpf.addEventListener('mousemove', function (ev) {
        mostrarTooltip(ev, '<strong>' + f.faixa + ' anos</strong><br>Sem CPF: ' + formatarNumero(f.semCpf));
      });
      barraSemCpf.addEventListener('mouseleave', esconderTooltip);
      svg.appendChild(barraSemCpf);

      var rotuloFaixa = svgEl('text', { x: meioX, y: cy + alturaLinha / 2 + 4, 'text-anchor': 'middle', class: 'rotulo-eixo' });
      rotuloFaixa.textContent = f.faixa;
      svg.appendChild(rotuloFaixa);

      var valComCpf = svgEl('text', {
        x: meioX - margem.centro / 2 - largComCpf - 6, y: cy + alturaLinha / 2 + 4, 'text-anchor': 'end', class: 'rotulo-direto',
      });
      valComCpf.textContent = formatarNumero(f.comCpf);
      svg.appendChild(valComCpf);

      var valSemCpf = svgEl('text', {
        x: meioX + margem.centro / 2 + largSemCpf + 6, y: cy + alturaLinha / 2 + 4, class: 'rotulo-direto',
      });
      valSemCpf.textContent = formatarNumero(f.semCpf);
      svg.appendChild(valSemCpf);
    });

    var legenda = svgEl('text', { x: margem.esq, y: altura - 2, class: 'rotulo-eixo' });
    legenda.textContent = '← Com CPF';
    svg.appendChild(legenda);
    var legenda2 = svgEl('text', { x: largura - margem.dir, y: altura - 2, 'text-anchor': 'end', class: 'rotulo-eixo' });
    legenda2.textContent = 'Sem CPF →';
    svg.appendChild(legenda2);

    return svg;
  }

  // ══════════════════════════════════════════════════════════════════════
  //  CROMO COMPARTILHADO (cabeçalho, tema, seletor de competência)
  // ══════════════════════════════════════════════════════════════════════

  function aplicarTemaSalvo() {
    try {
      var salvo = localStorage.getItem('tema');
      if (salvo) document.documentElement.setAttribute('data-theme', salvo);
    } catch (e) { /* localStorage pode falhar em aba privada — ignora */ }
  }

  function alternarTema() {
    var atual = document.documentElement.getAttribute('data-theme');
    var prefereEscuro = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    var novo;
    if (!atual) novo = prefereEscuro ? 'light' : 'dark';
    else novo = atual === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', novo);
    try { localStorage.setItem('tema', novo); } catch (e) { /* ignora */ }
  }

  function montarCabecalho(dados, rota) {
    var cab = document.createElement('header');
    cab.className = 'cabecalho';

    var marca = document.createElement('a');
    marca.className = 'cabecalho__marca';
    marca.href = linkPara({ tipo: 'rede' }, rota);
    marca.innerHTML = '<img src="' + (rota.prefixo || '') + 'assets/logo_prima.png" alt=""> Painel de Cadastro';
    cab.appendChild(marca);

    var nav = document.createElement('nav');
    nav.className = 'cabecalho__nav';
    [{ tipo: 'ranking', rotulo: 'Ranking' }, { tipo: 'metodologia', rotulo: 'Metodologia' }].forEach(function (item) {
      var a = document.createElement('a');
      a.href = linkPara({ tipo: item.tipo }, rota);
      a.textContent = item.rotulo;
      if (rota.tipo === item.tipo) a.setAttribute('aria-current', 'page');
      nav.appendChild(a);
    });
    cab.appendChild(nav);

    var controles = document.createElement('div');
    controles.className = 'cabecalho__controles';

    var seletorMunicipio = document.createElement('select');
    seletorMunicipio.className = 'seletor';
    seletorMunicipio.setAttribute('aria-label', 'Selecionar município');
    var optRede = document.createElement('option');
    optRede.value = linkPara({ tipo: 'rede' }, rota);
    optRede.textContent = 'Rede (todos os municípios)';
    seletorMunicipio.appendChild(optRede);
    dados.municipios.forEach(function (m) {
      var opt = document.createElement('option');
      opt.value = linkPara({ tipo: 'municipio', municipio: m.slug }, rota);
      opt.textContent = m.municipio;
      seletorMunicipio.appendChild(opt);
    });
    var valorAtual = rota.tipo === 'municipio' || rota.tipo === 'equipe'
      ? linkPara({ tipo: 'municipio', municipio: rota.municipio }, rota) : linkPara({ tipo: 'rede' }, rota);
    seletorMunicipio.value = valorAtual;
    seletorMunicipio.addEventListener('change', function () { window.location.href = seletorMunicipio.value; });
    controles.appendChild(seletorMunicipio);

    var botaoTema = document.createElement('button');
    botaoTema.className = 'botao-tema';
    botaoTema.type = 'button';
    botaoTema.title = 'Alternar modo claro/escuro';
    botaoTema.textContent = '◐';
    botaoTema.addEventListener('click', alternarTema);
    controles.appendChild(botaoTema);

    cab.appendChild(controles);
    return cab;
  }

  function elementoAjuda(texto) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = 'ajuda';
    b.textContent = '?';
    b.setAttribute('aria-label', texto);
    b.addEventListener('click', function () {
      var aberto = b.hasAttribute('data-aberto');
      document.querySelectorAll('.ajuda[data-aberto]').forEach(function (o) { o.removeAttribute('data-aberto'); });
      if (!aberto) b.setAttribute('data-aberto', '');
    });
    return b;
  }
  if (typeof document !== 'undefined') {
    document.addEventListener('click', function (ev) {
      if (!ev.target.closest || !ev.target.closest('.ajuda')) {
        document.querySelectorAll('.ajuda[data-aberto]').forEach(function (o) { o.removeAttribute('data-aberto'); });
      }
    });
  }

  function cartao(tituloTexto, textoAjuda) {
    var div = document.createElement('div');
    div.className = 'cartao anima-entrada';
    var cab = document.createElement('div');
    cab.className = 'cartao__cabecalho';
    var titulo = document.createElement('div');
    titulo.className = 'cartao__titulo';
    titulo.textContent = tituloTexto;
    if (textoAjuda) titulo.appendChild(elementoAjuda(textoAjuda));
    cab.appendChild(titulo);
    div.appendChild(cab);
    return div;
  }

  // ══════════════════════════════════════════════════════════════════════
  //  PÁGINA: VISÃO GERAL (rede ou município)
  // ══════════════════════════════════════════════════════════════════════

  function renderizarVisaoGeral(container, dados, rota) {
    var equipesEscopo, tituloEscopo, brasao;
    if (rota.tipo === 'municipio') {
      var municipio = dados.municipios.filter(function (m) { return m.slug === rota.municipio; })[0];
      equipesEscopo = municipio.equipes;
      tituloEscopo = municipio.municipio;
      brasao = municipio.brasao;
    } else {
      equipesEscopo = [];
      dados.municipios.forEach(function (m) { equipesEscopo = equipesEscopo.concat(m.equipes); });
      tituloEscopo = 'Visão geral da rede';
    }

    var competencias = [];
    equipesEscopo.forEach(function (eq) {
      eq.serie.forEach(function (s) { if (competencias.indexOf(s.competencia) === -1) competencias.push(s.competencia); });
    });
    competencias.sort();

    var h1 = document.createElement('h1');
    h1.className = 'titulo-pagina';
    if (brasao) h1.innerHTML = '<img class="brasao" src="' + (rota.prefixo || '') + brasao + '" alt="">';
    h1.appendChild(document.createTextNode(tituloEscopo));
    container.appendChild(h1);

    // ── seletor de competência (padrão: mais recente) ──────────────────
    var linhaFiltros = document.createElement('div');
    linhaFiltros.style.marginBottom = '14px';
    var rotuloComp = document.createElement('label');
    rotuloComp.style.fontSize = '13px';
    rotuloComp.style.color = 'var(--texto-suave)';
    rotuloComp.textContent = 'Competência: ';
    var seletorComp = document.createElement('select');
    seletorComp.className = 'seletor';
    competencias.forEach(function (c) {
      var opt = document.createElement('option');
      opt.value = c;
      opt.textContent = formatarCompetencia(c);
      seletorComp.appendChild(opt);
    });
    seletorComp.value = competencias[competencias.length - 1];
    rotuloComp.appendChild(seletorComp);
    linhaFiltros.appendChild(rotuloComp);
    container.appendChild(linhaFiltros);

    var areaConteudo = document.createElement('div');
    container.appendChild(areaConteudo);

    function desenhar(competenciaAtual) {
      areaConteudo.innerHTML = '';
      var competenciaAnterior = competenciaAnteriorNaSerie(competencias, competenciaAtual);

      var atual = scoreAgregado(equipesEscopo, competenciaAtual, dados.indicadores);
      var anterior = competenciaAnterior ? scoreAgregado(equipesEscopo, competenciaAnterior, dados.indicadores) : null;
      var pctGlobal = percentualInconsistenciaGlobal(equipesEscopo, competenciaAtual, dados.indicadores);

      var grade = document.createElement('div');
      grade.className = 'grade';

      // Card: Score
      var cScore = cartao('Score de Qualidade', 'Score = 100 − (Σ das inconsistências que entram no score ÷ total de cadastrados × 100). Quanto maior, melhor.');
      var vScore = document.createElement('div');
      vScore.className = 'stat-valor';
      vScore.textContent = atual.score === null ? '—' : formatarNumero(atual.score) === atual.score.toFixed(0) ? atual.score.toFixed(2).replace('.', ',') : atual.score.toFixed(2).replace('.', ',');
      cScore.appendChild(vScore);
      var linhaSelos = document.createElement('div');
      linhaSelos.style.marginTop = '8px';
      linhaSelos.appendChild(elementoSelo(classificarCriticidade(atual.score, dados.metas)));
      if (atual.dadoIncompleto) {
        var s2 = elementoSelo(null, 'Dado incompleto');
        s2.style.marginLeft = '6px';
        linhaSelos.appendChild(s2);
      }
      cScore.appendChild(linhaSelos);
      if (anterior && anterior.score !== null && atual.score !== null) {
        var delta = atual.score - anterior.score;
        var variacao = document.createElement('div');
        variacao.className = 'stat-variacao ' + (delta >= 0 ? 'stat-variacao--boa' : 'stat-variacao--ruim');
        variacao.textContent = formatarPontoPercentual(delta) + ' pp vs ' + formatarCompetencia(competenciaAnterior);
        cScore.appendChild(variacao);
      }
      grade.appendChild(cScore);

      // Card: % Inconsistência Global
      var cGlobal = cartao('% Inconsistência Global', 'Soma das inconsistências ativas (incluindo as que não entram no score) dividida pelo total de cadastrados.');
      var vGlobal = document.createElement('div');
      vGlobal.className = 'stat-valor';
      vGlobal.textContent = pctGlobal === null ? '—' : pctGlobal.toFixed(1).replace('.', ',') + '%';
      cGlobal.appendChild(vGlobal);
      grade.appendChild(cGlobal);

      // Card: equipes dentro da meta
      var itensAtuais = itensDaCompetencia(equipesEscopo, competenciaAtual);
      var equipesComScore = equipesEscopo.map(function (eq) {
        var it = eq.serie.filter(function (s) { return s.competencia === competenciaAtual; })[0];
        return it ? { equipe: eq, score: it.score } : null;
      }).filter(Boolean);
      var dentroMeta = equipesComScore.filter(function (e) { return e.score !== null && e.score >= dados.metas.meta_score; }).length;
      var cMeta = cartao('Equipes dentro da meta', 'Equipes com Score de Qualidade ≥ ' + dados.metas.meta_score + '.');
      var vMeta = document.createElement('div');
      vMeta.className = 'stat-valor';
      vMeta.textContent = dentroMeta + ' de ' + equipesComScore.length;
      cMeta.appendChild(vMeta);
      grade.appendChild(cMeta);

      // Cards: maior melhora / maior piora (automático, mais recente vs anterior)
      if (competenciaAnterior) {
        var variacoes = [];
        equipesEscopo.forEach(function (eq) {
          var itA = eq.serie.filter(function (s) { return s.competencia === competenciaAtual; })[0];
          var itP = eq.serie.filter(function (s) { return s.competencia === competenciaAnterior; })[0];
          if (itA && itP && itA.score !== null && itP.score !== null) {
            variacoes.push({ equipe: eq, delta: itA.score - itP.score });
          }
        });
        if (variacoes.length) {
          variacoes.sort(function (a, b) { return b.delta - a.delta; });
          var melhora = variacoes[0], piora = variacoes[variacoes.length - 1];

          var cMelhora = cartao('Maior melhora', 'Maior variação positiva de score em relação ao mês anterior.');
          var vMelhora = document.createElement('div');
          vMelhora.className = 'stat-valor';
          vMelhora.style.fontSize = '18px';
          vMelhora.textContent = melhora.equipe.equipe;
          cMelhora.appendChild(vMelhora);
          var legMelhora = document.createElement('div');
          legMelhora.className = 'stat-variacao stat-variacao--boa';
          legMelhora.textContent = formatarPontoPercentual(melhora.delta) + ' pp';
          cMelhora.appendChild(legMelhora);
          grade.appendChild(cMelhora);

          if (piora !== melhora) {
            var cPiora = cartao('Maior piora', 'Maior variação negativa de score em relação ao mês anterior.');
            var vPiora = document.createElement('div');
            vPiora.className = 'stat-valor';
            vPiora.style.fontSize = '18px';
            vPiora.textContent = piora.equipe.equipe;
            cPiora.appendChild(vPiora);
            var legPiora = document.createElement('div');
            legPiora.className = 'stat-variacao ' + (piora.delta >= 0 ? 'stat-variacao--boa' : 'stat-variacao--ruim');
            legPiora.textContent = formatarPontoPercentual(piora.delta) + ' pp';
            cPiora.appendChild(legPiora);
            grade.appendChild(cPiora);
          }
        }
      }

      areaConteudo.appendChild(grade);

      // ── Lista de equipes (só faz sentido na visão do município) ──────
      if (rota.tipo === 'municipio') {
        var cEquipes = cartao('Equipes');
        cEquipes.className += ' cartao--largo';
        var tabela = document.createElement('table');
        tabela.className = 'tabela';
        var corpo = document.createElement('tbody');
        var rankAtual = posicaoRanking(equipesEscopo, competenciaAtual);
        equipesEscopo.slice().sort(function (a, b) {
          return (rankAtual.posicoes[a.slug] || 999) - (rankAtual.posicoes[b.slug] || 999);
        }).forEach(function (eq) {
          var it = eq.serie.filter(function (s) { return s.competencia === competenciaAtual; })[0];
          var tr = document.createElement('tr');
          var tdNome = document.createElement('td');
          var a = document.createElement('a');
          a.className = 'linha-equipe';
          a.href = linkPara({ tipo: 'equipe', municipio: rota.municipio, equipe: eq.slug }, rota);
          a.textContent = (rankAtual.posicoes[eq.slug] ? rankAtual.posicoes[eq.slug] + 'º · ' : '') + eq.equipe;
          tdNome.appendChild(a);
          tr.appendChild(tdNome);
          var tdScore = document.createElement('td');
          tdScore.className = 'num';
          tdScore.textContent = it && it.score !== null ? it.score.toFixed(2).replace('.', ',') : '—';
          tr.appendChild(tdScore);
          var tdSelo = document.createElement('td');
          if (it) tdSelo.appendChild(elementoSelo(classificarCriticidade(it.score, dados.metas)));
          tr.appendChild(tdSelo);
          corpo.appendChild(tr);
        });
        tabela.appendChild(corpo);
        cEquipes.appendChild(tabela);
        areaConteudo.appendChild(cEquipes);
      }

      // ── Evolução do score ──────────────────────────────────────────
      var cEvolucao = cartao('Evolução do score', 'Linha tracejada horizontal = meta (' + dados.metas.meta_score + ').');
      cEvolucao.className += ' cartao--largo';
      var pontos = competencias.map(function (c) {
        var r = scoreAgregado(equipesEscopo, c, dados.indicadores);
        return { rotulo: formatarCompetencia(c), valor: r.score, extra: r.dadoIncompleto ? ' (dado incompleto)' : '' };
      });
      cEvolucao.appendChild(graficoEvolucao({
        pontos: pontos, meta: dados.metas.meta_score, metas: dados.metas,
        formatarValor: function (v) { return v.toFixed(1).replace('.', ','); },
        aria: 'Evolução do score de qualidade ao longo das competências',
      }));
      areaConteudo.appendChild(cEvolucao);

      // ── Distribuição por inconsistência ─────────────────────────────
      var cDist = cartao('Distribuição por inconsistência', 'Cada barra é a taxa daquela inconsistência sobre o total de cadastrados na competência selecionada. Linha tracejada = meta (' + dados.metas.meta_inconsistencia_pct + '%).');
      cDist.className += ' cartao--largo';
      var itensDist = dados.indicadores.map(function (ind) {
        var d = detalheIndicador(equipesEscopo, competenciaAtual, ind.id);
        return {
          rotulo: ind.titulo, valor: d ? d.pct : null, ausente: d === null,
          n: d ? d.soma : 0, total: d ? d.total : 0,
        };
      });
      cDist.appendChild(graficoBarrasDiretas({
        itens: itensDist, meta: dados.metas.meta_inconsistencia_pct,
        aria: 'Distribuição das inconsistências ativas na competência selecionada',
      }));
      areaConteudo.appendChild(cDist);

      // ── Mapa de calor equipe × inconsistência ───────────────────────
      // Indicador nas LINHAS (rótulo longo, cabe no espaço já reservado à
      // esquerda) e equipe nas COLUNAS (rótulo curto, não precisa rotacionar).
      if (equipesEscopo.length > 1 || rota.tipo === 'rede') {
        var cMapa = cartao('Equipe × inconsistência', 'Cor derivada da distância até a meta de cada indicador (verde = dentro da meta, laranja = até 2x a meta, vermelho = acima disso).');
        cMapa.className += ' cartao--largo';
        var linhasMapa = dados.indicadores.map(function (ind) { return { titulo: ind.titulo, meta: ind.meta_pct, id: ind.id }; });
        var colunasMapa = equipesEscopo.map(function (eq) { return { titulo: eq.equipe, eq: eq }; });
        cMapa.appendChild(graficoMapaCalor({
          linhas: linhasMapa, colunas: colunasMapa,
          obterValor: function (linha, col) {
            var it = col.eq.serie.filter(function (s) { return s.competencia === competenciaAtual; })[0];
            if (!it) return null;
            var ind = it.indicadores.filter(function (i) { return i.id === linha.id; })[0];
            if (!ind || ind.ausente) return null;
            return (ind.valor / it.total_cadastrados) * 100;
          },
          obterMeta: function (linha) { return linha.meta; },
          aria: 'Mapa de calor: taxa de cada inconsistência por equipe',
        }));
        areaConteudo.appendChild(cMapa);
      }
    }

    desenhar(seletorComp.value);
    seletorComp.addEventListener('change', function () { desenhar(seletorComp.value); });
  }

  // ══════════════════════════════════════════════════════════════════════
  //  PÁGINA: EQUIPE
  // ══════════════════════════════════════════════════════════════════════

  function renderizarEquipe(container, dados, rota) {
    var municipio = dados.municipios.filter(function (m) { return m.slug === rota.municipio; })[0];
    var equipe = municipio.equipes.filter(function (e) { return e.slug === rota.equipe; })[0];
    var competencias = equipe.serie.map(function (s) { return s.competencia; });

    var migalha = document.createElement('div');
    migalha.className = 'migalha';
    var linkMun = document.createElement('a');
    linkMun.href = linkPara({ tipo: 'municipio', municipio: municipio.slug }, rota);
    linkMun.textContent = municipio.municipio;
    migalha.appendChild(linkMun);
    migalha.appendChild(document.createTextNode(' / ' + equipe.equipe));
    container.appendChild(migalha);

    var h1 = document.createElement('h1');
    h1.className = 'titulo-pagina';
    h1.textContent = equipe.equipe;
    container.appendChild(h1);

    var linhaFiltros = document.createElement('div');
    linhaFiltros.style.marginBottom = '14px';
    var rotuloComp = document.createElement('label');
    rotuloComp.style.fontSize = '13px';
    rotuloComp.style.color = 'var(--texto-suave)';
    rotuloComp.textContent = 'Competência: ';
    var seletorComp = document.createElement('select');
    seletorComp.className = 'seletor';
    competencias.forEach(function (c) {
      var opt = document.createElement('option');
      opt.value = c;
      opt.textContent = formatarCompetencia(c);
      seletorComp.appendChild(opt);
    });
    seletorComp.value = competencias[competencias.length - 1];
    rotuloComp.appendChild(seletorComp);
    linhaFiltros.appendChild(rotuloComp);
    container.appendChild(linhaFiltros);

    var areaConteudo = document.createElement('div');
    container.appendChild(areaConteudo);

    function desenhar(competenciaAtual) {
      areaConteudo.innerHTML = '';
      var item = equipe.serie.filter(function (s) { return s.competencia === competenciaAtual; })[0];
      var rank = posicaoRanking(municipio.equipes, competenciaAtual);
      var posicao = rank.posicoes[equipe.slug];

      // ── a) Score + criticidade + posição ──────────────────────────
      var grade = document.createElement('div');
      grade.className = 'grade';

      var cScore = cartao('Score de Qualidade', 'Score = 100 − (Σ das inconsistências que entram no score ÷ total de cadastrados × 100).');
      var vScore = document.createElement('div');
      vScore.className = 'stat-valor';
      vScore.textContent = item.score === null ? '—' : item.score.toFixed(2).replace('.', ',');
      cScore.appendChild(vScore);
      var linhaSelos = document.createElement('div');
      linhaSelos.style.marginTop = '8px';
      linhaSelos.appendChild(elementoSelo(classificarCriticidade(item.score, dados.metas)));
      if (item.dado_incompleto) {
        var selo2 = elementoSelo(null, 'Dado incompleto');
        selo2.style.marginLeft = '6px';
        linhaSelos.appendChild(selo2);
      }
      cScore.appendChild(linhaSelos);
      grade.appendChild(cScore);

      var cPosicao = cartao('Posição no município', 'Ranking por score dentro do município nesta competência — 1º = melhor score.');
      var vPosicao = document.createElement('div');
      vPosicao.className = 'stat-valor';
      vPosicao.textContent = posicao ? (posicao + 'º de ' + rank.total) : '—';
      cPosicao.appendChild(vPosicao);
      grade.appendChild(cPosicao);

      var cTotal = cartao('Total de cadastrados ativos');
      var vTotal = document.createElement('div');
      vTotal.className = 'stat-valor';
      vTotal.textContent = formatarNumero(item.total_cadastrados);
      cTotal.appendChild(vTotal);
      grade.appendChild(cTotal);

      areaConteudo.appendChild(grade);

      // ── a-continuação) Estado do cadastro: waffle de limpos + distribuição ──
      if (item.estado) {
        var cEstado = cartao('Estado do cadastro', '"Limpo" = nenhuma das ' + dados.indicadores.length + ' inconsistências ativas presente no cadastro.');
        cEstado.className += ' cartao--largo';
        var linhaEstado = document.createElement('div');
        linhaEstado.style.display = 'flex';
        linhaEstado.style.gap = '28px';
        linhaEstado.style.flexWrap = 'wrap';
        linhaEstado.style.alignItems = 'center';

        var totalEstado = item.estado.limpos + item.estado.com_1 + item.estado.com_2 + item.estado.com_3_mais;
        var pctLimpos = totalEstado ? (item.estado.limpos / totalEstado) * 100 : 0;
        var blocoWaffle = document.createElement('div');
        blocoWaffle.appendChild(construirWaffle(pctLimpos));
        var legendaWaffle = document.createElement('div');
        legendaWaffle.className = 'stat-legenda';
        legendaWaffle.style.marginTop = '8px';
        legendaWaffle.textContent = formatarPercentual(item.estado.limpos, totalEstado) + ' limpos (' + formatarNumero(item.estado.limpos) + ' de ' + formatarNumero(totalEstado) + ')';
        blocoWaffle.appendChild(legendaWaffle);
        linhaEstado.appendChild(blocoWaffle);

        var blocoDist = document.createElement('div');
        blocoDist.style.flex = '1';
        blocoDist.style.minWidth = '220px';
        [
          ['0 inconsistências', item.estado.limpos],
          ['1 inconsistência', item.estado.com_1],
          ['2 inconsistências', item.estado.com_2],
          ['3 ou mais', item.estado.com_3_mais],
        ].forEach(function (par) {
          var linha = document.createElement('div');
          linha.style.display = 'flex';
          linha.style.justifyContent = 'space-between';
          linha.style.gap = '10px';
          linha.style.padding = '3px 0';
          linha.style.fontSize = '13px';
          var rot = document.createElement('span');
          rot.className = 'texto-suave';
          rot.textContent = par[0];
          var val = document.createElement('span');
          val.textContent = formatarNumero(par[1]) + ' (' + formatarPercentual(par[1], totalEstado) + ')';
          linha.appendChild(rot);
          linha.appendChild(val);
          blocoDist.appendChild(linha);
        });
        var linhaMedia = document.createElement('div');
        linhaMedia.className = 'texto-fraco';
        linhaMedia.style.marginTop = '8px';
        linhaMedia.textContent = 'Média de ' + item.estado.media_inconsistencias.toFixed(2).replace('.', ',') + ' inconsistências por cadastro.';
        blocoDist.appendChild(linhaMedia);
        if (item.estado.fora_area > 0) {
          var linhaFora = document.createElement('div');
          linhaFora.className = 'texto-fraco';
          linhaFora.textContent = formatarNumero(item.estado.fora_area) + ' cadastro(s) "FORA DE ÁREA" (não entram em nenhuma estatística acima).';
          blocoDist.appendChild(linhaFora);
        }
        linhaEstado.appendChild(blocoDist);

        cEstado.appendChild(linhaEstado);
        areaConteudo.appendChild(cEstado);
      }

      // ── b) Faltam N por indicador, com maior concentração por microárea ──
      var cFaltam = cartao('Faltam para a meta', 'Quantos cadastros a mais cada indicador precisa corrigir para a equipe ficar dentro da meta.');
      cFaltam.className += ' cartao--largo';
      var gradeFaltam = document.createElement('div');
      gradeFaltam.className = 'grade';
      dados.indicadores.forEach(function (ind) {
        var indItem = item.indicadores.filter(function (i) { return i.id === ind.id; })[0];
        var sub = document.createElement('div');
        sub.style.padding = '10px 0';

        var titulo = document.createElement('div');
        titulo.className = 'texto-suave';
        titulo.style.fontSize = '12.5px';
        titulo.style.fontWeight = '700';
        titulo.textContent = ind.titulo;
        sub.appendChild(titulo);

        if (!indItem || indItem.ausente) {
          var semDado = document.createElement('div');
          semDado.className = 'stat-legenda';
          semDado.textContent = 'Dado incompleto nesta competência.';
          sub.appendChild(semDado);
          gradeFaltam.appendChild(sub);
          return;
        }

        var metaN = Math.ceil((item.total_cadastrados * ind.meta_pct) / 100);
        var faltam = Math.max(0, indItem.valor - metaN);
        var valorFaltam = document.createElement('div');
        valorFaltam.className = 'stat-valor';
        valorFaltam.style.fontSize = '24px';
        valorFaltam.style.color = faltam > 0 ? 'var(--cor-critica)' : 'var(--cor-adequada)';
        valorFaltam.textContent = faltam > 0 ? formatarNumero(faltam) : 'Meta OK';
        sub.appendChild(valorFaltam);

        var legenda = document.createElement('div');
        legenda.className = 'stat-legenda';
        legenda.textContent = formatarNumero(indItem.valor) + ' de ' + formatarNumero(item.total_cadastrados) + ' (' + formatarPercentual(indItem.valor, item.total_cadastrados) + ')';
        sub.appendChild(legenda);

        if (item.microareas && faltam > 0) {
          var maior = null;
          item.microareas.forEach(function (ma) {
            var mind = ma.indicadores.filter(function (i) { return i.id === ind.id; })[0];
            if (!mind || mind.ausente) return;
            var metaMa = Math.ceil((ma.total_cadastrados * ind.meta_pct) / 100);
            var faltamMa = Math.max(0, mind.valor - metaMa);
            if (faltamMa > 0 && (!maior || faltamMa > maior.faltam)) maior = { nome: ma.microarea, faltam: faltamMa };
          });
          if (maior) {
            var conc = document.createElement('div');
            conc.className = 'texto-fraco';
            conc.textContent = 'Maior concentração: ' + maior.nome + ' (' + formatarNumero(maior.faltam) + ')';
            sub.appendChild(conc);
          }
        }
        gradeFaltam.appendChild(sub);
      });
      cFaltam.appendChild(gradeFaltam);
      areaConteudo.appendChild(cFaltam);

      // ── c) Trajetória de cada inconsistência + projeção ─────────────
      dados.indicadores.forEach(function (ind) {
        var pontosSerie = competencias.map(function (c, i) {
          return { i: i, rotulo: formatarCompetencia(c), valor: taxaIndicador([equipe], c, ind.id) };
        });
        if (pontosSerie.every(function (p) { return p.valor === null; })) return;

        var proj = projetarMeta(pontosSerie, ind.meta_pct, competenciaAtual);
        var cTraj = cartao(ind.titulo, ind.descricao_curta);
        cTraj.className += ' cartao--largo';
        cTraj.appendChild(graficoEvolucao({
          pontos: pontosSerie, meta: ind.meta_pct,
          corPonto: function (v) { return corPorTaxaVsMeta(v, ind.meta_pct); },
          projecao: proj.segmentoProjecao,
          formatarValor: function (v) { return v.toFixed(1).replace('.', ',') + '%'; },
          aria: 'Trajetória de ' + ind.titulo,
        }));

        var frase = document.createElement('p');
        frase.className = 'stat-legenda';
        if (proj.status === 'meta_atingida') frase.textContent = 'Meta já atingida nesta competência.';
        else if (proj.status === 'sem_queda') frase.textContent = 'No ritmo atual, sem queda — a taxa não vem melhorando nas últimas competências.';
        else if (proj.status === 'projetado') frase.textContent = 'No ritmo atual, meta em ' + formatarCompetencia(proj.competenciaMeta) + '.';
        else frase.textContent = 'Ainda não há competências suficientes para estimar uma tendência.';
        cTraj.appendChild(frase);
        areaConteudo.appendChild(cTraj);
      });

      // ── d) Mapa de calor microárea × indicador + Pareto ─────────────
      if (item.microareas && item.microareas.length) {
        var cMapaMa = cartao('Microárea × inconsistência', 'Cor derivada da distância até a meta de cada indicador.');
        cMapaMa.className += ' cartao--largo';
        var linhasMapaMa = dados.indicadores.map(function (ind) { return { titulo: ind.titulo, meta: ind.meta_pct, id: ind.id }; });
        var colunasMapaMa = item.microareas.map(function (ma) { return { titulo: ma.microarea, ma: ma }; });
        cMapaMa.appendChild(graficoMapaCalor({
          linhas: linhasMapaMa, colunas: colunasMapaMa,
          obterValor: function (linha, col) {
            var ind = col.ma.indicadores.filter(function (i) { return i.id === linha.id; })[0];
            if (!ind || ind.ausente) return null;
            return (ind.valor / col.ma.total_cadastrados) * 100;
          },
          obterMeta: function (linha) { return linha.meta; },
          aria: 'Mapa de calor: taxa de cada inconsistência por microárea',
        }));
        areaConteudo.appendChild(cMapaMa);

        var cPareto = cartao('Concentração por microárea (Pareto)', 'Soma de todas as inconsistências ativas da microárea (barras) e % acumulado (linha) — mostra se o problema está espalhado ou concentrado em poucas microáreas.');
        cPareto.className += ' cartao--largo';
        var itensPareto = item.microareas.map(function (ma) {
          var soma = ma.indicadores.reduce(function (s, ind) { return s + (ind.ausente ? 0 : ind.valor); }, 0);
          return { rotulo: ma.microarea, valor: soma, total: ma.total_cadastrados };
        });
        cPareto.appendChild(graficoPareto({ itens: itensPareto, aria: 'Curva de Pareto das inconsistências por microárea' }));

        var ordenadasPareto = itensPareto.slice().sort(function (a, b) { return b.valor - a.valor; });
        var nPioresMa = Math.max(1, Math.round(ordenadasPareto.length * 0.2));
        var piores = ordenadasPareto.slice(0, nPioresMa);
        var somaTotalInc = itensPareto.reduce(function (s, it) { return s + it.valor; }, 0);
        var somaPioresInc = piores.reduce(function (s, it) { return s + it.valor; }, 0);
        var somaTotalCad = itensPareto.reduce(function (s, it) { return s + it.total; }, 0);
        var somaPioresCad = piores.reduce(function (s, it) { return s + it.total; }, 0);
        var concentracao = document.createElement('p');
        concentracao.className = 'stat-legenda';
        concentracao.textContent = 'As ' + nPioresMa + ' microáreas mais críticas (20% do total) concentram ' +
          formatarPercentual(somaPioresInc, somaTotalInc) + ' das inconsistências, e representam ' +
          formatarPercentual(somaPioresCad, somaTotalCad) + ' dos cadastros.';
        cPareto.appendChild(concentracao);

        areaConteudo.appendChild(cPareto);
      } else {
        var semMicro = document.createElement('p');
        semMicro.className = 'texto-fraco';
        semMicro.textContent = 'Sem dado de microárea para esta competência.';
        areaConteudo.appendChild(semMicro);
      }

      // ── e) Pirâmide etária espelhada (com CPF × sem CPF) ─────────────
      if (item.faixa_etaria && item.faixa_etaria.length) {
        var cPiramide = cartao('Sem CPF por faixa etária', 'Cada faixa mostra quantos cadastros ativos têm CPF (esquerda) e quantos não têm (direita).');
        cPiramide.className += ' cartao--largo';
        var ordemFaixas = ['<1', '1-4', '5-9', '10-17', '18-59', '60+'];
        var faixasOrdenadas = item.faixa_etaria
          .filter(function (f) { return f.indicador_id === 'sem_cpf'; })
          .map(function (f) { return { faixa: f.faixa, semCpf: f.valor, comCpf: f.total_cadastrados - f.valor }; })
          .sort(function (a, b) { return ordemFaixas.indexOf(a.faixa) - ordemFaixas.indexOf(b.faixa); });
        cPiramide.appendChild(graficoPiramide(faixasOrdenadas, 'Cadastros com e sem CPF por faixa etária'));
        areaConteudo.appendChild(cPiramide);
      }

      // ── Impacto em grupos prioritários ────────────────────────────────
      if (item.grupos && item.grupos.length) {
        var cGrupos = cartao('Impacto em grupos prioritários', '% sem CPF, % sem CNS e % com alguma falha de identificação ou vínculo (sem CPF, sem CNS ou não vinculado à família), dentro de cada grupo.');
        cGrupos.className += ' cartao--largo';
        var tabelaG = document.createElement('table');
        tabelaG.className = 'tabela';
        var theadG = document.createElement('thead');
        theadG.innerHTML = '<tr><th>Grupo</th><th>Cadastros</th><th>Sem CPF</th><th>Sem CNS</th><th>Falha de identificação/vínculo</th></tr>';
        tabelaG.appendChild(theadG);
        var corpoG = document.createElement('tbody');
        item.grupos.forEach(function (g) {
          var catalogo = dados.grupos.filter(function (dg) { return dg.id === g.id; })[0];
          var tr = document.createElement('tr');
          var tdNome = document.createElement('td');
          tdNome.style.fontWeight = '700';
          tdNome.textContent = catalogo ? catalogo.titulo : g.id;
          tr.appendChild(tdNome);
          var tdTotal = document.createElement('td');
          tdTotal.className = 'num';
          tdTotal.textContent = formatarNumero(g.total_grupo);
          tr.appendChild(tdTotal);
          ['sem_cpf', 'sem_cns', 'falha_identificacao_vinculo'].forEach(function (idMetrica) {
            var m = g.metricas.filter(function (x) { return x.id === idMetrica; })[0];
            var td = document.createElement('td');
            td.className = 'num';
            td.textContent = m ? formatarPercentual(m.valor, g.total_grupo) + ' (' + formatarNumero(m.valor) + ')' : '—';
            tr.appendChild(td);
          });
          corpoG.appendChild(tr);
        });
        tabelaG.appendChild(corpoG);
        cGrupos.appendChild(tabelaG);
        areaConteudo.appendChild(cGrupos);
      }

      // ── Perfil de condições de saúde + alertas de plausibilidade ──────
      if (item.condicoes && item.condicoes.length) {
        var alertasPlausibilidade = calcularAlertasPlausibilidade(dados.condicoes, item.condicoes, dados.minimo_cadastros_para_alerta);

        if (alertasPlausibilidade.length) {
          var cAlertas = cartao('Alertas de plausibilidade', 'Sem uma taxa clínica esperada para a maioria das condições, o critério é: zero casos numa equipe deste tamanho é estatisticamente improvável — quase sempre sinal de campo não preenchido, não de "está tudo bem".');
          cAlertas.className += ' cartao--largo';
          alertasPlausibilidade.forEach(function (a) {
            var linha = document.createElement('div');
            linha.style.display = 'flex';
            linha.style.gap = '8px';
            linha.style.padding = '6px 0';
            linha.style.fontSize = '13.5px';
            linha.style.borderLeft = '3px solid var(--cor-critica)';
            linha.style.paddingLeft = '10px';
            var texto = document.createElement('span');
            texto.textContent = a.mensagem;
            linha.appendChild(texto);
            cAlertas.appendChild(linha);
          });
          areaConteudo.appendChild(cAlertas);
        }

        var cCondicoes = cartao('Condições de saúde', 'Distribuição das condições registradas em CONDICOES DE SAUDE, sobre o total de cadastrados ativos.');
        cCondicoes.className += ' cartao--largo';
        var tabelaC = document.createElement('table');
        tabelaC.className = 'tabela';
        var theadC = document.createElement('thead');
        theadC.innerHTML = '<tr><th>Condição</th><th>Cadastros</th><th>%</th></tr>';
        tabelaC.appendChild(theadC);
        var corpoC = document.createElement('tbody');
        item.condicoes.slice().sort(function (a, b) { return b.valor - a.valor; }).forEach(function (c) {
          var cat = dados.condicoes.filter(function (dc) { return dc.id === c.id; })[0];
          var tr = document.createElement('tr');
          var tdNome = document.createElement('td');
          tdNome.textContent = cat ? cat.titulo : c.id;
          tr.appendChild(tdNome);
          var tdN = document.createElement('td');
          tdN.className = 'num';
          tdN.textContent = formatarNumero(c.valor);
          tr.appendChild(tdN);
          var tdPct = document.createElement('td');
          tdPct.className = 'num';
          tdPct.textContent = formatarPercentual(c.valor, c.total_cadastrados);
          tr.appendChild(tdPct);
          corpoC.appendChild(tr);
        });
        tabelaC.appendChild(corpoC);
        cCondicoes.appendChild(tabelaC);
        areaConteudo.appendChild(cCondicoes);
      }

      // ── h) Orientação de correção ────────────────────────────────────
      var cOrient = cartao('Como corrigir cada inconsistência');
      cOrient.className += ' cartao--largo';
      dados.indicadores.forEach(function (ind) {
        var det = document.createElement('details');
        det.style.marginBottom = '6px';
        var sum = document.createElement('summary');
        sum.style.cursor = 'pointer';
        sum.style.fontWeight = '700';
        sum.textContent = ind.titulo;
        det.appendChild(sum);
        var p = document.createElement('p');
        p.className = 'texto-suave';
        p.style.marginLeft = '6px';
        p.textContent = ind.orientacao;
        det.appendChild(p);
        cOrient.appendChild(det);
      });
      areaConteudo.appendChild(cOrient);

      // ── i) Botão grande + última atualização ─────────────────────────
      var cAcao = cartao('Planilhas de correção');
      cAcao.className += ' cartao--largo';
      if (equipe.link_drive) {
        var botao = document.createElement('a');
        botao.className = 'botao-grande';
        botao.href = equipe.link_drive;
        botao.target = '_blank';
        botao.rel = 'noopener';
        botao.textContent = '📂 Abrir planilhas da equipe';
        cAcao.appendChild(botao);
      } else {
        var semLink = document.createElement('p');
        semLink.className = 'texto-suave';
        semLink.textContent = 'Link da pasta no Drive ainda não cadastrado em config/equipes.yaml.';
        cAcao.appendChild(semLink);
      }
      var dataAtt = document.createElement('p');
      dataAtt.className = 'texto-fraco';
      dataAtt.style.marginTop = '10px';
      dataAtt.textContent = 'Última atualização dos dados: ' + new Date(dados.gerado_em).toLocaleString('pt-BR');
      cAcao.appendChild(dataAtt);
      areaConteudo.appendChild(cAcao);
    }

    desenhar(seletorComp.value);
    seletorComp.addEventListener('change', function () { desenhar(seletorComp.value); });
  }

  // ══════════════════════════════════════════════════════════════════════
  //  PÁGINA: RANKING (rede inteira — um único /ranking/, sem variante por município)
  // ══════════════════════════════════════════════════════════════════════

  function todasEquipes(dados) {
    var lista = [];
    dados.municipios.forEach(function (m) {
      m.equipes.forEach(function (e) {
        lista.push({
          slug: e.slug, equipe: e.equipe, serie: e.serie,
          municipioSlug: m.slug, municipio: m.municipio,
        });
      });
    });
    return lista;
  }

  // Cor de IDENTIFICAÇÃO (não de significado) para distinguir as linhas do
  // bump chart — aqui a cor não diz "bom/ruim", só "essa é a equipe X", por
  // isso cada linha também termina com o nome por extenso (rótulo direto),
  // a cor nunca é a única forma de saber qual equipe é qual.
  function corIdentidade(indice) {
    var matiz = (indice * 47) % 360;
    return 'hsl(' + matiz + ', 50%, 42%)';
  }

  function renderizarRanking(container, dados, rota) {
    var equipes = todasEquipes(dados);
    var competencias = [];
    equipes.forEach(function (eq) {
      eq.serie.forEach(function (s) { if (competencias.indexOf(s.competencia) === -1) competencias.push(s.competencia); });
    });
    competencias.sort();

    var h1 = document.createElement('h1');
    h1.className = 'titulo-pagina';
    h1.textContent = 'Ranking';
    container.appendChild(h1);

    var linhaFiltros = document.createElement('div');
    linhaFiltros.style.marginBottom = '14px';
    var rotuloComp = document.createElement('label');
    rotuloComp.style.fontSize = '13px';
    rotuloComp.style.color = 'var(--texto-suave)';
    rotuloComp.textContent = 'Competência: ';
    var seletorComp = document.createElement('select');
    seletorComp.className = 'seletor';
    competencias.forEach(function (c) {
      var opt = document.createElement('option');
      opt.value = c;
      opt.textContent = formatarCompetencia(c);
      seletorComp.appendChild(opt);
    });
    seletorComp.value = competencias[competencias.length - 1];
    rotuloComp.appendChild(seletorComp);
    linhaFiltros.appendChild(rotuloComp);
    container.appendChild(linhaFiltros);

    var areaConteudo = document.createElement('div');
    container.appendChild(areaConteudo);
    var modoPrioridade = false;

    function linkEquipe(eq) {
      var a = document.createElement('a');
      a.className = 'linha-equipe';
      a.href = linkPara({ tipo: 'equipe', municipio: eq.municipioSlug, equipe: eq.slug }, rota);
      a.textContent = eq.equipe;
      return a;
    }

    function desenhar(competenciaAtual) {
      areaConteudo.innerHTML = '';
      var compAnterior = competenciaAnteriorNaSerie(competencias, competenciaAtual);
      var rank = posicaoRanking(equipes, competenciaAtual);

      var linhas = equipes.map(function (eq) {
        var it = eq.serie.filter(function (s) { return s.competencia === competenciaAtual; })[0];
        var itAnt = compAnterior ? eq.serie.filter(function (s) { return s.competencia === compAnterior; })[0] : null;
        var variacao = (it && itAnt && it.score !== null && itAnt.score !== null) ? it.score - itAnt.score : null;
        return { eq: eq, score: it ? it.score : null, variacao: variacao };
      }).filter(function (l) { return l.score !== null; });

      // ── Ranking por score / Prioridade de apoio (mesma tabela, ordem invertida) ──
      var cRank = cartao(modoPrioridade ? 'Prioridade de apoio' : 'Ranking por score',
        modoPrioridade
          ? 'Mesma lista do ranking, em ordem inversa — quem mais precisa de apoio primeiro.'
          : '1º = melhor Score de Qualidade.');
      cRank.className += ' cartao--largo';

      var botaoModo = document.createElement('button');
      botaoModo.type = 'button';
      botaoModo.className = 'seletor';
      botaoModo.style.marginBottom = '10px';
      botaoModo.style.cursor = 'pointer';
      botaoModo.textContent = modoPrioridade ? 'Ver ranking por score' : 'Ver prioridade de apoio';
      botaoModo.addEventListener('click', function () { modoPrioridade = !modoPrioridade; desenhar(competenciaAtual); });
      cRank.appendChild(botaoModo);

      var ordenadas = linhas.slice().sort(function (a, b) {
        return modoPrioridade ? a.score - b.score : b.score - a.score;
      });
      var tabela = document.createElement('table');
      tabela.className = 'tabela';
      var thead = document.createElement('thead');
      thead.innerHTML = '<tr><th>#</th><th>Equipe</th><th>Município</th><th>Score</th><th>Situação</th><th>Variação</th></tr>';
      tabela.appendChild(thead);
      var corpo = document.createElement('tbody');
      ordenadas.forEach(function (l) {
        var tr = document.createElement('tr');
        var tdPos = document.createElement('td');
        tdPos.textContent = rank.posicoes[l.eq.slug] + 'º';
        tr.appendChild(tdPos);
        var tdEq = document.createElement('td');
        tdEq.appendChild(linkEquipe(l.eq));
        tr.appendChild(tdEq);
        var tdMun = document.createElement('td');
        tdMun.className = 'texto-suave';
        tdMun.textContent = l.eq.municipio;
        tr.appendChild(tdMun);
        var tdScore = document.createElement('td');
        tdScore.className = 'num';
        tdScore.textContent = l.score.toFixed(2).replace('.', ',');
        tr.appendChild(tdScore);
        var tdSelo = document.createElement('td');
        tdSelo.appendChild(elementoSelo(classificarCriticidade(l.score, dados.metas)));
        tr.appendChild(tdSelo);
        var tdVar = document.createElement('td');
        tdVar.className = 'num';
        if (l.variacao !== null) {
          tdVar.textContent = formatarPontoPercentual(l.variacao) + ' pp';
          tdVar.style.color = l.variacao >= 0 ? 'var(--cor-adequada)' : 'var(--cor-critica)';
        } else {
          tdVar.textContent = '—';
        }
        tr.appendChild(tdVar);
        corpo.appendChild(tr);
      });
      tabela.appendChild(corpo);
      cRank.appendChild(tabela);
      areaConteudo.appendChild(cRank);

      // ── Ranking de evolução (maior melhora primeiro) ────────────────
      if (compAnterior) {
        var cEvo = cartao('Ranking de evolução', 'Variação do score em relação a ' + formatarCompetencia(compAnterior) + ' — maior melhora no topo.');
        cEvo.className += ' cartao--largo';
        var comVariacao = linhas.filter(function (l) { return l.variacao !== null; })
          .sort(function (a, b) { return b.variacao - a.variacao; });
        if (comVariacao.length) {
          var tabelaEvo = document.createElement('table');
          tabelaEvo.className = 'tabela';
          var theadEvo = document.createElement('thead');
          theadEvo.innerHTML = '<tr><th>#</th><th>Equipe</th><th>Município</th><th>Variação</th><th>Score atual</th></tr>';
          tabelaEvo.appendChild(theadEvo);
          var corpoEvo = document.createElement('tbody');
          comVariacao.forEach(function (l, i) {
            var tr = document.createElement('tr');
            var tdPos = document.createElement('td'); tdPos.textContent = (i + 1) + 'º'; tr.appendChild(tdPos);
            var tdEq = document.createElement('td'); tdEq.appendChild(linkEquipe(l.eq)); tr.appendChild(tdEq);
            var tdMun = document.createElement('td'); tdMun.className = 'texto-suave'; tdMun.textContent = l.eq.municipio; tr.appendChild(tdMun);
            var tdVar = document.createElement('td');
            tdVar.className = 'num';
            tdVar.style.color = l.variacao >= 0 ? 'var(--cor-adequada)' : 'var(--cor-critica)';
            tdVar.textContent = formatarPontoPercentual(l.variacao) + ' pp';
            tr.appendChild(tdVar);
            var tdScore = document.createElement('td'); tdScore.className = 'num'; tdScore.textContent = l.score.toFixed(2).replace('.', ','); tr.appendChild(tdScore);
            corpoEvo.appendChild(tr);
          });
          tabelaEvo.appendChild(corpoEvo);
          cEvo.appendChild(tabelaEvo);
        }
        areaConteudo.appendChild(cEvo);
      }

      // ── Bump chart: posição ao longo dos meses ──────────────────────
      if (competencias.length > 1) {
        var cBump = cartao('Posição ao longo do tempo', 'Cada linha é uma equipe; a cor só identifica (não indica bom/ruim) — por isso todo mundo tem o nome escrito na ponta.');
        cBump.className += ' cartao--largo';
        cBump.appendChild(graficoBump({ equipes: equipes, competencias: competencias, aria: 'Posição de cada equipe no ranking ao longo das competências' }));
        areaConteudo.appendChild(cBump);
      }
    }

    desenhar(seletorComp.value);
    seletorComp.addEventListener('change', function () { desenhar(seletorComp.value); });
  }

  // ── Bump chart genérico ────────────────────────────────────────────────
  function graficoBump(opcoes) {
    var margem = { esq: 16, dir: 150, topo: 20, baixo: 30 };
    var alturaLinha = 26;
    var w = Math.max(360, opcoes.competencias.length * 130);
    var h = opcoes.equipes.length * alturaLinha;
    var largura = margem.esq + w + margem.dir, altura = margem.topo + h + margem.baixo;

    var svg = svgEl('svg', { class: 'grafico', viewBox: '0 0 ' + largura + ' ' + altura, role: 'img', 'aria-label': opcoes.aria });
    var g = svgEl('g', { transform: 'translate(' + margem.esq + ',' + margem.topo + ')' });
    svg.appendChild(g);

    var passoX = opcoes.competencias.length > 1 ? w / (opcoes.competencias.length - 1) : 0;
    function x(i) { return passoX * i; }
    function y(posicao) { return (posicao - 1) * alturaLinha + alturaLinha / 2; }

    opcoes.competencias.forEach(function (c, i) {
      var t = svgEl('text', { class: 'rotulo-eixo', x: x(i), y: h + 20, 'text-anchor': 'middle' });
      t.textContent = formatarCompetencia(c);
      g.appendChild(t);
    });

    opcoes.equipes.forEach(function (eq, indiceEquipe) {
      var cor = corIdentidade(indiceEquipe);
      var pontos = [];
      opcoes.competencias.forEach(function (c, i) {
        var temDado = eq.serie.some(function (s) { return s.competencia === c && s.score !== null; });
        if (!temDado) return;
        var rank = posicaoRanking(opcoes.equipes.filter(function (e) {
          return e.serie.some(function (s) { return s.competencia === c && s.score !== null; });
        }), c);
        pontos.push({ i: i, posicao: rank.posicoes[eq.slug] });
      });
      if (!pontos.length) return;

      var d = pontos.map(function (p, k) { return (k === 0 ? 'M' : 'L') + x(p.i).toFixed(1) + ',' + y(p.posicao).toFixed(1); }).join(' ');
      var linha = svgEl('path', { d: d, fill: 'none', 'stroke-width': 2.5 });
      linha.style.stroke = cor;
      g.appendChild(linha);

      pontos.forEach(function (p) {
        var ponto = svgEl('circle', { cx: x(p.i), cy: y(p.posicao), r: 4 });
        ponto.style.fill = cor;
        ponto.addEventListener('mousemove', function (ev) {
          mostrarTooltip(ev, '<strong>' + eq.equipe + '</strong><br>' + formatarCompetencia(opcoes.competencias[p.i]) + ': ' + p.posicao + 'º lugar');
        });
        ponto.addEventListener('mouseleave', esconderTooltip);
        g.appendChild(ponto);
      });

      var ultimo = pontos[pontos.length - 1];
      var rotulo = svgEl('text', {
        x: x(ultimo.i) + 10, y: y(ultimo.posicao) + 4, class: 'rotulo-direto',
      });
      rotulo.textContent = eq.equipe;
      rotulo.style.fill = cor;
      g.appendChild(rotulo);
    });

    return svg;
  }

  // ══════════════════════════════════════════════════════════════════════
  //  PÁGINA: METODOLOGIA — gerada a partir dos YAML, não é texto fixo: os
  //  números (metas, lista de indicadores, se cada um entra no score) vêm
  //  todos de dados.metas/dados.indicadores/dados.grupos. Só a explicação
  //  de COMO o cálculo funciona (a fórmula em si) é texto fixo aqui.
  // ══════════════════════════════════════════════════════════════════════

  function secaoTexto(titulo, paragrafos) {
    var cartaoEl = cartao(titulo);
    cartaoEl.className += ' cartao--largo';
    paragrafos.forEach(function (texto) {
      var p = document.createElement('p');
      p.className = 'texto-suave';
      p.innerHTML = texto;
      cartaoEl.appendChild(p);
    });
    return cartaoEl;
  }

  function renderizarMetodologia(container, dados) {
    var h1 = document.createElement('h1');
    h1.className = 'titulo-pagina';
    h1.textContent = 'Metodologia';
    container.appendChild(h1);

    container.appendChild(secaoTexto('Score de Qualidade', [
      'Score = 100 − (Σ das inconsistências que <strong>entram no score</strong> ÷ total de cadastrados ativos × 100).',
      'Meta: <strong>' + dados.metas.meta_score + '</strong> ou mais. O score é sempre travado entre 0 e 100.',
      'Quando faltar algum indicador que entra no score numa competência (ex.: a exportação daquele mês não trouxe dado suficiente), ele é calculado SEM esse indicador — nunca tratando o valor ausente como zero — e a competência ganha o selo "dado incompleto".',
      'Entre equipes ou entre meses, a soma é sempre dos numeradores e dos denominadores primeiro, e só depois se calcula o percentual — nunca a média dos percentuais de cada equipe.',
    ]));

    container.appendChild(secaoTexto('Criticidade', [
      '<span class="selo selo--adequada">Adequada</span> score ≥ ' + dados.metas.criticidade.adequada + '.',
      '<span class="selo selo--atencao">Atenção</span> score ≥ ' + dados.metas.criticidade.atencao + ' e < ' + dados.metas.criticidade.adequada + '.',
      '<span class="selo selo--critica">Crítica</span> score < ' + dados.metas.criticidade.atencao + '.',
    ]));

    container.appendChild(secaoTexto('Ranking', [
      'Ranking por score: <strong>1º = melhor score</strong>. Empate divide a mesma posição (ranking denso — não pula número).',
      '"Prioridade de apoio" é a mesma lista em ordem inversa: quem mais precisa de apoio aparece primeiro.',
      'Maior melhora / maior piora são calculados automaticamente comparando a competência mais recente com a anterior — não é preciso escolher o mês.',
    ]));

    var cInd = cartao('Indicadores de inconsistência', 'Meta de cada indicador: no máximo ' + dados.metas.meta_inconsistencia_pct + '% do total de cadastrados ativos da equipe.');
    cInd.className += ' cartao--largo';
    var tabela = document.createElement('table');
    tabela.className = 'tabela';
    var thead = document.createElement('thead');
    thead.innerHTML = '<tr><th>Indicador</th><th>O que é</th><th>Meta</th><th>Entra no score?</th></tr>';
    tabela.appendChild(thead);
    var corpo = document.createElement('tbody');
    dados.indicadores.forEach(function (ind) {
      var tr = document.createElement('tr');
      var tdNome = document.createElement('td');
      tdNome.style.fontWeight = '700';
      tdNome.textContent = ind.titulo;
      tr.appendChild(tdNome);
      var tdDesc = document.createElement('td');
      tdDesc.className = 'texto-suave';
      tdDesc.textContent = ind.descricao_curta;
      tr.appendChild(tdDesc);
      var tdMeta = document.createElement('td');
      tdMeta.className = 'num';
      tdMeta.textContent = '≤ ' + ind.meta_pct + '%';
      tr.appendChild(tdMeta);
      var tdScore = document.createElement('td');
      tdScore.textContent = ind.entra_no_score ? 'Sim' : 'Não';
      tr.appendChild(tdScore);
      corpo.appendChild(tr);
    });
    tabela.appendChild(corpo);
    cInd.appendChild(tabela);
    container.appendChild(cInd);

    if (dados.grupos && dados.grupos.length) {
      var cGrupos = cartao('Grupos prioritários', 'Usados nos indicadores complementares de impacto em populações prioritárias.');
      cGrupos.className += ' cartao--largo';
      var lista = document.createElement('ul');
      lista.style.margin = '0';
      dados.grupos.forEach(function (g) {
        var li = document.createElement('li');
        li.className = 'texto-suave';
        li.textContent = g.titulo;
        lista.appendChild(li);
      });
      cGrupos.appendChild(lista);
      container.appendChild(cGrupos);
    }
  }

  // ══════════════════════════════════════════════════════════════════════
  //  BOOT
  // ══════════════════════════════════════════════════════════════════════

  function iniciar() {
    aplicarTemaSalvo();
    var dados = JSON.parse(document.getElementById('dados-pagina').textContent);
    var rota = JSON.parse(document.getElementById('rota-pagina').textContent);
    var raiz = document.getElementById('app');

    raiz.appendChild(montarCabecalho(dados, rota));

    var pagina = document.createElement('main');
    pagina.className = 'pagina';
    raiz.appendChild(pagina);

    if (rota.tipo === 'rede' || rota.tipo === 'municipio') {
      renderizarVisaoGeral(pagina, dados, rota);
    } else if (rota.tipo === 'equipe') {
      renderizarEquipe(pagina, dados, rota);
    } else if (rota.tipo === 'ranking') {
      renderizarRanking(pagina, dados, rota);
    } else if (rota.tipo === 'metodologia') {
      renderizarMetodologia(pagina, dados);
    } else {
      var aviso = document.createElement('p');
      aviso.className = 'texto-suave';
      aviso.textContent = 'Esta página ainda está em construção.';
      pagina.appendChild(aviso);
    }

    var rodape = document.createElement('div');
    rodape.className = 'rodape';
    rodape.textContent = 'Dados gerados em ' + new Date(dados.gerado_em).toLocaleString('pt-BR') + ' · Prima Qualitá Saúde';
    pagina.appendChild(rodape);
  }

  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', iniciar);
    } else {
      iniciar();
    }
  }

  // Exposto só para os testes automatizados (Node, sem DOM) — nenhuma
  // função aqui embaixo toca no document/window além do necessário.
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
      formatarNumero: formatarNumero,
      formatarPercentual: formatarPercentual,
      formatarPontoPercentual: formatarPontoPercentual,
      formatarCompetencia: formatarCompetencia,
      classificarCriticidade: classificarCriticidade,
      scoreAgregado: scoreAgregado,
      percentualInconsistenciaGlobal: percentualInconsistenciaGlobal,
      taxaIndicador: taxaIndicador,
      detalheIndicador: detalheIndicador,
      competenciaAnteriorNaSerie: competenciaAnteriorNaSerie,
      linkPara: linkPara,
      corPorTaxaVsMeta: corPorTaxaVsMeta,
      posicaoRanking: posicaoRanking,
      calcularAlertasPlausibilidade: calcularAlertasPlausibilidade,
      regressaoLinear: regressaoLinear,
      projetarMeta: projetarMeta,
    };
  }
})();
