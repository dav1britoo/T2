const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Codeforces 449B - Jzzhu and Cities";

// ---------- Paleta "Midnight Executive" ----------
const NAVY = "1E2761";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const ACCENT = "F96167"; // coral
const GOLD = "F9E795";
const MUTED = "8A93B8";
const DARK = "0F1638";
const CODEBG = "12193C";
const SOFT = "F4F6FC";

const FONT_HEAD = "Georgia";
const FONT_BODY = "Calibri";
const FONT_MONO = "Consolas";

const TOTAL = 10;

// ---------- Helpers ----------
function footer(slide, n) {
  slide.addShape("rect", {
    x: 0, y: 7.15, w: 13.333, h: 0.35,
    fill: { color: NAVY }, line: { color: NAVY },
  });
  slide.addText("Grupo F • Codeforces 449B — Jzzhu and Cities", {
    x: 0.4, y: 7.18, w: 9, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, color: ICE,
  });
  slide.addText(`${n} / ${TOTAL}`, {
    x: 12.4, y: 7.18, w: 0.6, h: 0.3,
    fontFace: FONT_BODY, fontSize: 10, color: ICE, align: "right",
  });
}
function title(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 0.35, w: 12.3, h: 0.7,
    fontFace: FONT_HEAD, fontSize: 30, bold: true, color: NAVY,
  });
}
// Desenha um pequeno grafo do exemplo 1 dentro de um retângulo escuro
// distLabels: opcional {nodeId: "valor"} mostrado em pílula
// highlight: opcional set de nodeIds para destacar
function drawExemplo1(slide, x0, y0, w, h, opts = {}) {
  slide.addShape("roundRect", {
    x: x0, y: y0, w, h,
    fill: { color: DARK }, line: { color: NAVY, width: 1 }, rectRadius: 0.08,
  });
  // Posições relativas, escaladas para a caixa
  const padX = 0.5, padY = 0.4;
  const innerW = w - 2 * padX;
  const innerH = h - 2 * padY;
  // Layout dos nós em coords [0..1]
  const layout = {
    1: { rx: 0.10, ry: 0.50 },
    2: { rx: 0.45, ry: 0.18 },
    3: { rx: 0.62, ry: 0.50 },
    4: { rx: 0.92, ry: 0.78 },
    5: { rx: 0.45, ry: 0.85 },
  };
  const nodes = {};
  for (const k of Object.keys(layout)) {
    nodes[k] = {
      x: x0 + padX + layout[k].rx * innerW,
      y: y0 + padY + layout[k].ry * innerH,
    };
  }
  // Arestas (estradas)
  const edges = [[1,2,"1"],[2,3,"2"],[1,3,"3"],[3,4,"4"],[1,5,"5"]];
  for (const [a,b,w2] of edges) {
    const A = nodes[a], B = nodes[b];
    slide.addShape("line", {
      x: A.x, y: A.y, w: B.x - A.x, h: B.y - A.y,
      line: { color: ICE, width: 1.5 },
    });
    slide.addText(w2, {
      x: (A.x + B.x)/2 - 0.18, y: (A.y + B.y)/2 - 0.16, w: 0.36, h: 0.3,
      fontFace: FONT_MONO, fontSize: 11, color: GOLD, align: "center", bold: true,
    });
  }
  // Trens (linha tracejada)
  const trains = [[3,"5"],[4,"5"],[5,"5"]];
  for (const [t,] of trains) {
    const A = nodes[1], B = nodes[t];
    slide.addShape("line", {
      x: A.x, y: A.y, w: B.x - A.x, h: B.y - A.y,
      line: { color: ACCENT, width: 1.1, dashType: "dash" },
    });
  }
  // Nós
  for (const [id, p] of Object.entries(nodes)) {
    const isCap = id === "1";
    const highlighted = opts.highlight && opts.highlight.includes(id);
    slide.addShape("ellipse", {
      x: p.x - 0.22, y: p.y - 0.22, w: 0.44, h: 0.44,
      fill: { color: isCap ? GOLD : (highlighted ? ACCENT : ICE) },
      line: { color: NAVY, width: 2 },
    });
    slide.addText(id, {
      x: p.x - 0.22, y: p.y - 0.23, w: 0.44, h: 0.44,
      fontFace: FONT_HEAD, fontSize: 14, bold: true,
      color: highlighted && !isCap ? WHITE : NAVY,
      align: "center", valign: "middle",
    });
    // Etiqueta de distância
    if (opts.distLabels && opts.distLabels[id] !== undefined) {
      slide.addShape("roundRect", {
        x: p.x + 0.16, y: p.y - 0.45, w: 0.6, h: 0.32,
        fill: { color: NAVY }, line: { color: GOLD, width: 1 },
        rectRadius: 0.04,
      });
      slide.addText(String(opts.distLabels[id]), {
        x: p.x + 0.16, y: p.y - 0.47, w: 0.6, h: 0.32,
        fontFace: FONT_MONO, fontSize: 11, bold: true, color: GOLD,
        align: "center", valign: "middle",
      });
    }
  }
}

// =============================================================
// Slide 1 — Capa
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape("rect", { x: 0, y: 0, w: 0.4, h: 7.5, fill: { color: ACCENT }, line: { color: ACCENT } });
  s.addText("Resolução de Problemas com Grafos", {
    x: 0.9, y: 0.7, w: 11.5, h: 0.5,
    fontFace: FONT_BODY, fontSize: 16, color: ICE, italic: true,
  });
  s.addText("Codeforces 449B", {
    x: 0.9, y: 1.3, w: 11.5, h: 0.7,
    fontFace: FONT_HEAD, fontSize: 28, bold: true, color: GOLD,
  });
  s.addText("Jzzhu and Cities", {
    x: 0.9, y: 2.0, w: 11.5, h: 1.5,
    fontFace: FONT_HEAD, fontSize: 60, bold: true, color: WHITE,
  });
  s.addText("Dijkstra com análise de arestas redundantes", {
    x: 0.9, y: 3.5, w: 11.5, h: 0.5,
    fontFace: FONT_BODY, fontSize: 20, color: ICE,
  });
  s.addShape("roundRect", {
    x: 0.9, y: 4.6, w: 11.5, h: 2.0,
    fill: { color: DARK }, line: { color: ACCENT, width: 1 }, rectRadius: 0.1,
  });
  s.addText("Grupo F", {
    x: 1.2, y: 4.7, w: 5, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 22, bold: true, color: GOLD,
  });
  s.addText([
    { text: "Integrantes:  ", options: { bold: true, color: ICE } },
    { text: "João Vitor Silva  •  Antonio Davi  •  Pablo Dornelles", options: { color: WHITE } },
  ], { x: 1.2, y: 5.25, w: 11, h: 0.4, fontFace: FONT_BODY, fontSize: 16 });
  s.addText([
    { text: "Orientador:    ", options: { bold: true, color: ICE } },
    { text: "Prof. Me Ricardo Carubbi", options: { color: WHITE } },
  ], { x: 1.2, y: 5.7, w: 11, h: 0.4, fontFace: FONT_BODY, fontSize: 16 });
  s.addText([
    { text: "Linguagem:   ", options: { bold: true, color: ICE } },
    { text: "Python 3   •   Plataforma: Codeforces", options: { color: WHITE } },
  ], { x: 1.2, y: 6.15, w: 11, h: 0.4, fontFace: FONT_BODY, fontSize: 16 });
  s.addText("Trabalho Prático 2 — Unidade 3", {
    x: 0.9, y: 6.85, w: 11.5, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, color: MUTED, italic: true,
  });
}

// =============================================================
// Slide 2 — O problema
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "1. O problema");

  // Bloco esquerdo — enunciado + saída
  s.addShape("roundRect", {
    x: 0.5, y: 1.3, w: 6.2, h: 5.5,
    fill: { color: SOFT }, line: { color: ICE, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Enunciado", {
    x: 0.8, y: 1.45, w: 5.6, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: NAVY,
  });
  s.addText([
    { text: "n", options: { bold: true } },
    { text: " cidades, sendo a cidade " },
    { text: "1", options: { bold: true } },
    { text: " a capital.\n" },
    { text: "m", options: { bold: true } },
    { text: " estradas bidirecionais com peso " },
    { text: "x", options: { italic: true } },
    { text: ".\n" },
    { text: "k", options: { bold: true } },
    { text: " trens, cada um direto da capital a uma cidade " },
    { text: "s", options: { italic: true } },
    { text: " com custo " },
    { text: "y", options: { italic: true } },
    { text: "." },
  ], {
    x: 0.8, y: 1.95, w: 5.6, h: 1.6,
    fontFace: FONT_BODY, fontSize: 15, color: DARK,
  });

  s.addText("Objetivo", {
    x: 0.8, y: 3.55, w: 5.6, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: ACCENT,
  });
  s.addText(
    "Fechar o máximo possível de trens sem alterar a distância mínima da capital a nenhuma cidade.",
    {
      x: 0.8, y: 4.0, w: 5.6, h: 1.0,
      fontFace: FONT_BODY, fontSize: 15, color: DARK,
    });

  s.addText("Saída para o exemplo:", {
    x: 0.8, y: 5.2, w: 5.6, h: 0.35,
    fontFace: FONT_BODY, fontSize: 13, color: NAVY, bold: true,
  });
  s.addShape("roundRect", {
    x: 0.8, y: 5.55, w: 2.0, h: 0.7,
    fill: { color: NAVY }, line: { color: GOLD, width: 1 }, rectRadius: 0.05,
  });
  s.addText("2", {
    x: 0.8, y: 5.55, w: 2.0, h: 0.7,
    fontFace: FONT_HEAD, fontSize: 30, bold: true, color: GOLD,
    align: "center", valign: "middle",
  });
  s.addText("trens podem ser fechados", {
    x: 2.95, y: 5.65, w: 3.5, h: 0.5,
    fontFace: FONT_BODY, fontSize: 13, color: DARK, italic: true, valign: "middle",
  });

  s.addText("Limites:  n ≤ 10⁵   •   m ≤ 3·10⁵   •   k ≤ 10⁵   •   pesos até 10⁹", {
    x: 0.8, y: 6.4, w: 5.6, h: 0.4,
    fontFace: FONT_MONO, fontSize: 11, color: NAVY, italic: true,
  });

  // Bloco direito — diagrama
  drawExemplo1(s, 7.0, 1.3, 5.85, 5.5);
  s.addText("Exemplo 1", {
    x: 7.2, y: 1.45, w: 5.5, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });
  // Legenda
  s.addShape("line", { x: 7.25, y: 6.4, w: 0.6, h: 0, line: { color: ICE, width: 2 } });
  s.addText("estrada", { x: 7.95, y: 6.25, w: 1.3, h: 0.3, fontFace: FONT_BODY, fontSize: 11, color: ICE });
  s.addShape("line", { x: 9.3, y: 6.4, w: 0.6, h: 0, line: { color: ACCENT, width: 2, dashType: "dash" } });
  s.addText("trem (saindo da capital)", { x: 10.0, y: 6.25, w: 2.7, h: 0.3, fontFace: FONT_BODY, fontSize: 11, color: ACCENT });

  footer(s, 2);
}

// =============================================================
// Slide 3 — Construção do grafo (enunciado → grafo)
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "2. Do enunciado ao grafo");

  s.addText("Cada linha da entrada vira um elemento do grafo. A imagem cresce em três etapas:", {
    x: 0.5, y: 1.05, w: 12.3, h: 0.35,
    fontFace: FONT_BODY, fontSize: 14, italic: true, color: MUTED,
  });

  // Três painéis lado a lado
  const stages = [
    {
      titulo: "1. Cidades viram vértices",
      cor: ICE,
      // só pontos
      drawNodes: true, drawRoads: false, drawTrains: false,
      desc: "n nós no plano — cidade 1 é a capital (origem).",
    },
    {
      titulo: "2. Estradas viram arestas",
      cor: GOLD,
      drawNodes: true, drawRoads: true, drawTrains: false,
      desc: "Cada estrada (u, v, x) é uma aresta bidirecional de peso x.",
    },
    {
      titulo: "3. Trens viram arestas extras",
      cor: ACCENT,
      drawNodes: true, drawRoads: true, drawTrains: true,
      desc: "Cada trem (s, y) é uma aresta da capital 1 para s, marcada como trem.",
    },
  ];

  let cx = 0.5;
  for (const st of stages) {
    // cartão
    s.addShape("roundRect", {
      x: cx, y: 1.55, w: 4.1, h: 5.3,
      fill: { color: DARK }, line: { color: st.cor, width: 2 }, rectRadius: 0.08,
    });
    s.addText(st.titulo, {
      x: cx + 0.2, y: 1.65, w: 3.7, h: 0.45,
      fontFace: FONT_HEAD, fontSize: 15, bold: true, color: st.cor,
    });
    // mini-grafo dentro do cartão
    const gx = cx + 0.2, gy = 2.2, gw = 3.7, gh = 3.2;
    // Layout reduzido
    const lay = {
      1: { rx: 0.12, ry: 0.55 },
      2: { rx: 0.42, ry: 0.10 },
      3: { rx: 0.60, ry: 0.55 },
      4: { rx: 0.90, ry: 0.92 },
      5: { rx: 0.42, ry: 0.92 },
    };
    const np = {};
    for (const k of Object.keys(lay)) {
      np[k] = { x: gx + lay[k].rx * gw, y: gy + lay[k].ry * gh };
    }
    // estradas
    if (st.drawRoads) {
      for (const [a,b] of [[1,2],[2,3],[1,3],[3,4],[1,5]]) {
        s.addShape("line", {
          x: np[a].x, y: np[a].y, w: np[b].x - np[a].x, h: np[b].y - np[a].y,
          line: { color: ICE, width: 1.2 },
        });
      }
    }
    // trens
    if (st.drawTrains) {
      for (const t of [3,4,5]) {
        s.addShape("line", {
          x: np[1].x, y: np[1].y, w: np[t].x - np[1].x, h: np[t].y - np[1].y,
          line: { color: ACCENT, width: 1.0, dashType: "dash" },
        });
      }
    }
    // nós
    for (const [id, p] of Object.entries(np)) {
      const isCap = id === "1";
      s.addShape("ellipse", {
        x: p.x - 0.18, y: p.y - 0.18, w: 0.36, h: 0.36,
        fill: { color: isCap ? GOLD : ICE },
        line: { color: NAVY, width: 1.5 },
      });
      s.addText(id, {
        x: p.x - 0.18, y: p.y - 0.19, w: 0.36, h: 0.36,
        fontFace: FONT_HEAD, fontSize: 11, bold: true, color: NAVY,
        align: "center", valign: "middle",
      });
    }
    s.addText(st.desc, {
      x: cx + 0.2, y: 5.6, w: 3.7, h: 1.15,
      fontFace: FONT_BODY, fontSize: 12, color: ICE,
    });
    cx += 4.27;
  }

  footer(s, 3);
}

// =============================================================
// Slide 4 — Modelagem (vértices, arestas, pesos, Dijkstra ok)
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "3. Modelagem como grafo");

  // Tabela de correspondência
  s.addText("Correspondência enunciado ↔ grafo", {
    x: 0.5, y: 1.15, w: 7.0, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: NAVY,
  });

  const rows = [
    ["Enunciado", "Grafo G(V, E)"],
    ["n cidades, capital = 1", "n vértices, origem v = 1"],
    ["Estrada (u, v, x) bidirecional", "Aresta (u, v) com peso x — tipo: estrada"],
    ["Trem (s, y) direto da capital", "Aresta (1, s) com peso y — tipo: trem"],
    ["Distância mínima até cada cidade", "Dijkstra de origem única a partir de 1"],
  ];
  // Cabeçalho
  s.addTable(rows, {
    x: 0.5, y: 1.6, w: 7.0, colW: [2.8, 4.2],
    fontSize: 13, fontFace: FONT_BODY,
    color: DARK, fill: { color: WHITE },
    border: { type: "solid", pt: 1, color: ICE },
    rowH: 0.55,
  });

  // Painel à direita: por que Dijkstra
  s.addShape("roundRect", {
    x: 7.8, y: 1.15, w: 5.05, h: 3.4,
    fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText("Por que Dijkstra é aplicável?", {
    x: 8.0, y: 1.3, w: 4.75, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: GOLD,
  });
  s.addText([
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Pesos não negativos:  ", options: { color: ICE, bold: true } },
    { text: "x, y ≥ 1.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Caminho mais curto único a partir da capital:  ", options: { color: ICE, bold: true } },
    { text: "origem única para todos os vértices.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Trens como arestas:  ", options: { color: ICE, bold: true } },
    { text: "naturalmente integrados ao grafo, sem mudar o algoritmo.", options: { color: WHITE } },
  ], {
    x: 8.0, y: 1.85, w: 4.75, h: 2.6,
    fontFace: FONT_BODY, fontSize: 13, paraSpaceAfter: 4,
  });

  // Painel inferior — representação
  s.addShape("roundRect", {
    x: 0.5, y: 5.0, w: 12.35, h: 1.95,
    fill: { color: CODEBG }, line: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText("Representação: lista de adjacência", {
    x: 0.7, y: 5.1, w: 11, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 15, bold: true, color: GOLD,
  });
  s.addText([
    { text: "adj[u] = [ (v, peso, é_trem), ... ]\n\n", options: { color: WHITE, bold: true } },
    { text: "• estradas:  ", options: { color: ICE } },
    { text: "adj[u].append((v, x, False))   e   adj[v].append((u, x, False))\n", options: { color: WHITE } },
    { text: "• trens:       ", options: { color: ICE } },
    { text: "adj[1].append((s, y, True))     (só saindo da capital)", options: { color: WHITE } },
  ], {
    x: 0.7, y: 5.5, w: 12, h: 1.45,
    fontFace: FONT_MONO, fontSize: 13,
  });

  // Painel direita-inferior: tamanhos
  s.addShape("roundRect", {
    x: 7.8, y: 4.65, w: 5.05, h: 0.35,
    fill: { color: SOFT }, line: { color: ICE, width: 1 }, rectRadius: 0.04,
  });
  s.addText("V = n, E = 2m + k", {
    x: 7.85, y: 4.65, w: 5.0, h: 0.35,
    fontFace: FONT_MONO, fontSize: 12, color: NAVY, align: "center", valign: "middle", bold: true,
  });

  footer(s, 4);
}

// =============================================================
// Slide 5 — Estratégia: Dijkstra + via_estrada
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "4. Estratégia: Dijkstra + marcação de alternativa");

  s.addText([
    { text: "Variação de Dijkstra:  ", options: { bold: true, color: ACCENT } },
    { text: "uma única execução do Dijkstra clássico de origem única, com uma anotação extra por vértice (", options: { color: MUTED, italic: true } },
    { text: "via_estrada", options: { color: NAVY, italic: true, bold: true } },
    { text: ") — sem expansão de estado nem múltiplas execuções.", options: { color: MUTED, italic: true } },
  ], {
    x: 0.5, y: 1.05, w: 12.3, h: 0.35,
    fontFace: FONT_BODY, fontSize: 13,
  });

  // Três blocos: Init, Fila de prioridade, Relaxamento
  const cards = [
    {
      h: "Inicialização",
      lines: [
        "dist[1] = 0",
        "dist[v] = ∞  ∀v ≠ 1",
        "via_estrada[1] = True",
        "via_estrada[v] = False",
        "pq = [(0, 1)]",
      ],
      cor: GOLD,
    },
    {
      h: "Fila de prioridade",
      lines: [
        "heapq (heap binário mín.)",
        "extrai sempre o vértice",
        "de menor distância.",
        "Descarta entradas obsoletas",
        "via   if d > dist[u]: continue",
      ],
      cor: ACCENT,
    },
    {
      h: "Critério de parada",
      lines: [
        "fila vazia.",
        "Ao final, dist[] tem o caminho",
        "mínimo de 1 até cada vértice,",
        "incluindo trens como arestas.",
        "",
      ],
      cor: ICE,
    },
  ];

  let cx = 0.5;
  for (const c of cards) {
    s.addShape("roundRect", {
      x: cx, y: 1.55, w: 4.1, h: 2.65,
      fill: { color: DARK }, line: { color: c.cor, width: 2 }, rectRadius: 0.08,
    });
    s.addText(c.h, {
      x: cx + 0.25, y: 1.65, w: 3.65, h: 0.4,
      fontFace: FONT_HEAD, fontSize: 16, bold: true, color: c.cor,
    });
    s.addText(c.lines.map((t, i) => ({
      text: t + (i < c.lines.length - 1 ? "\n" : ""),
      options: { color: i === 0 ? WHITE : ICE, bold: i === 0 },
    })), {
      x: cx + 0.25, y: 2.1, w: 3.65, h: 2.0,
      fontFace: FONT_MONO, fontSize: 12, paraSpaceAfter: 2,
    });
    cx += 4.27;
  }

  // Bloco grande: relaxamento adaptado
  s.addShape("roundRect", {
    x: 0.5, y: 4.35, w: 12.35, h: 2.6,
    fill: { color: CODEBG }, line: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText("Relaxamento da aresta  u → v  (peso w, é_trem t)", {
    x: 0.7, y: 4.45, w: 12, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 15, bold: true, color: GOLD,
  });
  s.addText([
    { text: "nd = dist[u] + w\n\n", options: { color: WHITE } },
    { text: "se  nd < dist[v]:                              ", options: { color: WHITE, bold: true } },
    { text: "# encontrou caminho mais curto\n", options: { color: MUTED, italic: true } },
    { text: "    dist[v] = nd;  via_estrada[v] = (not t);  push(pq, (nd, v))\n\n", options: { color: WHITE } },
    { text: "senão se  nd == dist[v]  e  not t:    ", options: { color: WHITE, bold: true } },
    { text: "# alternativa de mesmo custo por estrada\n", options: { color: MUTED, italic: true } },
    { text: "    via_estrada[v] = True", options: { color: GOLD, bold: true } },
  ], {
    x: 0.7, y: 4.85, w: 12, h: 2.05,
    fontFace: FONT_MONO, fontSize: 13,
  });

  footer(s, 5);
}

// =============================================================
// Slide 6 — Walkthrough do exemplo 1
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "5. Dijkstra em ação — Exemplo 1");

  // Diagrama com distâncias finais
  drawExemplo1(s, 0.5, 1.3, 6.0, 5.5, {
    distLabels: { 1: "0", 2: "1", 3: "3", 4: "5", 5: "5" },
  });
  s.addText("Grafo + dist finais", {
    x: 0.7, y: 1.45, w: 5.5, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: GOLD,
  });

  // Tabela de iterações (extrações da fila)
  s.addText("Iterações (extração da fila)", {
    x: 6.7, y: 1.3, w: 6.2, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: NAVY,
  });
  const itRows = [
    ["pop", "d", "atualizações"],
    ["1", "0", "dist[2]=1  dist[3]=3  dist[5]=5  dist[4]=5*"],
    ["2", "1", "—  (nd=3 == dist[3], via estrada → via_estrada[3]=T)"],
    ["3", "3", "—  (3→4 daria 7 > 5)"],
    ["4", "5", "—"],
    ["5", "5", "—"],
  ];
  s.addTable(itRows, {
    x: 6.7, y: 1.75, w: 6.2, colW: [0.7, 0.7, 4.8],
    fontSize: 12, fontFace: FONT_MONO,
    color: DARK,
    border: { type: "solid", pt: 1, color: ICE },
    rowH: 0.4,
  });
  s.addText("*  dist[4]=5 via trem (1→4, y=5). via_estrada[4] = False.", {
    x: 6.7, y: 4.35, w: 6.2, h: 0.35,
    fontFace: FONT_BODY, fontSize: 11, italic: true, color: MUTED,
  });

  // Bloco de resultado final
  s.addShape("roundRect", {
    x: 6.7, y: 4.85, w: 6.2, h: 2.0,
    fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText("Estado final dos vetores", {
    x: 6.9, y: 4.95, w: 5.8, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 14, bold: true, color: GOLD,
  });
  s.addText([
    { text: "dist          = [_, 0, 1, 3, 5, 5]\n", options: { color: WHITE } },
    { text: "via_estrada = [_, T, T, T, ", options: { color: WHITE } },
    { text: "F", options: { color: ACCENT, bold: true } },
    { text: ", T]\n\n", options: { color: WHITE } },
    { text: "→ apenas a cidade 4 depende de trem", options: { color: GOLD, italic: true } },
  ], {
    x: 6.9, y: 5.4, w: 5.8, h: 1.4,
    fontFace: FONT_MONO, fontSize: 12,
  });

  footer(s, 6);
}

// =============================================================
// Slide 7 — Regras de fechamento dos trens
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "6. Regras de fechamento dos trens");

  // Observação-chave no topo
  s.addShape("roundRect", {
    x: 0.5, y: 1.15, w: 12.35, h: 1.0,
    fill: { color: SOFT }, line: { color: ICE, width: 1 }, rectRadius: 0.06,
  });
  s.addText([
    { text: "Observação-chave:  ", options: { bold: true, color: ACCENT } },
    { text: "todo trem sai da capital, então em qualquer caminho mínimo um trem só pode aparecer como ", options: { color: DARK } },
    { text: "a primeira (e única) aresta", options: { bold: true, color: NAVY } },
    { text: ". Para vértices que não são destino direto de trem o problema é Dijkstra padrão — basta decidir, para cada trem, se ele pode ser fechado.", options: { color: DARK } },
  ], {
    x: 0.7, y: 1.25, w: 12, h: 0.85,
    fontFace: FONT_BODY, fontSize: 13,
  });

  // Quatro casos
  const cases = [
    { titulo: "Trem inútil",             cond: "y > dist[s]",                        res: "FECHA",     txt: "O melhor caminho até s já é menor que o trem.", cor: ACCENT },
    { titulo: "Alternativa por estrada", cond: "y = dist[s] e via_estrada[s]",       res: "FECHA",     txt: "Há um caminho ótimo só por estradas — trem redundante.", cor: ICE },
    { titulo: "Trem essencial",          cond: "y = dist[s] e ¬via_estrada[s]",      res: "MANTÉM 1",  txt: "Sem trem, dist[s] aumentaria. Duplicatas são fechadas.", cor: GOLD },
    { titulo: "Caso impossível",         cond: "y < dist[s]",                        res: "—",         txt: "Não ocorre: o trem já entrou no Dijkstra e fixaria dist[s].", cor: MUTED },
  ];
  let cx = 0.5;
  const cardW = 3.02;
  for (const c of cases) {
    s.addShape("roundRect", {
      x: cx, y: 2.45, w: cardW, h: 3.0,
      fill: { color: DARK }, line: { color: c.cor, width: 2 }, rectRadius: 0.08,
    });
    s.addText(c.titulo, {
      x: cx + 0.2, y: 2.55, w: cardW - 0.4, h: 0.4,
      fontFace: FONT_HEAD, fontSize: 14, bold: true, color: c.cor,
    });
    s.addText(c.cond, {
      x: cx + 0.2, y: 3.0, w: cardW - 0.4, h: 0.4,
      fontFace: FONT_MONO, fontSize: 11, color: WHITE,
    });
    s.addShape("line", { x: cx + 0.2, y: 3.5, w: cardW - 0.4, h: 0, line: { color: c.cor, width: 1 } });
    s.addText("→ " + c.res, {
      x: cx + 0.2, y: 3.6, w: cardW - 0.4, h: 0.4,
      fontFace: FONT_BODY, fontSize: 12, bold: true, color: c.cor,
    });
    s.addText(c.txt, {
      x: cx + 0.2, y: 4.05, w: cardW - 0.4, h: 1.35,
      fontFace: FONT_BODY, fontSize: 11, color: ICE,
    });
    cx += cardW + 0.08;
  }

  // Aplicação aos trens do exemplo
  s.addShape("roundRect", {
    x: 0.5, y: 5.6, w: 12.35, h: 1.4,
    fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.06,
  });
  s.addText("Aplicação no exemplo 1", {
    x: 0.7, y: 5.7, w: 12, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 14, bold: true, color: GOLD,
  });
  s.addText([
    { text: "trem (3, 5):  ", options: { color: ICE, bold: true } },
    { text: "y=5 > dist[3]=3            → ", options: { color: WHITE } },
    { text: "fecha\n", options: { color: ACCENT, bold: true } },
    { text: "trem (4, 5):  ", options: { color: ICE, bold: true } },
    { text: "y=5 = dist[4],  ¬via_estrada[4]  → ", options: { color: WHITE } },
    { text: "mantém\n", options: { color: GOLD, bold: true } },
    { text: "trem (5, 5):  ", options: { color: ICE, bold: true } },
    { text: "y=5 = dist[5],  via_estrada[5]   → ", options: { color: WHITE } },
    { text: "fecha", options: { color: ACCENT, bold: true } },
  ], {
    x: 0.7, y: 6.1, w: 12, h: 0.9,
    fontFace: FONT_MONO, fontSize: 12,
  });

  footer(s, 7);
}

// =============================================================
// Slide 8 — Complexidade
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "7. Complexidade");

  // Bloco principal — tempo
  s.addShape("roundRect", {
    x: 0.5, y: 1.3, w: 6.2, h: 5.5,
    fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.1,
  });
  s.addText("Tempo", {
    x: 0.8, y: 1.45, w: 5.6, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });
  s.addText("O((V + E) · log V)", {
    x: 0.8, y: 2.0, w: 5.6, h: 1.0,
    fontFace: FONT_MONO, fontSize: 32, bold: true, color: WHITE,
  });
  s.addText([
    { text: "V = n           ", options: { color: ICE, bold: true } },
    { text: "(vértices)\n", options: { color: MUTED, italic: true } },
    { text: "E = 2m + k    ", options: { color: ICE, bold: true } },
    { text: "(estradas duplicadas + trens)\n", options: { color: MUTED, italic: true } },
    { text: "log V          ", options: { color: ICE, bold: true } },
    { text: "(heap binário, heapq)\n\n", options: { color: MUTED, italic: true } },
    { text: "+ O(k)          ", options: { color: ICE, bold: true } },
    { text: "contagem final sobre os trens", options: { color: MUTED, italic: true } },
  ], {
    x: 0.8, y: 3.05, w: 5.6, h: 2.4,
    fontFace: FONT_MONO, fontSize: 13,
  });
  s.addText("Cada vértice entra na fila no máximo uma vez por relaxamento bem-sucedido; cada aresta é processada O(1) vez na ramificação ativa.", {
    x: 0.8, y: 5.7, w: 5.6, h: 1.0,
    fontFace: FONT_BODY, fontSize: 12, italic: true, color: ICE,
  });

  // Memória
  s.addShape("roundRect", {
    x: 7.0, y: 1.3, w: 5.85, h: 2.5,
    fill: { color: SOFT }, line: { color: ICE, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Memória", {
    x: 7.2, y: 1.45, w: 5.5, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: ACCENT,
  });
  s.addText("O(V + E)", {
    x: 7.2, y: 1.95, w: 5.5, h: 0.6,
    fontFace: FONT_MONO, fontSize: 24, bold: true, color: NAVY,
  });
  s.addText([
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "lista de adjacência ", options: { color: DARK, bold: true } },
    { text: "≈ 2m + k entradas\n", options: { color: DARK } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "vetores ", options: { color: DARK, bold: true } },
    { text: "dist e via_estrada de tamanho n\n", options: { color: DARK } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "heap ", options: { color: DARK, bold: true } },
    { text: "com no máximo O(E) entradas vivas", options: { color: DARK } },
  ], {
    x: 7.2, y: 2.65, w: 5.5, h: 1.15,
    fontFace: FONT_BODY, fontSize: 12,
  });

  // No problema (limites)
  s.addShape("roundRect", {
    x: 7.0, y: 3.95, w: 5.85, h: 2.85,
    fill: { color: DARK }, line: { color: ACCENT, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Nos limites do problema", {
    x: 7.2, y: 4.05, w: 5.5, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: GOLD,
  });
  s.addText([
    { text: "n ≤ 10⁵    m ≤ 3·10⁵    k ≤ 10⁵\n", options: { color: WHITE, bold: true } },
    { text: "⇒ ~7·10⁵ arestas e ~6·10⁶ ops de heap.\n\n", options: { color: ICE } },
    { text: "Submissão real:\n", options: { color: ICE, bold: true } },
    { text: "tempo:    ", options: { color: ICE } },
    { text: "1625 ms", options: { color: GOLD, bold: true } },
    { text: "   (limite 2000 ms)\n", options: { color: ICE } },
    { text: "memória:  ", options: { color: ICE } },
    { text: "180 MB", options: { color: GOLD, bold: true } },
    { text: "      (limite 256 MB)", options: { color: ICE } },
  ], {
    x: 7.2, y: 4.45, w: 5.5, h: 2.3,
    fontFace: FONT_MONO, fontSize: 12,
  });

  footer(s, 8);
}

// =============================================================
// Slide 9 — Casos especiais
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "8. Casos especiais");

  const items = [
    { h: "Pesos grandes (até 10⁹)",                    t: "Distâncias podem chegar a ~10¹⁴. Em Python inteiros são arbitrários — sem overflow.\nEm Java, usar long." },
    { h: "Múltiplos trens para a mesma cidade",         t: "Agrupamos por destino. No máximo um é mantido quando estritamente necessário; os demais empatados são fechados." },
    { h: "Arestas paralelas entre o mesmo par",         t: "Tratadas naturalmente — cada uma vira uma entrada própria na lista de adjacência. O Dijkstra escolhe a melhor." },
    { h: "Empate trem × estrada (y = dist[s])",         t: "A marcação via_estrada[s] detecta a alternativa por estrada e permite fechar o trem redundante." },
    { h: "Trem que melhora a distância (y < dist[s])",  t: "Tratado dentro do próprio Dijkstra — o trem é apenas mais uma aresta. dist[s] cai e via_estrada[s] fica False." },
    { h: "Conectividade garantida",                     t: "O enunciado garante que toda cidade alcança a capital, então nenhum dist[v] fica em ∞ no final." },
  ];

  let i = 0;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 3; col++) {
      const x = 0.5 + col * 4.27;
      const y = 1.3 + row * 2.85;
      const it = items[i++];
      s.addShape("roundRect", {
        x, y, w: 4.06, h: 2.6,
        fill: { color: SOFT }, line: { color: NAVY, width: 1 }, rectRadius: 0.06,
      });
      s.addShape("rect", {
        x, y, w: 0.12, h: 2.6,
        fill: { color: ACCENT }, line: { color: ACCENT },
      });
      s.addText(it.h, {
        x: x + 0.3, y: y + 0.2, w: 3.7, h: 0.7,
        fontFace: FONT_HEAD, fontSize: 14, bold: true, color: NAVY,
      });
      s.addText(it.t, {
        x: x + 0.3, y: y + 0.95, w: 3.7, h: 1.55,
        fontFace: FONT_BODY, fontSize: 12, color: DARK,
      });
    }
  }

  footer(s, 9);
}

// =============================================================
// Slide 10 — Conclusão / Accepted
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape("rect", { x: 0, y: 0, w: 0.4, h: 7.5, fill: { color: ACCENT }, line: { color: ACCENT } });

  s.addText("Resultado", {
    x: 0.9, y: 0.5, w: 11.5, h: 0.5,
    fontFace: FONT_BODY, fontSize: 16, italic: true, color: ICE,
  });
  s.addText("Accepted", {
    x: 0.9, y: 1.05, w: 11.5, h: 1.3,
    fontFace: FONT_HEAD, fontSize: 80, bold: true, color: GOLD,
  });

  const stats = [
    { v: "#376408159", l: "ID da submissão" },
    { v: "1625 ms",     l: "tempo (limite 2 s)" },
    { v: "180 MB",      l: "memória (limite 256 MB)" },
  ];
  let x = 0.9;
  for (const st of stats) {
    s.addShape("roundRect", {
      x, y: 2.7, w: 3.9, h: 1.5,
      fill: { color: DARK }, line: { color: ACCENT, width: 1 }, rectRadius: 0.08,
    });
    s.addText(st.v, {
      x: x + 0.2, y: 2.85, w: 3.5, h: 0.7,
      fontFace: FONT_HEAD, fontSize: 24, bold: true, color: WHITE,
    });
    s.addText(st.l, {
      x: x + 0.2, y: 3.55, w: 3.5, h: 0.55,
      fontFace: FONT_BODY, fontSize: 12, color: ICE, italic: true,
    });
    x += 4.0;
  }

  s.addShape("roundRect", {
    x: 0.9, y: 4.45, w: 11.5, h: 2.35,
    fill: { color: DARK }, line: { color: GOLD, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Em resumo", {
    x: 1.1, y: 4.55, w: 11.1, h: 0.45,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });
  s.addText([
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Modelagem unificada: trens são apenas arestas extras a partir da capital.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Uma única execução de Dijkstra com a marcação via_estrada já resolve.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Contagem direta sobre os trens, agrupados por destino, fornece a resposta.", options: { color: WHITE } },
  ], {
    x: 1.1, y: 5.1, w: 11.1, h: 1.6,
    fontFace: FONT_BODY, fontSize: 15, paraSpaceAfter: 3,
  });

  s.addText("Obrigado!  •  Grupo F  •  Codeforces 449B — Jzzhu and Cities", {
    x: 0.9, y: 6.95, w: 11.5, h: 0.4,
    fontFace: FONT_BODY, fontSize: 12, italic: true, color: ICE, align: "center",
  });
}

pres.writeFile({ fileName: "apresentacao.pptx" }).then((f) => {
  console.log("escrito:", f);
});
