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

const TOTAL = 8;

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
// Nó circular reutilizável
function node(slide, cx, cy, label, opts = {}) {
  const r = opts.r || 0.24;
  slide.addShape("ellipse", {
    x: cx - r, y: cy - r, w: 2 * r, h: 2 * r,
    fill: { color: opts.fill || ICE },
    line: { color: opts.line || NAVY, width: opts.lw || 2 },
  });
  slide.addText(label, {
    x: cx - r, y: cy - r - 0.01, w: 2 * r, h: 2 * r,
    fontFace: FONT_HEAD, fontSize: opts.fs || 14, bold: true,
    color: opts.text || NAVY, align: "center", valign: "middle",
  });
}
// Aresta com rótulo opcional
function edge(slide, a, b, opts = {}) {
  slide.addShape("line", {
    x: a.x, y: a.y, w: b.x - a.x, h: b.y - a.y,
    line: {
      color: opts.color || ICE,
      width: opts.width || 1.5,
      dashType: opts.dash ? "dash" : "solid",
    },
  });
  if (opts.label !== undefined) {
    slide.addText(String(opts.label), {
      x: (a.x + b.x) / 2 - 0.2 + (opts.lx || 0),
      y: (a.y + b.y) / 2 - 0.16 + (opts.ly || 0),
      w: 0.4, h: 0.3,
      fontFace: FONT_MONO, fontSize: 11, bold: true,
      color: opts.labelColor || GOLD, align: "center",
    });
  }
}
// Grafo do exemplo 1
function drawExemplo1(slide, x0, y0, w, h, opts = {}) {
  slide.addShape("roundRect", {
    x: x0, y: y0, w, h,
    fill: { color: DARK }, line: { color: NAVY, width: 1 }, rectRadius: 0.08,
  });
  const padX = 0.5, padY = 0.4;
  const iw = w - 2 * padX, ih = h - 2 * padY;
  const layout = {
    1: { rx: 0.10, ry: 0.50 }, 2: { rx: 0.45, ry: 0.18 },
    3: { rx: 0.62, ry: 0.50 }, 4: { rx: 0.92, ry: 0.78 },
    5: { rx: 0.45, ry: 0.85 },
  };
  const np = {};
  for (const k of Object.keys(layout))
    np[k] = { x: x0 + padX + layout[k].rx * iw, y: y0 + padY + layout[k].ry * ih };
  for (const [a, b, lbl] of [[1,2,"1"],[2,3,"2"],[1,3,"3"],[3,4,"4"],[1,5,"5"]])
    edge(slide, np[a], np[b], { label: lbl });
  for (const t of [3, 4, 5])
    edge(slide, np[1], np[t], { color: ACCENT, width: 1.1, dash: true });
  for (const [id, p] of Object.entries(np)) {
    const isCap = id === "1";
    const hl = opts.highlight && opts.highlight.includes(id);
    node(slide, p.x, p.y, id, {
      r: 0.22, fill: isCap ? GOLD : (hl ? ACCENT : ICE),
      text: hl && !isCap ? WHITE : NAVY,
    });
    if (opts.distLabels && opts.distLabels[id] !== undefined) {
      slide.addShape("roundRect", {
        x: p.x + 0.16, y: p.y - 0.45, w: 0.6, h: 0.32,
        fill: { color: NAVY }, line: { color: GOLD, width: 1 }, rectRadius: 0.04,
      });
      slide.addText(String(opts.distLabels[id]), {
        x: p.x + 0.16, y: p.y - 0.47, w: 0.6, h: 0.32,
        fontFace: FONT_MONO, fontSize: 11, bold: true, color: GOLD,
        align: "center", valign: "middle",
      });
    }
  }
}
// Ícone em círculo colorido (motif recorrente)
function iconBadge(slide, cx, cy, glyph, color, r = 0.42) {
  slide.addShape("ellipse", {
    x: cx - r, y: cy - r, w: 2 * r, h: 2 * r,
    fill: { color }, line: { color, width: 1 },
  });
  slide.addText(glyph, {
    x: cx - r, y: cy - r, w: 2 * r, h: 2 * r,
    fontFace: FONT_HEAD, fontSize: r * 46, bold: true,
    color: NAVY, align: "center", valign: "middle",
  });
}

// =============================================================
// Slide 1 — Capa
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape("rect", { x: 0, y: 0, w: 0.4, h: 7.5, fill: { color: ACCENT }, line: { color: ACCENT } });
  s.addText("Resolução de Problemas com Grafos", {
    x: 0.9, y: 0.7, w: 11.5, h: 0.5, fontFace: FONT_BODY, fontSize: 16, color: ICE, italic: true,
  });
  s.addText("Codeforces 449B", {
    x: 0.9, y: 1.3, w: 11.5, h: 0.7, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: GOLD,
  });
  s.addText("Jzzhu and Cities", {
    x: 0.9, y: 2.0, w: 11.5, h: 1.5, fontFace: FONT_HEAD, fontSize: 60, bold: true, color: WHITE,
  });
  s.addText("Dijkstra com análise de arestas redundantes", {
    x: 0.9, y: 3.5, w: 11.5, h: 0.5, fontFace: FONT_BODY, fontSize: 20, color: ICE,
  });
  s.addShape("roundRect", {
    x: 0.9, y: 4.6, w: 11.5, h: 2.0, fill: { color: DARK }, line: { color: ACCENT, width: 1 }, rectRadius: 0.1,
  });
  s.addText("Grupo F", {
    x: 1.2, y: 4.7, w: 5, h: 0.5, fontFace: FONT_HEAD, fontSize: 22, bold: true, color: GOLD,
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
    x: 0.9, y: 6.85, w: 11.5, h: 0.4, fontFace: FONT_BODY, fontSize: 12, color: MUTED, italic: true,
  });
  s.addNotes(
    "[~10 s] Apresentação de 5 minutos. Abertura rápida: problema 449B do Codeforces, " +
    "resolvido com Dijkstra. Divisão: 1 min problema/modelagem, 2 min estratégia, " +
    "1 min complexidade/casos, 1 min conclusão."
  );
}

// =============================================================
// Slide 2 — O problema (ícones + grafo)
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "1. O problema");

  // Três ícones-resumo da entrada
  const facts = [
    { glyph: "⬢", color: ICE,    n: "n", t: "cidades\n(capital = 1)" },
    { glyph: "—", color: GOLD,   n: "m", t: "estradas\nbidirecionais" },
    { glyph: "⇢", color: ACCENT, n: "k", t: "trens\nda capital" },
  ];
  let fy = 1.5;
  for (const f of facts) {
    iconBadge(s, 1.1, fy + 0.4, f.glyph, f.color, 0.42);
    s.addText(f.n, {
      x: 1.7, y: fy, w: 1.0, h: 0.8, fontFace: FONT_HEAD, fontSize: 40, bold: true,
      color: NAVY, valign: "middle",
    });
    s.addText(f.t, {
      x: 2.7, y: fy, w: 3.6, h: 0.8, fontFace: FONT_BODY, fontSize: 15, color: DARK, valign: "middle",
    });
    fy += 1.0;
  }

  // Objetivo curto
  s.addShape("roundRect", {
    x: 0.6, y: 4.6, w: 5.7, h: 1.05, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText([
    { text: "Objetivo:  ", options: { bold: true, color: GOLD } },
    { text: "fechar o máximo de trens sem aumentar a distância da capital a nenhuma cidade.", options: { color: WHITE } },
  ], { x: 0.85, y: 4.7, w: 5.2, h: 0.85, fontFace: FONT_BODY, fontSize: 15, valign: "middle" });

  // Resposta do exemplo
  iconBadge(s, 1.05, 6.35, "2", GOLD, 0.42);
  s.addText("trens podem ser fechados  (saída do exemplo 1)", {
    x: 1.65, y: 6.05, w: 4.7, h: 0.6, fontFace: FONT_BODY, fontSize: 14, color: DARK, italic: true, valign: "middle",
  });

  // Grafo à direita
  drawExemplo1(s, 7.0, 1.3, 5.85, 5.5);
  s.addText("Exemplo 1", {
    x: 7.2, y: 1.45, w: 5.5, h: 0.4, fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });
  edge(s, { x: 7.25, y: 6.4 }, { x: 7.85, y: 6.4 }, { color: ICE, width: 2 });
  s.addText("estrada", { x: 7.95, y: 6.25, w: 1.3, h: 0.3, fontFace: FONT_BODY, fontSize: 11, color: ICE });
  edge(s, { x: 9.3, y: 6.4 }, { x: 9.9, y: 6.4 }, { color: ACCENT, width: 2, dash: true });
  s.addText("trem (da capital)", { x: 10.0, y: 6.25, w: 2.7, h: 0.3, fontFace: FONT_BODY, fontSize: 11, color: ACCENT });

  s.addNotes(
    "[~1 min p/ problema + modelagem — começa aqui]\n" +
    "n cidades (1 = capital), m estradas com peso, k trens diretos da capital. " +
    "Queremos fechar o MÁXIMO de trens sem aumentar a distância mínima de nenhuma cidade " +
    "até a capital. No exemplo, a resposta é 2."
  );
  footer(s, 2);
}

// =============================================================
// Slide 3 — Modelagem: enunciado → grafo + por que Dijkstra
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "2. Modelagem como grafo");
  s.addText("Cada linha da entrada vira um elemento do grafo — em três etapas:", {
    x: 0.5, y: 1.05, w: 12.3, h: 0.35, fontFace: FONT_BODY, fontSize: 14, italic: true, color: MUTED,
  });

  // Três mini-grafos das etapas (esquerda)
  const stages = [
    { titulo: "Cidades → vértices", cor: ICE, roads: false, trains: false },
    { titulo: "Estradas → arestas", cor: GOLD, roads: true, trains: false },
    { titulo: "Trens → arestas extras", cor: ACCENT, roads: true, trains: true },
  ];
  let cx = 0.5;
  for (const st of stages) {
    s.addShape("roundRect", {
      x: cx, y: 1.5, w: 2.5, h: 3.95, fill: { color: DARK }, line: { color: st.cor, width: 2 }, rectRadius: 0.08,
    });
    s.addText(st.titulo, {
      x: cx + 0.12, y: 1.58, w: 2.26, h: 0.6, fontFace: FONT_HEAD, fontSize: 12.5, bold: true, color: st.cor,
    });
    const gx = cx + 0.18, gy = 2.35, gw = 2.14, gh = 2.7;
    const lay = {
      1: { rx: 0.12, ry: 0.50 }, 2: { rx: 0.48, ry: 0.08 }, 3: { rx: 0.64, ry: 0.50 },
      4: { rx: 0.92, ry: 0.92 }, 5: { rx: 0.44, ry: 0.92 },
    };
    const np = {};
    for (const k of Object.keys(lay)) np[k] = { x: gx + lay[k].rx * gw, y: gy + lay[k].ry * gh };
    if (st.roads) for (const [a, b] of [[1,2],[2,3],[1,3],[3,4],[1,5]]) edge(s, np[a], np[b], { width: 1.0 });
    if (st.trains) for (const t of [3, 4, 5]) edge(s, np[1], np[t], { color: ACCENT, width: 0.9, dash: true });
    for (const [id, p] of Object.entries(np))
      node(s, p.x, p.y, id, { r: 0.15, fs: 10, fill: id === "1" ? GOLD : ICE });
    cx += 2.62;
  }

  // Painel direito: por que Dijkstra
  s.addShape("roundRect", {
    x: 8.5, y: 1.5, w: 4.35, h: 3.95, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText("Por que Dijkstra?", {
    x: 8.7, y: 1.62, w: 4.0, h: 0.5, fontFace: FONT_HEAD, fontSize: 17, bold: true, color: GOLD,
  });
  const reasons = [
    { i: "≥0", t: "Pesos não negativos (x, y ≥ 1)" },
    { i: "1", t: "Origem única: a capital" },
    { i: "+", t: "Trens são só arestas extras" },
  ];
  let py = 2.35;
  for (const rsn of reasons) {
    iconBadge(s, 9.0, py + 0.3, rsn.i, ACCENT, 0.3);
    s.addText(rsn.t, {
      x: 9.5, y: py, w: 3.2, h: 0.65, fontFace: FONT_BODY, fontSize: 14, color: WHITE, valign: "middle",
    });
    py += 1.0;
  }

  // Faixa inferior: representação
  s.addShape("roundRect", {
    x: 0.5, y: 5.7, w: 12.35, h: 1.15, fill: { color: CODEBG }, line: { color: NAVY }, rectRadius: 0.06,
  });
  s.addText("Representação", {
    x: 0.75, y: 5.8, w: 4, h: 0.4, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: GOLD,
  });
  s.addText([
    { text: "adj[u] = [ (vizinho, peso, é_trem) ]         ", options: { color: WHITE, bold: true } },
    { text: "lista de adjacência   •   V = n   •   E = 2m + k", options: { color: ICE } },
  ], { x: 0.75, y: 6.25, w: 12, h: 0.5, fontFace: FONT_MONO, fontSize: 14 });

  s.addNotes(
    "[~1 min p/ problema + modelagem]\n" +
    "Mostrar que cada cidade é um vértice, cada estrada uma aresta com peso, e o " +
    "pulo do gato: o trem é só uma aresta EXTRA da capital até a cidade. " +
    "Como todos os pesos são >= 1 (não negativos) e a origem é única (capital), " +
    "Dijkstra é a ferramenta certa."
  );
  footer(s, 3);
}

// =============================================================
// Slide 4 — Estratégia (relaxamento como diagrama)
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "3. Estratégia: Dijkstra + via_estrada");
  s.addText([
    { text: "Variação:  ", options: { bold: true, color: ACCENT } },
    { text: "uma única execução de Dijkstra, com uma anotação extra por vértice — sem expansão de estado.", options: { color: MUTED, italic: true } },
  ], { x: 0.5, y: 1.05, w: 12.3, h: 0.35, fontFace: FONT_BODY, fontSize: 13 });

  // Diagrama do relaxamento (visual dominante à esquerda)
  s.addShape("roundRect", {
    x: 0.5, y: 1.55, w: 6.4, h: 5.3, fill: { color: DARK }, line: { color: NAVY, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Relaxar  u → v", {
    x: 0.7, y: 1.65, w: 6.0, h: 0.45, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: GOLD,
  });
  // dois nós e a aresta
  const u = { x: 1.9, y: 3.0 }, v = { x: 5.4, y: 3.0 };
  edge(s, u, v, { color: ICE, width: 2, label: "w", ly: -0.22 });
  node(s, u.x, u.y, "u", { r: 0.35, fs: 20, fill: GOLD });
  node(s, v.x, v.y, "v", { r: 0.35, fs: 20 });
  s.addText("dist[u]", { x: 1.4, y: 3.4, w: 1.0, h: 0.3, fontFace: FONT_MONO, fontSize: 12, color: ICE, align: "center" });
  s.addText("dist[v]", { x: 4.9, y: 3.4, w: 1.0, h: 0.3, fontFace: FONT_MONO, fontSize: 12, color: ICE, align: "center" });
  s.addText("nd = dist[u] + w", {
    x: 0.7, y: 4.0, w: 6.0, h: 0.4, fontFace: FONT_MONO, fontSize: 15, bold: true, color: WHITE, align: "center",
  });
  // duas setas de decisão
  s.addShape("roundRect", { x: 0.8, y: 4.65, w: 5.8, h: 0.85, fill: { color: NAVY }, line: { color: GOLD, width: 1 }, rectRadius: 0.05 });
  s.addText([
    { text: "nd < dist[v]   ", options: { color: GOLD, bold: true } },
    { text: "→  dist[v]=nd,  via_estrada[v] = (não é trem)", options: { color: WHITE } },
  ], { x: 1.0, y: 4.65, w: 5.4, h: 0.85, fontFace: FONT_MONO, fontSize: 12, valign: "middle" });
  s.addShape("roundRect", { x: 0.8, y: 5.6, w: 5.8, h: 0.85, fill: { color: NAVY }, line: { color: ICE, width: 1 }, rectRadius: 0.05 });
  s.addText([
    { text: "nd = dist[v] e estrada   ", options: { color: ICE, bold: true } },
    { text: "→  via_estrada[v] = True", options: { color: WHITE } },
  ], { x: 1.0, y: 5.6, w: 5.4, h: 0.85, fontFace: FONT_MONO, fontSize: 12, valign: "middle" });

  // Coluna direita: 3 ingredientes com ícones
  const ingr = [
    { i: "⛁", c: GOLD,   h: "Fila de prioridade", t: "heapq sempre extrai o vértice de menor distância." },
    { i: "0", c: ACCENT, h: "Inicialização", t: "dist[1]=0, demais ∞.  via_estrada[1]=True." },
    { i: "✓", c: ICE,    h: "Parada", t: "Fila vazia → dist[] final, com trens já incluídos." },
  ];
  let iy = 1.55;
  for (const it of ingr) {
    s.addShape("roundRect", {
      x: 7.1, y: iy, w: 5.75, h: 1.62, fill: { color: SOFT }, line: { color: ICE, width: 1 }, rectRadius: 0.06,
    });
    iconBadge(s, 7.75, iy + 0.55, it.i, it.c, 0.34);
    s.addText(it.h, {
      x: 8.3, y: iy + 0.15, w: 4.4, h: 0.45, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: NAVY,
    });
    s.addText(it.t, {
      x: 8.3, y: iy + 0.6, w: 4.4, h: 0.9, fontFace: FONT_BODY, fontSize: 13, color: DARK,
    });
    iy += 1.78;
  }
  s.addNotes(
    "[~2 min p/ estratégia — começa aqui]\n" +
    "Rodamos UM Dijkstra a partir da capital. A novidade é o vetor via_estrada: " +
    "ele guarda se a melhor distância até cada cidade pode ser obtida só por estradas. " +
    "No relaxamento, se acho caminho mais curto atualizo a distância e marco via_estrada " +
    "conforme a aresta seja estrada ou trem; em empate por estrada, marco via_estrada = True."
  );
  footer(s, 4);
}

// =============================================================
// Slide 5 — Dijkstra em ação (grafo + tabela enxuta)
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "4. Dijkstra em ação — Exemplo 1");

  drawExemplo1(s, 0.5, 1.3, 6.2, 5.5, { distLabels: { 1: "0", 2: "1", 3: "3", 4: "5", 5: "5" } });
  s.addText("Distâncias finais (pílulas douradas)", {
    x: 0.7, y: 1.45, w: 5.7, h: 0.4, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: GOLD,
  });

  s.addText("Vetores ao final", {
    x: 7.0, y: 1.3, w: 6.0, h: 0.4, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: NAVY,
  });
  s.addShape("roundRect", {
    x: 7.0, y: 1.8, w: 5.85, h: 1.7, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08,
  });
  s.addText([
    { text: "índice         1  2  3  4  5\n", options: { color: MUTED } },
    { text: "dist            0  1  3  5  5\n", options: { color: WHITE, bold: true } },
    { text: "via_estrada     T  T  T  ", options: { color: WHITE, bold: true } },
    { text: "F", options: { color: ACCENT, bold: true } },
    { text: "  T", options: { color: WHITE, bold: true } },
  ], { x: 7.25, y: 1.95, w: 5.4, h: 1.4, fontFace: FONT_MONO, fontSize: 14, valign: "middle" });

  // Leitura visual do resultado
  s.addShape("roundRect", {
    x: 7.0, y: 3.75, w: 5.85, h: 3.1, fill: { color: SOFT }, line: { color: ICE, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Como ler", {
    x: 7.25, y: 3.85, w: 5.4, h: 0.4, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: NAVY,
  });
  const reads = [
    { c: ICE,    t: "Cidades 2, 3, 5 têm caminho ótimo por estradas." },
    { c: ACCENT, t: "Só a cidade 4 depende do trem (via_estrada = F)." },
    { c: GOLD,   t: "Logo, no exemplo, 2 dos 3 trens podem ser fechados." },
  ];
  let ry = 4.35;
  for (const rd of reads) {
    s.addShape("ellipse", { x: 7.3, y: ry + 0.06, w: 0.18, h: 0.18, fill: { color: rd.c }, line: { color: rd.c } });
    s.addText(rd.t, {
      x: 7.65, y: ry - 0.08, w: 5.0, h: 0.7, fontFace: FONT_BODY, fontSize: 14, color: DARK, valign: "middle",
    });
    ry += 0.82;
  }
  s.addNotes(
    "[~2 min p/ estratégia — continua]\n" +
    "Aqui mostro o resultado no exemplo: as distâncias finais nas pílulas douradas. " +
    "Olhando via_estrada, só a cidade 4 está marcada com F — ou seja, depende do trem. " +
    "As outras têm caminho ótimo por estrada."
  );
  footer(s, 5);
}

// =============================================================
// Slide 6 — Regras de fechamento (4 mini-grafos de situação)
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "5. Quando fechar cada trem?");

  // desenha uma "situação": capital 1, cidade s, trem (tracejado) e
  // opcionalmente uma estrada alternativa por um nó intermediário.
  function situacao(x0, y0, w, h, opt) {
    s.addShape("roundRect", {
      x: x0, y: y0, w, h, fill: { color: DARK }, line: { color: opt.cor, width: 2 }, rectRadius: 0.08,
    });
    // título + veredito
    s.addText(opt.titulo, {
      x: x0 + 0.2, y: y0 + 0.12, w: w - 0.4, h: 0.4, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: opt.cor,
    });
    s.addText(opt.cond, {
      x: x0 + 0.2, y: y0 + 0.52, w: w - 0.4, h: 0.32, fontFace: FONT_MONO, fontSize: 11, color: WHITE,
    });
    // mini grafo
    const cap = { x: x0 + 0.55, y: y0 + 1.7 };
    const dst = { x: x0 + w - 0.55, y: y0 + 1.7 };
    const mid = { x: x0 + w / 2, y: y0 + 1.15 };
    // estrada alternativa (cap-mid-dst) se houver
    if (opt.estrada) {
      edge(s, cap, mid, { color: ICE, width: 1.5 });
      edge(s, mid, dst, { color: ICE, width: 1.5, label: opt.estradaLbl, ly: -0.18 });
      node(s, mid.x, mid.y, "·", { r: 0.14, fs: 12 });
    }
    // trem direto
    edge(s, cap, dst, {
      color: opt.tremApagado ? MUTED : ACCENT, width: opt.tremForte ? 2.4 : 1.3,
      dash: true, label: opt.tremLbl, labelColor: opt.tremApagado ? MUTED : ACCENT, ly: 0.16,
    });
    node(s, cap.x, cap.y, "1", { r: 0.22, fs: 13, fill: GOLD });
    node(s, dst.x, dst.y, "s", { r: 0.22, fs: 13 });
    // veredito grande
    s.addText(opt.veredito, {
      x: x0 + 0.2, y: y0 + h - 0.55, w: w - 0.4, h: 0.45, fontFace: FONT_HEAD, fontSize: 15, bold: true,
      color: opt.cor, align: "center",
    });
  }

  const W = 3.02, H = 4.0, Y = 1.45;
  situacao(0.5, Y, W, H, {
    titulo: "Trem inútil", cond: "y > dist[s]", cor: ACCENT,
    estrada: true, estradaLbl: "3", tremLbl: "5", tremApagado: true,
    veredito: "✕ FECHA",
  });
  situacao(0.5 + (W + 0.08), Y, W, H, {
    titulo: "Alternativa por estrada", cond: "y = dist[s], via_estrada", cor: ICE,
    estrada: true, estradaLbl: "5", tremLbl: "5",
    veredito: "✕ FECHA",
  });
  situacao(0.5 + 2 * (W + 0.08), Y, W, H, {
    titulo: "Trem essencial", cond: "y = dist[s], sem estrada", cor: GOLD,
    estrada: false, tremLbl: "5", tremForte: true,
    veredito: "✓ MANTÉM 1",
  });
  situacao(0.5 + 3 * (W + 0.08), Y, W, H, {
    titulo: "Caso impossível", cond: "y < dist[s]", cor: MUTED,
    estrada: true, estradaLbl: "9", tremLbl: "2", tremApagado: true,
    veredito: "— não ocorre",
  });

  // faixa: aplicação no exemplo (curta)
  s.addShape("roundRect", {
    x: 0.5, y: 5.65, w: 12.35, h: 1.35, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.06,
  });
  s.addText("No exemplo 1", {
    x: 0.7, y: 5.72, w: 11, h: 0.35, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: GOLD,
  });
  s.addText([
    { text: "trem (3,5): ", options: { color: ICE, bold: true } },
    { text: "5 > 3 → ", options: { color: WHITE } }, { text: "fecha       ", options: { color: ACCENT, bold: true } },
    { text: "trem (4,5): ", options: { color: ICE, bold: true } },
    { text: "5 = 5, sem estrada → ", options: { color: WHITE } }, { text: "mantém       ", options: { color: GOLD, bold: true } },
    { text: "trem (5,5): ", options: { color: ICE, bold: true } },
    { text: "5 = 5, há estrada → ", options: { color: WHITE } }, { text: "fecha", options: { color: ACCENT, bold: true } },
  ], { x: 0.7, y: 6.15, w: 12, h: 0.8, fontFace: FONT_MONO, fontSize: 12.5, valign: "middle" });

  s.addNotes(
    "[~2 min p/ estratégia — fecha aqui]\n" +
    "Esta é a regra de ouro. Como todo trem sai da capital, ele é a primeira aresta do " +
    "caminho. Para cada trem (s, y): se y > dist[s] é inútil; se y = dist[s] e existe " +
    "estrada equivalente, fecha; se y = dist[s] e NÃO há estrada, é essencial e mantém um. " +
    "No exemplo: fecha 2 dos 3 trens."
  );
  footer(s, 6);
}

// =============================================================
// Slide 7 — Complexidade e casos especiais (fundidos)
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "6. Complexidade e casos especiais");

  // --- Esquerda: complexidade ---
  s.addShape("roundRect", {
    x: 0.5, y: 1.35, w: 6.0, h: 1.75, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.1,
  });
  s.addText("Tempo", {
    x: 0.75, y: 1.45, w: 4, h: 0.45, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: GOLD,
  });
  s.addText("O((V + E) · log V)", {
    x: 0.55, y: 1.9, w: 5.9, h: 1.05, fontFace: FONT_MONO, fontSize: 32, bold: true, color: WHITE, align: "center",
  });

  // chips V/E/logV
  const chips = [
    { k: "V = n", v: "vértices" },
    { k: "E = 2m + k", v: "estradas ×2 + trens" },
    { k: "log V", v: "heap (heapq)" },
  ];
  let chy = 3.35;
  for (const c of chips) {
    s.addText(c.k, { x: 0.6, y: chy, w: 2.2, h: 0.42, fontFace: FONT_MONO, fontSize: 14, bold: true, color: NAVY, valign: "middle" });
    s.addText(c.v, { x: 2.85, y: chy, w: 3.5, h: 0.42, fontFace: FONT_BODY, fontSize: 12, color: DARK, italic: true, valign: "middle" });
    chy += 0.55;
  }

  // stats reais
  const stats = [
    { big: "1625 ms", small: "tempo (lim. 2 s)", c: GOLD },
    { big: "180 MB", small: "memória (lim. 256 MB)", c: ICE },
  ];
  let sx = 0.5;
  for (const st of stats) {
    s.addShape("roundRect", { x: sx, y: 5.15, w: 2.9, h: 1.7, fill: { color: DARK }, line: { color: st.c, width: 2 }, rectRadius: 0.1 });
    s.addText(st.big, { x: sx + 0.15, y: 5.45, w: 2.6, h: 0.7, fontFace: FONT_HEAD, fontSize: 26, bold: true, color: st.c, align: "center" });
    s.addText(st.small, { x: sx + 0.15, y: 6.15, w: 2.6, h: 0.5, fontFace: FONT_BODY, fontSize: 12, color: WHITE, align: "center" });
    sx += 3.1;
  }

  // --- Direita: casos especiais ---
  s.addText("Casos especiais", {
    x: 6.9, y: 1.3, w: 6, h: 0.45, fontFace: FONT_HEAD, fontSize: 17, bold: true, color: ACCENT,
  });
  const items = [
    { i: "10⁹", t: "Pesos grandes", d: "soma ~10¹⁴ — int do Python / long em Java." },
    { i: "=",   t: "Empate trem × estrada", d: "via_estrada detecta e fecha o trem." },
    { i: "≡",   t: "Trens repetidos", d: "agrupa por destino; mantém só o necessário." },
    { i: "↓",   t: "Trem melhora distância", d: "é só mais uma aresta no Dijkstra." },
  ];
  let iy = 1.85;
  for (const it of items) {
    s.addShape("roundRect", { x: 6.9, y: iy, w: 5.95, h: 1.15, fill: { color: SOFT }, line: { color: ICE, width: 1 }, rectRadius: 0.06 });
    iconBadge(s, 7.5, iy + 0.57, it.i, it.i.length > 2 ? ICE : ACCENT, 0.34);
    s.addText(it.t, { x: 8.05, y: iy + 0.1, w: 4.7, h: 0.45, fontFace: FONT_HEAD, fontSize: 14.5, bold: true, color: NAVY });
    s.addText(it.d, { x: 8.05, y: iy + 0.55, w: 4.7, h: 0.5, fontFace: FONT_BODY, fontSize: 12.5, color: DARK });
    iy += 1.27;
  }

  s.addNotes(
    "[~1 min p/ complexidade e casos]\n" +
    "Um único Dijkstra com heap: O((V+E) log V), memória O(V+E). Na prática rodou em " +
    "1,6 s e 180 MB, dentro dos limites. Casos a citar: pesos até 10^9 (cuidado com soma, " +
    "usar long em Java; em Python tranquilo); empate trem×estrada resolvido por via_estrada; " +
    "trens repetidos agrupados por destino."
  );
  footer(s, 7);
}

// =============================================================
// Slide 8 — Conclusão / Accepted
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape("rect", { x: 0, y: 0, w: 0.4, h: 7.5, fill: { color: ACCENT }, line: { color: ACCENT } });
  s.addText("Resultado", {
    x: 0.9, y: 0.5, w: 11.5, h: 0.5, fontFace: FONT_BODY, fontSize: 16, italic: true, color: ICE,
  });
  s.addText("Accepted", {
    x: 0.9, y: 1.05, w: 11.5, h: 1.3, fontFace: FONT_HEAD, fontSize: 80, bold: true, color: GOLD,
  });
  const stats = [
    { v: "#376408159", l: "submissão" },
    { v: "1625 ms", l: "tempo (limite 2 s)" },
    { v: "180 MB", l: "memória (limite 256 MB)" },
  ];
  let x = 0.9;
  for (const st of stats) {
    s.addShape("roundRect", { x, y: 2.7, w: 3.9, h: 1.5, fill: { color: DARK }, line: { color: ACCENT, width: 1 }, rectRadius: 0.08 });
    s.addText(st.v, { x: x + 0.2, y: 2.85, w: 3.5, h: 0.7, fontFace: FONT_HEAD, fontSize: 24, bold: true, color: WHITE });
    s.addText(st.l, { x: x + 0.2, y: 3.55, w: 3.5, h: 0.55, fontFace: FONT_BODY, fontSize: 12, color: ICE, italic: true });
    x += 4.0;
  }
  s.addShape("roundRect", { x: 0.9, y: 4.45, w: 11.5, h: 2.35, fill: { color: DARK }, line: { color: GOLD, width: 1 }, rectRadius: 0.08 });
  s.addText("Em resumo", {
    x: 1.1, y: 4.55, w: 11.1, h: 0.45, fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });
  s.addText([
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Trens são apenas arestas extras a partir da capital.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Um único Dijkstra + a marca via_estrada já resolve.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Conta-se direto, por destino, quantos trens podem fechar.", options: { color: WHITE } },
  ], { x: 1.1, y: 5.1, w: 11.1, h: 1.6, fontFace: FONT_BODY, fontSize: 16, paraSpaceAfter: 4 });
  s.addText("Obrigado!  •  Grupo F  •  Codeforces 449B — Jzzhu and Cities", {
    x: 0.9, y: 6.95, w: 11.5, h: 0.4, fontFace: FONT_BODY, fontSize: 12, italic: true, color: ICE, align: "center",
  });
  s.addNotes(
    "[~1 min p/ conclusão]\n" +
    "Recapitular as 3 ideias e mostrar o Accepted: 1625 ms / 180 MB no Codeforces. " +
    "Total alvo da apresentação: 5 minutos."
  );
}

pres.writeFile({ fileName: "apresentacao.pptx" }).then((f) => console.log("escrito:", f));
