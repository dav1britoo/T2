const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Codeforces 449B - Jzzhu and Cities";

// --- Paleta "Midnight Executive" ---
const NAVY = "1E2761";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const ACCENT = "F96167"; // coral para destaque
const GOLD = "F9E795";
const MUTED = "8A93B8";
const DARK = "0F1638";
const CODEBG = "12193C";

const FONT_HEAD = "Georgia";
const FONT_BODY = "Calibri";
const FONT_MONO = "Consolas";

// Helper: rodapé padrão em slides de conteúdo
function footer(slide, n) {
  slide.addShape("rect", { x: 0, y: 7.15, w: 13.333, h: 0.35, fill: { color: NAVY }, line: { color: NAVY } });
  slide.addText("Grupo F • Codeforces 449B — Jzzhu and Cities", {
    x: 0.4, y: 7.18, w: 9, h: 0.3, fontFace: FONT_BODY, fontSize: 10, color: ICE,
  });
  slide.addText(`${n} / 8`, {
    x: 12.4, y: 7.18, w: 0.6, h: 0.3, fontFace: FONT_BODY, fontSize: 10, color: ICE, align: "right",
  });
}

// Helper: título de slide de conteúdo
function title(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 0.35, w: 12.3, h: 0.7,
    fontFace: FONT_HEAD, fontSize: 30, bold: true, color: NAVY,
  });
}

// =============================================================
// Slide 1 — Capa
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  // Bloco lateral decorativo
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

  // Cartão de identificação
  s.addShape("roundRect", {
    x: 0.9, y: 4.6, w: 11.5, h: 2.0,
    fill: { color: DARK }, line: { color: ACCENT, width: 1 },
    rectRadius: 0.1,
  });
  s.addText("Grupo F", {
    x: 1.2, y: 4.7, w: 5, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 22, bold: true, color: GOLD,
  });
  s.addText([
    { text: "Integrantes:  ", options: { bold: true, color: ICE } },
    { text: "João Vitor Silva  •  Antonio Davi  •  Pablo Dornelles", options: { color: WHITE } },
  ], {
    x: 1.2, y: 5.25, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: 16,
  });
  s.addText([
    { text: "Orientador:    ", options: { bold: true, color: ICE } },
    { text: "Prof. Me Ricardo Carubbi", options: { color: WHITE } },
  ], {
    x: 1.2, y: 5.7, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: 16,
  });
  s.addText([
    { text: "Linguagem:   ", options: { bold: true, color: ICE } },
    { text: "Python 3   •   Plataforma: Codeforces", options: { color: WHITE } },
  ], {
    x: 1.2, y: 6.15, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: 16,
  });

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

  // Coluna esquerda — enunciado
  s.addShape("roundRect", {
    x: 0.5, y: 1.3, w: 6.2, h: 5.5,
    fill: { color: "F4F6FC" }, line: { color: ICE, width: 1 }, rectRadius: 0.08,
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
    { text: " estradas bidirecionais com pesos " },
    { text: "x", options: { italic: true } },
    { text: ".\n" },
    { text: "k", options: { bold: true } },
    { text: " rotas de trem ligando a capital a uma cidade " },
    { text: "s", options: { italic: true } },
    { text: " com custo " },
    { text: "y", options: { italic: true } },
    { text: "." },
  ], {
    x: 0.8, y: 1.95, w: 5.6, h: 1.6,
    fontFace: FONT_BODY, fontSize: 15, color: DARK, paraSpaceAfter: 4,
  });

  s.addText("Objetivo", {
    x: 0.8, y: 3.7, w: 5.6, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: ACCENT,
  });
  s.addText(
    "Fechar o maior número possível de rotas de trem sem alterar a " +
    "distância mínima da capital a nenhuma cidade.",
    {
      x: 0.8, y: 4.15, w: 5.6, h: 1.5,
      fontFace: FONT_BODY, fontSize: 15, color: DARK,
    });

  s.addText("Limites:  n ≤ 10⁵   •   m ≤ 3·10⁵   •   k ≤ 10⁵   •   pesos até 10⁹", {
    x: 0.8, y: 6.05, w: 5.6, h: 0.6,
    fontFace: FONT_MONO, fontSize: 12, color: NAVY, italic: true,
  });

  // Coluna direita — diagrama do exemplo
  s.addShape("roundRect", {
    x: 7.0, y: 1.3, w: 5.85, h: 5.5,
    fill: { color: DARK }, line: { color: NAVY, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Exemplo 1", {
    x: 7.2, y: 1.45, w: 5.5, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });

  // Nós (posição em polegadas, dentro do retângulo direito 7.0..12.85, 1.3..6.8)
  const nodes = {
    1: { x: 8.0,  y: 3.6 },   // capital
    2: { x: 9.6,  y: 2.4 },
    3: { x: 10.9, y: 3.6 },
    4: { x: 12.2, y: 4.8 },
    5: { x: 9.6,  y: 5.4 },
  };
  // Arestas (estradas) — desenhar como linhas
  const edges = [
    [1,2,"1"], [2,3,"2"], [1,3,"3"], [3,4,"4"], [1,5,"5"],
  ];
  for (const [a,b,w] of edges) {
    const A = nodes[a], B = nodes[b];
    s.addShape("line", {
      x: A.x, y: A.y, w: B.x - A.x, h: B.y - A.y,
      line: { color: ICE, width: 1.5 },
    });
    s.addText(w, {
      x: (A.x + B.x)/2 - 0.2, y: (A.y + B.y)/2 - 0.15, w: 0.4, h: 0.3,
      fontFace: FONT_MONO, fontSize: 11, color: GOLD, align: "center", bold: true,
    });
  }
  // Trens (linha tracejada/coral)
  const trains = [[3,"5"],[4,"5"],[5,"5"]];
  for (const [t,y] of trains) {
    const A = nodes[1], B = nodes[t];
    s.addShape("line", {
      x: A.x, y: A.y, w: B.x - A.x, h: B.y - A.y,
      line: { color: ACCENT, width: 1.2, dashType: "dash" },
    });
  }

  // Desenhar nós por cima
  for (const [id, p] of Object.entries(nodes)) {
    const isCap = id === "1";
    s.addShape("ellipse", {
      x: p.x - 0.25, y: p.y - 0.25, w: 0.5, h: 0.5,
      fill: { color: isCap ? GOLD : ICE },
      line: { color: isCap ? ACCENT : NAVY, width: 2 },
    });
    s.addText(id, {
      x: p.x - 0.25, y: p.y - 0.27, w: 0.5, h: 0.5,
      fontFace: FONT_HEAD, fontSize: 16, bold: true, color: NAVY, align: "center", valign: "middle",
    });
  }

  // Legenda
  s.addShape("line", {
    x: 7.25, y: 6.35, w: 0.6, h: 0,
    line: { color: ICE, width: 2 },
  });
  s.addText("estrada", { x: 7.95, y: 6.2, w: 1.3, h: 0.3, fontFace: FONT_BODY, fontSize: 11, color: ICE });
  s.addShape("line", {
    x: 9.3, y: 6.35, w: 0.6, h: 0,
    line: { color: ACCENT, width: 2, dashType: "dash" },
  });
  s.addText("trem (da capital)", { x: 10.0, y: 6.2, w: 2.5, h: 0.3, fontFace: FONT_BODY, fontSize: 11, color: ACCENT });

  footer(s, 2);
}

// =============================================================
// Slide 3 — Modelagem como grafo
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "2. Modelagem como grafo");

  const items = [
    { h: "Vértices",  t: "As n cidades. A capital é o vértice 1, origem do Dijkstra." },
    { h: "Arestas — estradas", t: "Cada estrada (u, v, x) vira aresta bidirecional de peso x." },
    { h: "Arestas — trens", t: "Cada trem (s, y) vira uma aresta extra ligando 1 → s com peso y, marcada como tipo \"trem\"." },
    { h: "Pesos", t: "Todos não negativos (x, y ≥ 1) ⇒ Dijkstra é aplicável." },
    { h: "Representação", t: "Lista de adjacência. Cada item: (vizinho, peso, é_trem)." },
  ];

  let y = 1.35;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    // Cartão
    s.addShape("roundRect", {
      x: 0.6, y, w: 12.1, h: 0.95,
      fill: { color: i % 2 === 0 ? "F4F6FC" : WHITE },
      line: { color: ICE, width: 1 }, rectRadius: 0.06,
    });
    // Número
    s.addShape("ellipse", {
      x: 0.85, y: y + 0.2, w: 0.55, h: 0.55,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(String(i + 1), {
      x: 0.85, y: y + 0.18, w: 0.55, h: 0.55,
      fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD, align: "center", valign: "middle",
    });
    s.addText(it.h, {
      x: 1.6, y: y + 0.1, w: 3.4, h: 0.4,
      fontFace: FONT_HEAD, fontSize: 16, bold: true, color: ACCENT,
    });
    s.addText(it.t, {
      x: 1.6, y: y + 0.5, w: 11.0, h: 0.45,
      fontFace: FONT_BODY, fontSize: 14, color: DARK,
    });
    y += 1.1;
  }

  footer(s, 3);
}

// =============================================================
// Slide 4 — Estratégia: Dijkstra + via_estrada
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "3. Estratégia: Dijkstra + marcação de alternativa");

  // Subtítulo
  s.addText("Uma única execução de Dijkstra a partir da capital, sobre estradas + trens.", {
    x: 0.5, y: 1.05, w: 12.3, h: 0.35,
    fontFace: FONT_BODY, fontSize: 14, italic: true, color: MUTED,
  });

  // Bloco esquerdo — pseudocódigo
  s.addShape("roundRect", {
    x: 0.5, y: 1.55, w: 6.5, h: 5.3,
    fill: { color: CODEBG }, line: { color: NAVY, width: 1 }, rectRadius: 0.06,
  });
  s.addText("Relaxamento adaptado", {
    x: 0.7, y: 1.65, w: 6.2, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: GOLD,
  });
  s.addText([
    { text: "ao relaxar  u → v  com custo  nd:\n\n", options: { italic: true, color: ICE } },
    { text: "se  nd < dist[v]:\n", options: { bold: true, color: WHITE } },
    { text: "    dist[v] = nd\n", options: { color: WHITE } },
    { text: "    via_estrada[v] = (não é trem)\n\n", options: { color: WHITE } },
    { text: "senão se  nd == dist[v]:\n", options: { bold: true, color: WHITE } },
    { text: "    se a aresta é estrada:\n", options: { color: WHITE } },
    { text: "        via_estrada[v] = True", options: { color: GOLD } },
  ], {
    x: 0.7, y: 2.1, w: 6.2, h: 4.6,
    fontFace: FONT_MONO, fontSize: 14, paraSpaceAfter: 2,
  });

  // Bloco direito — contagem final
  s.addShape("roundRect", {
    x: 7.2, y: 1.55, w: 5.6, h: 5.3,
    fill: { color: "F4F6FC" }, line: { color: ICE, width: 1 }, rectRadius: 0.06,
  });
  s.addText("Contagem final, por trem (s, y)", {
    x: 7.4, y: 1.65, w: 5.3, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: NAVY,
  });

  const rules = [
    { tag: "y > dist[s]",  res: "fecha", color: ACCENT, desc: "Trem mais caro que o caminho mínimo: inútil." },
    { tag: "y = dist[s] e via_estrada[s]", res: "fecha", color: ACCENT, desc: "Existe alternativa só por estradas." },
    { tag: "y = dist[s] e ¬via_estrada[s]", res: "mantém 1 / fecha demais", color: NAVY, desc: "Trem é a única forma de alcançar o ótimo." },
    { tag: "y < dist[s]", res: "impossível", color: MUTED, desc: "O próprio trem já estava no Dijkstra." },
  ];
  let yr = 2.15;
  for (const r of rules) {
    s.addText(r.tag, {
      x: 7.4, y: yr, w: 5.2, h: 0.32,
      fontFace: FONT_MONO, fontSize: 13, bold: true, color: NAVY,
    });
    s.addText("→ " + r.res, {
      x: 7.4, y: yr + 0.32, w: 5.2, h: 0.32,
      fontFace: FONT_BODY, fontSize: 13, bold: true, color: r.color,
    });
    s.addText(r.desc, {
      x: 7.4, y: yr + 0.62, w: 5.2, h: 0.4,
      fontFace: FONT_BODY, fontSize: 11, italic: true, color: DARK,
    });
    yr += 1.15;
  }

  footer(s, 4);
}

// =============================================================
// Slide 5 — Por que funciona
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "4. Por que funciona");

  // Texto principal
  s.addShape("roundRect", {
    x: 0.5, y: 1.3, w: 12.3, h: 1.8,
    fill: { color: "F4F6FC" }, line: { color: ICE, width: 1 }, rectRadius: 0.06,
  });
  s.addText([
    { text: "Observação-chave: ", options: { bold: true, color: ACCENT } },
    { text: "todo trem sai da capital. Portanto, em qualquer caminho mínimo, " +
            "um trem só pode aparecer como ", options: { color: DARK } },
    { text: "a primeira (e única) aresta", options: { bold: true, color: NAVY } },
    { text: ". Para vértices que não são destino direto de trem, o problema " +
            "é Dijkstra padrão — só precisamos decidir o destino de cada trem.",
      options: { color: DARK } },
  ], {
    x: 0.8, y: 1.45, w: 11.7, h: 1.55,
    fontFace: FONT_BODY, fontSize: 15,
  });

  // Três caixas com casos
  const cases = [
    {
      titulo: "Trem inútil",
      cond: "y > dist[s]",
      txt: "O melhor caminho até s já é menor que o trem; fechar não muda nada.",
      cor: ACCENT,
    },
    {
      titulo: "Alternativa por estrada",
      cond: "y = dist[s] e via_estrada[s]",
      txt: "Há um caminho ótimo cujo último passo é estrada — manter o trem é redundante.",
      cor: ICE,
    },
    {
      titulo: "Trem essencial",
      cond: "y = dist[s] e ¬via_estrada[s]",
      txt: "Sem trem, dist[s] aumentaria. Mantém-se exatamente um; duplicados são fechados.",
      cor: GOLD,
    },
  ];
  let cx = 0.5;
  for (const c of cases) {
    s.addShape("roundRect", {
      x: cx, y: 3.4, w: 4.06, h: 3.4,
      fill: { color: DARK }, line: { color: c.cor, width: 2 }, rectRadius: 0.08,
    });
    s.addText(c.titulo, {
      x: cx + 0.25, y: 3.55, w: 3.6, h: 0.5,
      fontFace: FONT_HEAD, fontSize: 18, bold: true, color: c.cor,
    });
    s.addText(c.cond, {
      x: cx + 0.25, y: 4.1, w: 3.6, h: 0.45,
      fontFace: FONT_MONO, fontSize: 13, color: GOLD,
    });
    s.addShape("line", {
      x: cx + 0.25, y: 4.7, w: 3.6, h: 0,
      line: { color: c.cor, width: 1 },
    });
    s.addText(c.txt, {
      x: cx + 0.25, y: 4.85, w: 3.6, h: 1.8,
      fontFace: FONT_BODY, fontSize: 13, color: ICE,
    });
    cx += 4.27;
  }

  footer(s, 5);
}

// =============================================================
// Slide 6 — Complexidade
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "5. Complexidade");

  // Estatística grande
  s.addShape("roundRect", {
    x: 0.5, y: 1.4, w: 6.2, h: 5.4,
    fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.1,
  });
  s.addText("Tempo", {
    x: 0.8, y: 1.55, w: 5.6, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });
  s.addText("O((V + E) · log V)", {
    x: 0.8, y: 2.15, w: 5.6, h: 1.2,
    fontFace: FONT_MONO, fontSize: 36, bold: true, color: WHITE,
  });
  s.addText([
    { text: "V = n           ", options: { color: ICE } },
    { text: "(vértices)\n", options: { color: MUTED, italic: true } },
    { text: "E = 2m + k    ", options: { color: ICE } },
    { text: "(estradas dos dois lados + trens)\n\n", options: { color: MUTED, italic: true } },
    { text: "+ O(k)          ", options: { color: ICE } },
    { text: "contagem final sobre os trens", options: { color: MUTED, italic: true } },
  ], {
    x: 0.8, y: 3.5, w: 5.6, h: 2.0,
    fontFace: FONT_MONO, fontSize: 14,
  });
  s.addText("Fila de prioridade: heapq (heap binário).", {
    x: 0.8, y: 6.2, w: 5.6, h: 0.5,
    fontFace: FONT_BODY, fontSize: 13, italic: true, color: ICE,
  });

  // Memória + limites
  s.addShape("roundRect", {
    x: 7.0, y: 1.4, w: 5.85, h: 2.5,
    fill: { color: "F4F6FC" }, line: { color: ICE, width: 1 }, rectRadius: 0.08,
  });
  s.addText("Memória", {
    x: 7.2, y: 1.55, w: 5.5, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: ACCENT,
  });
  s.addText("O(V + E)", {
    x: 7.2, y: 2.1, w: 5.5, h: 0.6,
    fontFace: FONT_MONO, fontSize: 26, bold: true, color: NAVY,
  });
  s.addText(
    "Listas de adjacência + vetores dist, via_estrada e a heap.",
    {
      x: 7.2, y: 2.85, w: 5.5, h: 1.0,
      fontFace: FONT_BODY, fontSize: 13, color: DARK,
    });

  // Caixa "no problema"
  s.addShape("roundRect", {
    x: 7.0, y: 4.1, w: 5.85, h: 2.7,
    fill: { color: DARK }, line: { color: ACCENT, width: 1 }, rectRadius: 0.08,
  });
  s.addText("No problema", {
    x: 7.2, y: 4.25, w: 5.5, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 18, bold: true, color: GOLD,
  });
  s.addText([
    { text: "n ≤ 10⁵   m ≤ 3·10⁵   k ≤ 10⁵\n", options: { color: WHITE, bold: true } },
    { text: "⇒ ~7·10⁵ arestas, ~6·10⁶ operações de heap.\n\n", options: { color: ICE } },
    { text: "Submissão real: ", options: { color: ICE } },
    { text: "1625 ms / 180 MB", options: { color: GOLD, bold: true } },
    { text: "  (limite 2 s / 256 MB).", options: { color: ICE } },
  ], {
    x: 7.2, y: 4.85, w: 5.5, h: 1.9,
    fontFace: FONT_MONO, fontSize: 13,
  });

  footer(s, 6);
}

// =============================================================
// Slide 7 — Casos especiais
// =============================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  title(s, "6. Casos especiais");

  const items = [
    {
      h: "Pesos grandes (até 10⁹)",
      t: "Distâncias podem chegar a ~10¹⁴. Python usa inteiros arbitrários, não há overflow.",
    },
    {
      h: "Múltiplos trens para a mesma cidade",
      t: "Agrupamos por destino. Mantemos no máximo um quando estritamente necessário; os demais empatados são fechados.",
    },
    {
      h: "Arestas paralelas entre o mesmo par",
      t: "Tratadas naturalmente — cada uma vira uma entrada própria na lista de adjacência.",
    },
    {
      h: "Empate trem × estrada (y = dist[s])",
      t: "A marcação via_estrada[s] garante reconhecer alternativa por estrada e fechar o trem redundante.",
    },
    {
      h: "Trem que melhora a distância",
      t: "É só mais uma aresta no Dijkstra. dist[s] é atualizado pelo trem e via_estrada[s] permanece falso ⇒ trem mantido.",
    },
    {
      h: "Conectividade garantida",
      t: "O enunciado garante que toda cidade alcança a capital, então nenhum vértice fica com distância infinita.",
    },
  ];

  // Grid 3x2
  let i = 0;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 3; col++) {
      const x = 0.5 + col * 4.27;
      const y = 1.3 + row * 2.85;
      const it = items[i++];
      s.addShape("roundRect", {
        x, y, w: 4.06, h: 2.6,
        fill: { color: "F4F6FC" }, line: { color: NAVY, width: 1 }, rectRadius: 0.06,
      });
      // Faixa de cor à esquerda
      s.addShape("rect", {
        x, y, w: 0.12, h: 2.6,
        fill: { color: ACCENT }, line: { color: ACCENT },
      });
      s.addText(it.h, {
        x: x + 0.3, y: y + 0.2, w: 3.7, h: 0.8,
        fontFace: FONT_HEAD, fontSize: 14, bold: true, color: NAVY,
      });
      s.addText(it.t, {
        x: x + 0.3, y: y + 1.05, w: 3.7, h: 1.45,
        fontFace: FONT_BODY, fontSize: 12, color: DARK,
      });
    }
  }

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
    x: 0.9, y: 0.5, w: 11.5, h: 0.5,
    fontFace: FONT_BODY, fontSize: 16, italic: true, color: ICE,
  });
  s.addText("Accepted", {
    x: 0.9, y: 1.05, w: 11.5, h: 1.3,
    fontFace: FONT_HEAD, fontSize: 80, bold: true, color: GOLD,
  });

  // Três caixas de stats
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

  // Conclusão
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
    { text: "Modelagem unificada: trens são apenas arestas extras da capital.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Uma única execução de Dijkstra com a marcação via_estrada já resolve.\n", options: { color: WHITE } },
    { text: "•  ", options: { color: ACCENT, bold: true } },
    { text: "Contagem direta sobre os trens, agrupados por destino, dá a resposta.", options: { color: WHITE } },
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
