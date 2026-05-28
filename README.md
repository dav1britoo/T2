# Trabalho Prático 2 — Grupo F

## Problema

**Codeforces 449B — Jzzhu and Cities**

- Link: <https://codeforces.com/problemset/problem/449/B>
- Plataforma: Codeforces
- Foco: Dijkstra com análise de arestas redundantes e alternativas especiais.

## Integrantes do grupo

- João Vitor Silva
- Antonio Davi
- Pablo Dornelles

## Linguagem

Python 3.

## Como executar

Requisitos: Python 3.8 ou superior. Sem dependências externas.

A solução lê da entrada padrão e escreve na saída padrão, exatamente no
formato exigido pelo Codeforces.

```bash
python src/main.py < dados/exemplo1.txt
# saída esperada: 2

python src/main.py < dados/exemplo2.txt
# saída esperada: 2
```

No Windows / PowerShell:

```powershell
Get-Content dados/exemplo1.txt | python src/main.py
```

## Enunciado resumido

Existem `n` cidades, sendo a cidade `1` a capital. Há `m` estradas
bidirecionais com pesos e `k` rotas de trem, cada uma ligando a capital
diretamente a uma cidade `s_i` com peso `y_i`. Deseja-se fechar o maior
número possível de rotas de trem **sem alterar a distância mínima da
capital até qualquer cidade**.

## Modelagem como grafo

- **Vértices**: as `n` cidades.
- **Arestas**:
  - cada estrada vira uma aresta não direcionada `(u, v)` com peso `x`;
  - cada trem `(s_i, y_i)` é tratado como uma aresta extra **da capital `1`
    até `s_i`** com peso `y_i`, marcada como sendo do tipo "trem".
- **Origem**: capital, vértice `1`. Calculamos a distância mínima de `1`
  até todos os outros vértices.
- **Pesos**: todos não negativos (`x, y ≥ 1`), portanto Dijkstra é aplicável.
- **Representação**: lista de adjacência. Cada item armazena
  `(vizinho, peso, é_trem)`.

A ideia central é que um trem só é necessário quando ele é, sozinho, a
única forma de atingir o melhor custo até sua cidade destino. Logo, basta
saber, para cada cidade `s` que recebe trens, qual é o caminho mínimo
`dist[s]` considerando estradas **e** trens, e se há uma alternativa
usando apenas estradas que alcança o mesmo custo.

## Algoritmo

**Dijkstra com fila de prioridade mínima** (`heapq`), executado uma única
vez a partir da capital, sobre o grafo unificado (estradas + trens).

Durante o relaxamento mantemos um vetor adicional
`via_estrada[v]`: vale `True` quando existe pelo menos um caminho mínimo
até `v` cuja **última aresta é uma estrada** (ou `v = 1`).

Regra de atualização ao relaxar `u → v` com custo `nd`:

- `nd < dist[v]`: atualiza `dist[v] = nd`, `via_estrada[v] = (não é trem)`.
- `nd == dist[v]`: se a aresta atual **é estrada**, marca
  `via_estrada[v] = True` (descobrimos alternativa por estrada com o
  mesmo custo ótimo).

Após o Dijkstra, percorremos os trens agrupados por destino `s`:

- `y > dist[s]`: trem inútil → **fecha**.
- `y == dist[s]` e `via_estrada[s]`: existe caminho ótimo só por estradas
  → **fecha**.
- `y == dist[s]` e **não** `via_estrada[s]`: precisamos manter exatamente
  um trem com esse custo; todos os demais empatados → **fecha**.
- `y < dist[s]`: impossível, pois o trem participou do Dijkstra.

A contagem total é a resposta.

### Variação de Dijkstra utilizada

Dijkstra clássico de origem única (fila de prioridade), com **anotação
extra por vértice** (`via_estrada`) que registra o tipo da última aresta
em algum caminho mínimo. Não é necessária a expansão de estado nem
múltiplas execuções: uma única passada resolve o problema.

## Complexidade

Sejam `V = n`, `E = 2m + k` (estradas em ambos sentidos + trens):

- Tempo: `O((V + E) · log V)` para o Dijkstra com `heapq`, mais `O(k)`
  para a contagem final.
- Memória: `O(V + E)` para a lista de adjacência e estruturas auxiliares.

Para os limites do problema (`n ≤ 10^5`, `m ≤ 3·10^5`, `k ≤ 10^5`),
isso é confortavelmente suficiente dentro de 2 s e 256 MB.

## Casos especiais tratados

- **Distâncias grandes**: pesos até `10^9` e somas até `~10^14`. Em
  Python os inteiros são arbitrários, não há overflow.
- **Múltiplos trens para a mesma cidade**: agrupamos por destino e
  mantemos no máximo um trem por destino quando ele é estritamente
  necessário.
- **Múltiplas arestas entre o mesmo par de cidades**: tratadas
  naturalmente, pois cada uma vira uma entrada independente na lista de
  adjacência.
- **Empate entre estrada e trem (`y = dist[s]`)**: a marcação
  `via_estrada[s]` garante que reconhecemos a existência de uma rota
  alternativa por estradas e podemos fechar o trem.
- **Trem que melhora a distância**: o trem é apenas mais uma aresta no
  Dijkstra, então pode efetivamente reduzir `dist[s]`. Nesse caso
  `via_estrada[s]` permanece `False` e o trem é mantido.

## Evidência de Accepted

Submissão aceita no Codeforces — ver
[`evidencias/accepted.png`](evidencias/accepted.png).

## Estrutura do repositório

```text
T2/
├── README.md
├── src/
│   └── main.py
├── evidencias/
│   └── accepted.png
├── apresentacao/
│   └── apresentacao.pdf
└── dados/
    ├── exemplo1.txt
    └── exemplo2.txt
```
