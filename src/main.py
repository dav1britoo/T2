import heapq
import sys
from collections import defaultdict


def main():
    data = sys.stdin.buffer.read().split()
    it = iter(data)
    n = int(next(it))
    m = int(next(it))
    k = int(next(it))

    # adj_full[u] = list of (v, peso, eh_trem)
    adj_full = [[] for _ in range(n + 1)]

    for _ in range(m):
        u = int(next(it))
        v = int(next(it))
        x = int(next(it))
        adj_full[u].append((v, x, False))
        adj_full[v].append((u, x, False))

    trens_por_destino = defaultdict(list)
    for _ in range(k):
        s = int(next(it))
        y = int(next(it))
        trens_por_destino[s].append(y)
        # Trem é uma aresta da capital (1) para s com peso y
        adj_full[1].append((s, y, True))

    INF = float('inf')
    dist = [INF] * (n + 1)
    dist[1] = 0

    # via_estrada[v] = True se existe um caminho mínimo até v cuja
    # última aresta é uma estrada (ou v == 1). Quando True, o trem
    # que termina em v pode ser fechado pois há alternativa por estrada.
    via_estrada = [False] * (n + 1)
    via_estrada[1] = True

    pq = [(0, 1)]
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist[u]:
            continue
        for v, w, eh_trem in adj_full[u]:
            nd = d + w
            if nd < dist[v]:
                dist[v] = nd
                via_estrada[v] = not eh_trem
                heapq.heappush(pq, (nd, v))
            elif nd == dist[v]:
                # Empate: se a aresta atual é estrada, marca alternativa por estrada.
                if not eh_trem:
                    via_estrada[v] = True

    fechaveis = 0
    for s, ys in trens_por_destino.items():
        d = dist[s]
        manteve = False
        for y in ys:
            if y > d:
                # Trem é maior que o caminho mínimo: inútil.
                fechaveis += 1
            elif y == d:
                if via_estrada[s]:
                    # Existe caminho de mesmo custo só com estradas: fecha o trem.
                    fechaveis += 1
                else:
                    # Precisamos manter exatamente um trem com este custo.
                    if manteve:
                        fechaveis += 1
                    else:
                        manteve = True
            # y < d é impossível, pois o trem foi considerado no Dijkstra.

    sys.stdout.write(str(fechaveis) + "\n")


if __name__ == "__main__":
    main()
