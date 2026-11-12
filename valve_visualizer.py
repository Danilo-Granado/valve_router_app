#!/usr/bin/env python3
"""
valve_visualizer.py - draw the piping network from the valve database as an
interactive diagram (HTML page you open in any browser).

Shows equipment, junctions (DV), segments and every valve in its position
along its segment. Optionally highlights a route found by valve_router.py,
colouring the valves to OPEN and CLOSE.

Usage
    python valve_visualizer.py valve_database.xlsx
    python valve_visualizer.py valve_database.xlsx --from TQ-101 --to RE-202 --product Leite
    python valve_visualizer.py valve_database.xlsx -o diagrama.html --layout saved --open

Layout: by default each area (column Área of the Nós sheet) gets its own block, flow runs
from left to right and items are ordered to avoid crossing lines. Lines are drawn with right
angles, routed around the items; lines that still cross without being connected are drawn
with a small jump.

Needs valve_router.py in the same folder.  pip install openpyxl networkx
The page loads the drawing library from the internet. For PCs without internet,
download vis-network.min.js once (see README note at the end) and pass --offline-lib.
"""
# --- BEGIN VISUALIZER CORE ---
import json
import html as _html

# Node styles: key -> (label, vis shape, fill colour, border colour, size)
NODE_STYLES = {
    "tank":      ("Tanque (TQ)",            "database", "#9DC3E6", "#1F4E79", 28),
    "reactor":   ("Reator (RE)",            "hexagon",  "#F4B183", "#843C0C", 26),
    "mixer":     ("Misturador (MS)",        "diamond",  "#C9B3E6", "#5B2C8F", 24),
    "pump":      ("Bomba (BB)",             "circle",   "#A9D18E", "#385723", 18),
    "hx":        ("Trocador de calor (TC)", "square",   "#8FD3D0", "#1B6B68", 20),
    "filter":    ("Filtro (FT)",            "triangleDown", "#D9D9D9", "#404040", 18),
    "junction":  ("Derivação (DV)",         "dot",      "#404040", "#000000", 7),
    "manifold":  ("Coletor (CL)",           "box",      "#D9D9D9", "#404040", 18),
    "drain":     ("Dreno (DR)",             "triangleDown", "#C9A27E", "#5E3B1C", 16),
    "cip":       ("Central CIP",            "database", "#FFE699", "#7F6000", 24),
    "equipment": ("Equipamento (EQ)",       "box",      "#F2F2F2", "#404040", 18),
}

# Node Type column values (English and Portuguese) -> style key
TYPE_WORDS = {
    "tank": "tank", "tanque": "tank", "silo": "tank",
    "reactor": "reactor", "reator": "reactor",
    "mixer": "mixer", "misturador": "mixer",
    "pump": "pump", "bomba": "pump",
    "heat exchanger": "hx", "trocador de calor": "hx",
    "filter": "filter", "filtro": "filter",
    "junction": "junction", "derivacao": "junction", "derivação": "junction",
    "manifold": "manifold", "coletor": "manifold",
    "drain": "drain", "dreno": "drain",
    "cip": "cip", "central cip": "cip",
}
STATUS_PT = {"in service": "Em operação", "out of service": "Fora de operação",
             "locked closed": "Travada fechada", "locked open": "Travada aberta"}
# Tag codes -> style key (used when Node Type is blank or generic "Equipment")
TAG_CODES = {
    "TQ": "tank", "T": "tank", "RE": "reactor", "MS": "mixer", "M": "mixer",
    "BB": "pump", "P": "pump", "TC": "hx", "HX": "hx", "FT": "filter",
    "DV": "junction", "J": "junction", "CL": "manifold", "DR": "drain",
    "CIP": "cip", "EQ": "equipment",
}

VALVE_NAMES = {
    "VES": "Esfera", "VBO": "Borboleta", "VGA": "Gaveta", "VGL": "Globo", "VDI": "Diafragma",
    "VMA": "Macho", "VAG": "Agulha", "VRT": "Retenção", "V3V": "Três vias",
}

ROUTE_COLORS = {"open": "#3BA55C", "close": "#E04848", "endpoint": "#2F75D6", "edge": "#F28C28"}


def classify_node(nid, ntype):
    t = (ntype or "").strip().lower()
    if t in TYPE_WORDS:
        return TYPE_WORDS[t]
    for token in nid.upper().split("-"):
        letters = "".join(ch for ch in token if ch.isalpha())
        if letters in TAG_CODES:
            return TAG_CODES[letters]
    return "equipment"


def valve_code(tag):
    for token in tag.upper().split("-"):
        if token in VALVE_NAMES:
            return token
    return None


def route_overlay(plant, G, path):
    """Return the sets the diagram needs to highlight one route."""
    open_v, close_v, endpoint_v, warns = plant.valve_states(path)
    route_edges = set()  # (segment, frozenset(a, b)) with hidden points collapsed
    pending = None
    for a, b in zip(path, path[1:]):
        sid = G[a][b]["segment"]
        if plant.is_hidden(b):
            pending = a
            continue
        if plant.is_hidden(a):
            route_edges.add((sid, frozenset((pending, b))))
        else:
            route_edges.add((sid, frozenset((a, b))))
    return {
        "nodes": {p for p in path if not plant.is_hidden(p)},
        "edges": route_edges,
        "open": list(open_v),
        "close": {t: list(v) for t, v in close_v.items()},
        "endpoint": list(endpoint_v),
        "warnings": warns,
        "path": [p for p in path if not plant.is_hidden(p)],
        "segments": plant.path_segments(G, path),
    }


# --------------------------------------------------------------------------- #
#  Layout: one horizontal band per area, flow from left to right, few crossings
# --------------------------------------------------------------------------- #
LAYOUT_GAP_X = 34     # free space between neighbouring columns
LAYOUT_GAP_Y = 34     # free space between items stacked in one column (room for lines between them)
LAYOUT_BAND_GAP = 150 # space between two area bands (room for the box, its title and lines)

# (half width, half height) of each drawing shape, label included
SHAPE_EXTENT = {"box": (46, 24), "database": (46, 30), "circle": (34, 28), "hexagon": (34, 42),
                "diamond": (32, 40), "triangleDown": (30, 34), "dot": (22, 18), "square": (30, 30)}
VALVE_EXTENT = (25, 21)


def point_areas(plant):
    """Area of every node and valve. A valve takes the area of the nearer end of its segment."""
    area = {nid: n.get("area", "") for nid, n in plant.nodes.items()}
    for s in plant.segments.values():
        vs = s["valves"]
        for i, t in enumerate(vs):
            area[t] = area.get(s["from"] if i < (len(vs) + 1) // 2 else s["to"], "")
    return area


def _natural(a):
    import re
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", a)]


def _band_orders(groups, links, keep=4):
    """Candidate top-to-bottom orders of the area bands, best first. groups: lists of areas, one list
    per row of the drawing (top row first); bands only move within their row. Cheapest = areas
    sharing many lines next to each other, upstream areas on top."""
    import itertools

    def cost(order):
        pos = {a: i for i, a in enumerate(order)}
        return sum(w * abs(pos[a] - pos[b]) + (0.01 * w if pos[a] > pos[b] else 0) for (a, b), w in links.items())

    choices = [list(itertools.permutations(g)) if len(g) <= 6 else [tuple(g)] for g in groups]
    total = 1
    for c in choices:
        total *= len(c)
    if total <= 5000:
        orders = [sum((list(p) for p in combo), []) for combo in itertools.product(*choices)]
    else:  # too many: rows in their natural order, then swap neighbours while it helps
        best, better = sum((list(g) for g in groups), []), True
        row = {a: i for i, g in enumerate(groups) for a in g}
        while better:
            better = False
            for i in range(len(best) - 1):
                if row[best[i]] == row[best[i + 1]]:
                    trial = best[:i] + [best[i + 1], best[i]] + best[i + 2:]
                    if cost(trial) < cost(best):
                        best, better = trial, True
        orders = [best]
    orders.sort(key=cost)  # stable: equal cost keeps the natural order
    return orders[:keep]


def _vessel(plant, p):
    """Tanks, reactors, mixers, CIP: where loops normally start and end."""
    n = plant.nodes.get(p)
    return n is not None and (not n["pass"] or classify_node(p, n["type"]) in ("tank", "reactor", "mixer", "cip"))


def _loop_vessels(points, edges, plant):
    """Vessels that lie on a loop, with (lines in, lines out) inside that loop."""
    import networkx as nx
    G = nx.DiGraph()
    G.add_nodes_from(points)
    G.add_edges_from(edges)
    out = {}
    for comp in nx.strongly_connected_components(G):
        if len(comp) > 1:
            for p in comp:
                if _vessel(plant, p):
                    out[p] = (sum(1 for q in G.predecessors(p) if q in comp), sum(1 for q in G.successors(p) if q in comp))
    return out


def _break_cycles_tanks(points, edges, plant, ends=frozenset()):
    """Edges (a, b) to draw reversed so the drawing has no loops. Within each loop the walk starts
    at a tank / reactor (a point routes can't pass through), so the line returning to it is reversed.
    Vessels in ends are walked last instead: they end their loop and their outlet line is reversed."""
    import networkx as nx
    G = nx.DiGraph()
    G.add_nodes_from(points)
    G.add_edges_from(edges)
    reverse = set()
    for comp in nx.strongly_connected_components(G):
        if len(comp) < 2:
            continue
        outside_in = {p: sum(1 for q in G.predecessors(p) if q not in comp) for p in comp}
        holds = {p: _vessel(plant, p) for p in comp}
        # how far the flow is from where it enters the loop: the first vessel it reaches starts the loop
        dist, frontier = {}, [p for p in comp if outside_in[p]]
        for p in frontier:
            dist[p] = 0
        while frontier:
            nxt = []
            for p in frontier:
                for q in G.successors(p):
                    if q in comp and q not in dist:
                        dist[q] = dist[p] + 1
                        nxt.append(q)
            frontier = nxt
        roots = sorted(comp, key=lambda p: (p in ends, not holds[p], dist.get(p, 0), -outside_in[p], p))
        state = {}
        for root in roots:
            if root in state:
                continue
            state[root] = "open"
            stack = [(root, iter(sorted((q for q in G.successors(root) if q in comp), key=lambda q: (q in ends, q))))]
            while stack:
                node, it = stack[-1]
                nxt = next(it, None)
                if nxt is None:
                    state[node] = "done"
                    stack.pop()
                elif state.get(nxt) == "open":
                    reverse.add((node, nxt))
                elif nxt not in state:
                    state[nxt] = "open"
                    stack.append((nxt, iter(sorted((q for q in G.successors(nxt) if q in comp), key=lambda q: (q in ends, q)))))
    return reverse


def _break_cycles_greedy(points, edges, plant):
    """Edges (a, b) to draw reversed so the drawing has no loops, as few as possible (greedy
    Eades-Lin-Smyth ordering). Ties go to tanks, reactors etc. as the start of their loop."""
    succ = {p: {} for p in points}
    pred = {p: {} for p in points}
    for a, b in edges:
        succ[a][b] = succ[a].get(b, 0) + 1
        pred[b][a] = pred[b].get(a, 0) + 1
    left = set(points)
    first, last = [], []

    def remove(p):
        left.discard(p)
        for q in succ[p]:
            pred[q].pop(p, None)
        for q in pred[p]:
            succ[q].pop(p, None)

    holds = {p: _vessel(plant, p) for p in points}
    while left:
        changed = True
        while changed:
            changed = False
            for p in sorted(left):
                if not succ[p]:
                    last.insert(0, p)
                    remove(p)
                    changed = True
                elif not pred[p]:
                    first.append(p)
                    remove(p)
                    changed = True
        if left:
            p = max(sorted(left), key=lambda p: (sum(succ[p].values()) - sum(pred[p].values()), holds[p]))
            first.append(p)
            remove(p)
    rank = {p: i for i, p in enumerate(first + last)}
    return {(a, b) for a, b in edges if rank[a] > rank[b]}


def _columns(points, dag):
    """Column of every point: each line goes at least one column to the right, and the total
    length of all lines is as small as possible (solved exactly as a min-cost flow problem)."""
    import networkx as nx
    mult = {}
    for a, b in dag:
        mult[(a, b)] = mult.get((a, b), 0) + 1
    F = nx.DiGraph()
    for p in points:
        F.add_node(p, demand=0)
    for (a, b), m in mult.items():
        F.add_edge(a, b, weight=-1)
        F.nodes[b]["demand"] += m
        F.nodes[a]["demand"] -= m
    _, flow = nx.network_simplex(F)
    # columns from the flow: every line spans >= 1 column, lines carrying flow exactly 1
    C = nx.DiGraph()
    C.add_nodes_from(points)
    for (a, b) in mult:
        C.add_edge(b, a, weight=-1)
        if flow[a][b] > 0:
            C.add_edge(a, b, weight=1)
    src = ("~src",)
    C.add_edges_from(((src, p) for p in points), weight=0)
    dist = nx.single_source_bellman_ford_path_length(C, src)
    low = min(dist[p] for p in points)
    return {p: dist[p] - low for p in points}


def _isotonic(values, lo=None, hi=None):
    """Least-squares non-decreasing fit of values (pool adjacent violators), clipped to [lo, hi]."""
    blocks = []  # [mean, count]
    for v in values:
        blocks.append([v, 1])
        while len(blocks) > 1 and blocks[-2][0] > blocks[-1][0]:
            m2, n2 = blocks.pop()
            m1, n1 = blocks.pop()
            blocks.append([(m1 * n1 + m2 * n2) / (n1 + n2), n1 + n2])
    out = []
    for m, n in blocks:
        if lo is not None:
            m = max(m, lo)
        if hi is not None:
            m = min(m, hi)
        out += [m] * n
    return out


ORTHO_GAP, ORTHO_TRACK, ORTHO_STUB, ORTHO_SHORT = 5, 10, 20, 26  # same as GAP, TRACK, STUB, SHORT in DIAGRAM_JS
_CELL = 160  # grid cell for finding nearby line pieces quickly


def _extent_js(shape, size):
    """[half width, above, below] of an item, as extent() in DIAGRAM_JS."""
    if shape in ("box", "database"):
        return 36, 24, 24
    if shape == "circle":
        return 26, 24, 24
    return max(size + 2, 20), size + 2, size + 17


def _pieces(pts):
    out = []
    for (ax, ay), (bx, by) in zip(pts, pts[1:]):
        if abs(ax - bx) < 0.5 and abs(ay - by) < 0.5:
            continue
        h = abs(ay - by) <= abs(ax - bx)
        out.append((h, (ay + by) / 2, min(ax, bx), max(ax, bx)) if h else (h, (ax + bx) / 2, min(ay, by), max(ay, by)))
    return out


def _simplify(pts):
    out = []
    for p in pts:
        if out and abs(out[-1][0] - p[0]) < 0.5 and abs(out[-1][1] - p[1]) < 0.5:
            continue
        out.append(p)
        while len(out) >= 3:
            (ax, ay), (bx, by), (cx, cy) = out[-3:]
            if (abs(ax - bx) < 0.5 and abs(bx - cx) < 0.5) or (abs(ay - by) < 0.5 and abs(by - cy) < 0.5):
                del out[-2]
            else:
                break
    return out


def _cells(seg, margin):
    h, c, lo, hi = seg
    x1, x2, y1, y2 = (lo, hi, c, c) if h else (c, c, lo, hi)
    return [(i, j) for i in range(int((x1 - margin) // _CELL), int((x2 + margin) // _CELL) + 1)
            for j in range(int((y1 - margin) // _CELL), int((y2 + margin) // _CELL) + 1)]


def ortho_routes(plant, pos):
    """Right-angle routes {(from, to, segment, index): [points]} for the drawing at pos: the same
    choice as orthoRoutes in DIAGRAM_JS (change both together)."""
    shape = {}
    for nid, n in plant.nodes.items():
        _, sh, _, _, size = NODE_STYLES[classify_node(nid, n["type"])]
        shape[nid] = (sh, size)
    for tag, v in plant.valves.items():
        check = v["type"].lower() == "check" or valve_code(tag) == "VRT"
        shape[tag] = ("triangle", 9) if check else ("square", 8)
    P, boxes, label_below = {}, [], {}
    for p, (x, y) in pos.items():
        if p not in shape:
            continue
        P[p] = (x, y)
        label_below[p] = shape[p][0] not in ("box", "database", "circle")
        w, up, down = _extent_js(*shape[p])
        boxes.append((p, x - w, x + w, y - up, y + down))
    cols = []
    for x in sorted(q[0] for q in P.values()):
        if cols and x - cols[-1][1] <= 20:
            cols[-1][1] = x
        else:
            cols.append([x, x])
    channels = [(cols[i - 1][1] + cols[i][0]) / 2 for i in range(1, len(cols))]

    edges = []
    for sid, sg in plant.segments.items():
        chain = [sg["from"], *sg["valves"], sg["to"]]
        for i, (a, b) in enumerate(zip(chain, chain[1:])):
            if a in P and b in P and a != b:
                edges.append((f"{sid}#{i}", a, b))

    def span(e):
        (ax, ay), (bx, by) = P[e[1]], P[e[2]]
        return abs(bx - ax) + abs(by - ay)

    edges.sort(key=lambda e: (span(e), e[0]))
    grid = {}  # cell -> [placed piece numbers]
    placed = []  # (h, c, lo, hi, ends)
    routes = {}
    G, T = ORTHO_GAP, ORTHO_TRACK
    for eid, fa, fb in edges:
        a, b = P[fa], P[fb]
        dx, dy = b[0] - a[0], b[1] - a[1]
        lo, hi = min(a[0], b[0]), max(a[0], b[0])
        after = next((c for c in channels if c > a[0] + 1), None)
        before = next((c for c in reversed(channels) if c < b[0] - 1), None)
        x1 = after if after is not None else a[0] + ORTHO_STUB
        x2 = before if before is not None else b[0] - ORTHO_STUB
        near = [o for o in boxes if o[0] != fa and o[0] != fb and o[2] > min(lo, x2) - T and o[1] < max(hi, x1) + T]
        merged = []
        for t, u in sorted(((o[3] - G, o[4] + G) for o in near), key=lambda q: q[0]):
            if merged and t <= merged[-1][1]:
                merged[-1][1] = max(merged[-1][1], u)
            else:
                merged.append([t, u])
        tracks = []
        if merged:
            tracks += [merged[0][0] - T, merged[-1][1] + T]
            tracks += [(merged[i - 1][1] + merged[i][0]) / 2 for i in range(1, len(merged)) if merged[i][0] - merged[i - 1][1] >= 2 * T]
        mid = (a[1] + b[1]) / 2
        ys = sorted(tracks, key=lambda t: abs(t - mid))[:4]

        options = []
        if abs(dy) < 1 or abs(dx) < 1:
            options.append([a, b])
        else:
            options += [[a, (b[0], a[1]), b], [a, (a[0], b[1]), b]]
        ch = [c for c in channels if lo + 1 < c < hi - 1]
        if not ch and hi - lo > 1:
            ch = [(lo + hi) / 2]
        if len(ch) > 3:
            ch = [ch[0], ch[len(ch) // 2], ch[-1]]
        if abs(dy) >= 1:
            for c in ch:
                for k in (0, -1, 1, -2, 2):
                    xm = c + k * T
                    options.append([a, (xm, a[1]), (xm, b[1]), b])
        for ym in ys:
            options.append([a, (a[0], ym), (b[0], ym), b])
            if dx < 0 or x1 < x2:
                options.append([a, (x1, a[1]), (x1, ym), (x2, ym), (x2, b[1]), b])

        best = None
        for raw in options:
            pts = _simplify(raw)
            segs = _pieces(pts)
            cost, length = 0.6 * (len(pts) - 2), 0.0
            if label_below[fa] and abs(pts[1][0] - pts[0][0]) < 0.5 and pts[1][1] > pts[0][1]:
                cost += 1.5
            if label_below[fb] and abs(pts[-2][0] - pts[-1][0]) < 0.5 and pts[-2][1] > pts[-1][1]:
                cost += 1.5
            if len(pts) > 2:
                for p, q in ((pts[0], pts[1]), (pts[-2], pts[-1])):
                    if abs(p[0] - q[0]) + abs(p[1] - q[1]) < ORTHO_SHORT:
                        cost += 2
            for h, c, slo, shi in segs:
                length += shi - slo
                for o in near:
                    if h:
                        hit = o[3] - G < c < o[4] + G and shi > o[1] - G and slo < o[2] + G
                    else:
                        hit = o[1] - G < c < o[2] + G and shi > o[3] - G and slo < o[4] + G
                    if hit:
                        cost += 100
                seen = set()
                for cell in _cells((h, c, slo, shi), T):
                    for k in grid.get(cell, ()):
                        if k in seen:
                            continue
                        seen.add(k)
                        rh, rc, rlo, rhi, ends = placed[k]
                        shares = fa in ends or fb in ends
                        if h == rh:
                            if abs(c - rc) < T - 2 and min(shi, rhi) - max(slo, rlo) > 2:
                                cost += 3 if shares else 12
                        else:
                            hz, vt = ((c, slo, shi), (rc, rlo, rhi)) if h else ((rc, rlo, rhi), (c, slo, shi))
                            if hz[1] + 0.5 < vt[0] < hz[2] - 0.5 and vt[1] + 0.5 < hz[0] < vt[2] - 0.5:
                                cost += 1
                if best is not None and cost >= best[0]:
                    break
            cost += 0.004 * length
            if best is None or cost < best[0]:
                best = (cost, pts)
        routes[(fa, fb, eid)] = best[1]
        for sg in _pieces(best[1]):
            placed.append((*sg, (fa, fb)))
            for cell in _cells(sg, 0):
                grid.setdefault(cell, []).append(len(placed) - 1)
    return routes, boxes


def ortho_score(plant, pos):
    """(crossings, lines lying on top of each other, lines over items) of the right-angle drawing."""
    routes, boxes = ortho_routes(plant, pos)
    G, T = ORTHO_GAP, ORTHO_TRACK
    segs = [(*sg, (fa, fb)) for (fa, fb, _), pts in routes.items() for sg in _pieces(pts)]
    grid = {}
    for k, sg in enumerate(segs):
        for cell in _cells(sg[:4], 0):
            grid.setdefault(cell, []).append(k)
    cross = overlap = 0
    for k, (h, c, lo, hi, ends) in enumerate(segs):
        seen = set()
        for cell in _cells((h, c, lo, hi), T):
            for j in grid.get(cell, ()):
                if j <= k or j in seen:
                    continue
                seen.add(j)
                rh, rc, rlo, rhi, rends = segs[j]
                shares = ends[0] in rends or ends[1] in rends
                if h == rh:
                    overlap += not shares and abs(c - rc) < T - 2 and min(hi, rhi) - max(lo, rlo) > 2
                else:
                    hz, vt = ((c, lo, hi), (rc, rlo, rhi)) if h else ((rc, rlo, rhi), (c, lo, hi))
                    cross += hz[1] + 0.5 < vt[0] < hz[2] - 0.5 and vt[1] + 0.5 < hz[0] < vt[2] - 0.5
    over = 0
    for h, c, lo, hi, ends in segs:
        for o in boxes:
            if o[0] in ends:
                continue
            if h:
                over += o[3] - G < c < o[4] + G and hi > o[1] - G and lo < o[2] + G
            else:
                over += o[1] - G < c < o[2] + G and hi > o[3] - G and lo < o[4] + G
    return cross, overlap, over


def area_layout(plant, budget=3.0):
    """Positions {tag: [x, y]} for every node and valve: each area in its own horizontal band,
    flow from left to right, item order chosen to keep line crossings low.

    Loops can be opened in more than one way, bands stacked in more than one order and long rows
    wrapped or not. Several combinations are laid out (most promising first, for about `budget`
    seconds) and the one whose drawing has the fewest crossings for its size is kept."""
    import time
    start = time.time()
    edges = []
    for s in plant.segments.values():
        chain = [s["from"], *s["valves"], s["to"]]
        edges += list(zip(chain, chain[1:]))
    vessels = _loop_vessels(list(plant.nodes) + list(plant.valves), edges, plant)
    collectors = frozenset(p for p, (i, o) in vessels.items() if i > o)
    # (how to open the loops, extra cost): the greedy way reverses the fewest lines but ignores the
    # process (it may draw a pump's outlet backwards), so it only wins when clearly better
    breakers = [(lambda pts, es, pl: _break_cycles_tanks(pts, es, pl), 0)]
    if collectors:
        breakers.append((lambda pts, es, pl: _break_cycles_tanks(pts, es, pl, collectors), 0))
    breakers.append((_break_cycles_greedy, 4))
    lines = [(a, b) for a, b in edges if a != b]

    def cost(pos, choice, extra=0):
        cross, overlap, over = ortho_score(plant, pos)
        over += overlap
        xs, ys = [p[0] for p in pos.values()], [p[1] for p in pos.values()]
        size = max(max(xs) - min(xs), (max(ys) - min(ys)) * 16 / 9)
        backwards = sum(1 for a, b in lines if a in pos and b in pos and pos[b][0] < pos[a][0] - 1)
        return cross + 3 * over + size / 250 + 1.5 * choice + backwards + extra

    best = best_cost = None
    seen = set()
    for choice in range(3):
        for wrap in (0.25, 0.6):
            for breaker, extra in breakers:
                if best is not None and time.time() - start > budget:
                    break
                found = _area_layout(plant, breaker, choice, wrap, seen)
                if found is None:
                    continue
                c = cost(found[0], choice, extra)
                if best is None or c < best_cost:
                    best, best_cost, best_choice, best_extra = found, c, choice, extra
    if best is None:
        return {}
    pos, columns, half_h = best
    polished = {p: list(v) for p, v in pos.items()}
    _polish(plant, polished, columns, half_h, deadline=start + budget * 1.5)
    if cost(polished, best_choice, best_extra) <= best_cost:
        pos = polished
    _level(plant, pos, columns, half_h)
    return pos


def _level(plant, pos, columns, half_h, reach=16):
    """Put connected items that are nearly level (less than reach apart) exactly level, so their line
    is straight instead of having a small step, when that keeps the spacing in their column."""
    nbrs = {p: set() for p in pos}
    for s in plant.segments.values():
        chain = [s["from"], *s["valves"], s["to"]]
        for a, b in zip(chain, chain[1:]):
            if a in pos and b in pos and a != b:
                nbrs[a].add(b)
                nbrs[b].add(a)
    column = {}
    for col in columns:
        for i, p in enumerate(col):
            column[p] = (col, i)
    for _ in range(3):
        moved = False
        for p in sorted(pos, key=lambda q: len(nbrs[q])):  # items with few lines move first
            near = sorted((abs(pos[q][1] - pos[p][1]), pos[q][1]) for q in nbrs[p] if 0 < abs(pos[q][1] - pos[p][1]) < reach)
            if not near:
                continue
            target = near[0][1]
            if p in column:
                col, i = column[p]
                if i > 0 and target - pos[col[i - 1]][1] < half_h[col[i - 1]] + half_h[p] + LAYOUT_GAP_Y - 2:
                    continue
                if i + 1 < len(col) and pos[col[i + 1]][1] - target < half_h[col[i + 1]] + half_h[p] + LAYOUT_GAP_Y - 2:
                    continue
            pos[p][1] = target
            moved = True
        if not moved:
            break


def _polish(plant, pos, columns, half_h, deadline=None):
    """Swap neighbours within a column while that removes crossings / lines over items in the
    drawing as it will be shown (straight lines), then respace each column."""
    lines = []
    for s in plant.segments.values():
        chain = [s["from"], *s["valves"], s["to"]]
        lines += [(a, b) for a, b in zip(chain, chain[1:]) if a in pos and b in pos and a != b]
    touching = {p: [] for p in pos}
    for i, (a, b) in enumerate(lines):
        touching[a].append(i)
        touching[b].append(i)

    def side(a, b, c):
        return (c[0] - a[0]) * (b[1] - a[1]) - (c[1] - a[1]) * (b[0] - a[0])

    def crosses(i, j):
        a, b = lines[i]
        c, d = lines[j]
        if a in (c, d) or b in (c, d):
            return False
        p1, p2, p3, p4 = pos[a], pos[b], pos[c], pos[d]
        return side(p3, p4, p1) * side(p3, p4, p2) < 0 and side(p1, p2, p3) * side(p1, p2, p4) < 0

    def over(i, n):
        a, b = lines[i]
        if n in (a, b):
            return False
        p1, p2, q = pos[a], pos[b], pos[n]
        dx, dy = p2[0] - p1[0], p2[1] - p1[1]
        t = max(0.0, min(1.0, ((q[0] - p1[0]) * dx + (q[1] - p1[1]) * dy) / ((dx * dx + dy * dy) or 1)))
        return (q[0] - p1[0] - t * dx) ** 2 + (q[1] - p1[1] - t * dy) ** 2 < 14 ** 2

    def cost(items):
        mine = {i for p in items for i in touching[p]}
        c = sum(1 for i in mine for j in range(len(lines)) if j not in mine and crosses(i, j))
        c += sum(1 for i in mine for j in mine if i < j and crosses(i, j))
        c += 2 * sum(1 for i in mine for n in pos if over(i, n))
        c += 2 * sum(1 for i in range(len(lines)) if i not in mine for n in items if over(i, n))
        return c

    import time
    for _ in range(4):
        changed = False
        if deadline and time.time() > deadline:
            break
        for col in columns:
            for k in range(len(col) - 1):
                u, v = col[k], col[k + 1]
                if not touching[u] and not touching[v]:
                    continue
                before = cost((u, v))
                pos[u][1], pos[v][1] = pos[v][1], pos[u][1]
                if cost((u, v)) < before:
                    col[k], col[k + 1] = v, u
                    changed = True
                else:
                    pos[u][1], pos[v][1] = pos[v][1], pos[u][1]
        if not changed:
            break
    for col in columns:  # keep the spacing after swapping items of different sizes
        offs = [0.0]
        for k in range(1, len(col)):
            offs.append(offs[-1] + half_h[col[k - 1]] + half_h[col[k]] + LAYOUT_GAP_Y)
        z = _isotonic([pos[p][1] - o for p, o in zip(col, offs)])
        for p, zz, o in zip(col, z, offs):
            pos[p][1] = round(zz + o)


def _area_layout(plant, break_cycles, band_choice=0, wrap=0.25, seen=None):
    """One layout attempt: (positions, columns, half heights), or None if there is no band
    order number band_choice (or this combination gave the same rows as one in seen).
    wrap: areas overlapping the row above by less than this share of the narrower one start a new row.
    columns: lists of items stacked in one column of one band, top first."""
    area = point_areas(plant)
    points = list(plant.nodes) + [v for v in plant.valves if v in area]
    if not points:
        return None
    pset = set(points)
    edges = []
    for s in plant.segments.values():
        chain = [s["from"], *s["valves"], s["to"]]
        edges += [(a, b) for a, b in zip(chain, chain[1:]) if a in pset and b in pset and a != b]

    def extent(p):
        if p in plant.valves:
            return VALVE_EXTENT
        return SHAPE_EXTENT.get(NODE_STYLES[classify_node(p, plant.nodes[p]["type"])][1], (30, 30))

    # 1. columns: open the loops, then the shortest total line length
    reverse = break_cycles(points, edges, plant)
    dag = [(b, a) if (a, b) in reverse else (a, b) for a, b in edges]
    succ = {p: [] for p in points}
    pred = {p: [] for p in points}
    for a, b in dag:
        succ[a].append(b)
        pred[b].append(a)
    layer = _columns(points, dag)
    for _ in range(2):  # items with slack sit midway between their neighbours
        for p in points:
            if not pred[p] or not succ[p]:
                continue
            lo, hi = max(layer[q] for q in pred[p]) + 1, min(layer[q] for q in succ[p]) - 1
            near = sorted(layer[q] for q in pred[p] + succ[p])
            layer[p] = min(max(near[len(near) // 2], lo), hi)
    low = min(layer.values())
    layer = {p: v - low for p, v in layer.items()}
    nlayers = max(layer.values()) + 1

    # 2. x: column positions, wide enough for the widest item in each column
    half_w = [6] * nlayers
    for p in points:
        half_w[layer[p]] = max(half_w[layer[p]], extent(p)[0])
    xs = [0.0]
    for li in range(1, nlayers):
        xs.append(xs[-1] + half_w[li - 1] + half_w[li] + LAYOUT_GAP_X)

    # 3. Area bands. Areas whose columns overlap share a row of the drawing; an area that starts
    # after the previous ones end goes back to the left on a new row (like wrapping text).
    reach = {}  # area -> (left, right)
    for p in points:
        lo_x, hi_x = reach.get(area[p], (float("inf"), float("-inf")))
        reach[area[p]] = (min(lo_x, xs[layer[p]] - half_w[layer[p]]), max(hi_x, xs[layer[p]] + half_w[layer[p]]))
    rows = []
    for name in sorted(reach, key=lambda n: (reach[n][0], n == "", _natural(n))):
        lo_x, hi_x = reach[name]
        if rows:
            overlap = rows[-1]["right"] - lo_x
            narrow = min(hi_x - lo_x, rows[-1]["right"] - rows[-1]["left"])
            if overlap > -LAYOUT_BAND_GAP and overlap > wrap * narrow:
                rows[-1]["names"].append(name)
                rows[-1]["right"] = max(rows[-1]["right"], hi_x)
                continue
        rows.append({"names": [name], "left": lo_x, "right": hi_x})
    links = {}
    for a, b in edges:
        if area[a] != area[b]:
            links[(area[a], area[b])] = links.get((area[a], area[b]), 0) + 1
    for r in rows:
        r["names"].sort(key=lambda n: (n == "", _natural(n)))
    orders = _band_orders([r["names"] for r in rows], links)
    if band_choice >= len(orders):
        return None
    if seen is not None:
        sig = (break_cycles, band_choice, tuple(tuple(r["names"]) for r in rows))
        if sig in seen:
            return None
        seen.add(sig)
    band = {a: i for i, a in enumerate(orders[band_choice])}
    row_of = {band[n]: i for i, r in enumerate(rows) for n in r["names"]}
    row_left = [r["left"] for r in rows]

    # 4. split lines that skip columns with hidden points, one per column
    key = {p: float(band[area[p]]) for p in points}  # band key (fractional for lines between bands)
    inner = {p: True for p in points}  # takes part in the vertical spacing of its band
    line_of = {}  # hidden point -> (from, to) of the line it belongs to
    up = {p: [] for p in points}
    down = {p: [] for p in points}
    for i, (a, b) in enumerate(dag):
        span = layer[b] - layer[a]
        prev = a
        for k in range(1, span):
            d = ("~", i, k)
            layer[d] = layer[a] + k
            key[d] = key[a] + (key[b] - key[a]) * k / span
            inner[d] = key[a] == key[b]
            line_of[d] = (a, b)
            up[d], down[d] = [], []
            down[prev].append(d)
            up[d].append(prev)
            prev = d
        down[prev].append(b)
        up[b].append(prev)
    layers = [[] for _ in range(nlayers)]
    for p in sorted(layer, key=lambda p: (key[p], str(p))):
        layers[layer[p]].append(p)

    # 5. order within each column: barycentre sweeps, keeping each area together
    def crossings(order):
        pos = {p: i for col in order for i, p in enumerate(col)}
        total = 0
        for col in order[:-1]:
            pairs = [(pos[p], pos[q]) for p in col for q in down[p]]
            total += sum(1 for i, (a1, b1) in enumerate(pairs) for a2, b2 in pairs[i + 1:]
                         if (a1 - a2) * (b1 - b2) < 0)
        return total

    def sweep(order, downward):
        order = [list(c) for c in order]
        rng = range(1, nlayers) if downward else range(nlayers - 2, -1, -1)
        for li in rng:
            ref = {p: i for i, p in enumerate(order[li - 1 if downward else li + 1])}
            col = order[li]
            bary = {}
            for i, p in enumerate(col):
                nb = [ref[q] for q in (up[p] if downward else down[p]) if q in ref]
                bary[p] = sum(nb) / len(nb) if nb else i * len(ref) / max(len(col), 1)
            col.sort(key=lambda p: (key[p], bary[p]))
        return order

    best, best_c = layers, crossings(layers)
    order = layers
    for it in range(24):
        order = sweep(order, it % 2 == 0)
        c = crossings(order)
        if c < best_c:
            best, best_c = order, c
    order = best

    def pair_cross(u, v, pos):  # crossings between the lines of u and v when u is above v
        c = 0
        for nb in (up, down):
            for a in nb[u]:
                for b in nb[v]:
                    if pos[a] > pos[b]:
                        c += 1
        return c

    improved = True
    for _ in range(10):  # swap neighbours in a column while that removes crossings
        if not improved:
            break
        improved = False
        pos = {p: i for col in order for i, p in enumerate(col)}
        for col in order:
            for i in range(len(col) - 1):
                u, v = col[i], col[i + 1]
                if key[u] != key[v]:
                    continue
                if pair_cross(v, u, pos) < pair_cross(u, v, pos):
                    col[i], col[i + 1] = v, u
                    pos[u], pos[v] = i + 1, i
                    improved = True

    x = {p: xs[layer[p]] for p in layer}

    def half_h(p):
        return extent(p)[1] if p in pset else 8

    # 6. y: within each band, pull every item level with its neighbours, keeping the order and spacing
    y = {}
    for col in order:
        by_band = {}
        for p in col:
            if inner[p]:
                by_band.setdefault(key[p], []).append(p)
        for seq in by_band.values():
            acc = 0.0
            for i, p in enumerate(seq):
                if i:
                    acc += half_h(seq[i - 1]) + half_h(p) + LAYOUT_GAP_Y
                y[p] = acc
            mid = acc / 2
            for p in seq:
                y[p] -= mid
    cross_nb = {p: [] for p in points}  # neighbours in other bands, for the last passes
    for a, b in dag:
        if key[a] != key[b]:
            cross_nb[a].append(b)
            cross_nb[b].append(a)

    def place(bounds=None):
        for it in range(14 if bounds is None else 4):
            cols = order if it % 2 == 0 else order[::-1]
            for col in cols:
                by_band = {}
                for p in col:
                    if inner[p]:
                        by_band.setdefault(key[p], []).append(p)
                for bk, seq in by_band.items():
                    want = []
                    for p in seq:
                        if p in line_of:  # hidden point: on the straight line between the real ends
                            a, b = line_of[p]
                            t = (x[p] - x[a]) / ((x[b] - x[a]) or 1)
                            want.append(y[a] + (y[b] - y[a]) * t)
                            continue
                        nb = [y[q] for q in up[p] + down[p] if inner[q] and key[q] == key[p]]
                        w = len(nb)
                        if bounds is not None and cross_nb[p]:
                            nb += [0.5 * y[q] for q in cross_nb[p]]
                            w += 0.5 * len(cross_nb[p])
                        want.append(sum(nb) / w if w else y[p])
                    offs = [0.0]
                    for i in range(1, len(seq)):
                        offs.append(offs[-1] + half_h(seq[i - 1]) + half_h(seq[i]) + LAYOUT_GAP_Y)
                    lo = hi = None
                    if bounds is not None:
                        top, bottom = bounds[bk]
                        lo, hi = top + half_h(seq[0]), bottom - half_h(seq[-1]) - offs[-1]
                        hi = max(hi, lo)
                    z = _isotonic([w - o for w, o in zip(want, offs)], lo, hi)
                    for p, zz, o in zip(seq, z, offs):
                        y[p] = zz + o

    place()
    # 7. stack the bands, row by row; each row starts again at the left
    bounds, cursor = {}, 0.0
    for row in range(len(rows)):
        for bk in sorted(b for b in set(key[p] for p in points) if row_of[b] == row):
            members = [p for p in y if inner[p] and key[p] == bk]
            top = min(y[p] - half_h(p) for p in members)
            bottom = max(y[p] + half_h(p) for p in members)
            for p in members:
                y[p] += cursor - top
            bounds[bk] = (cursor, cursor + bottom - top)
            cursor += bottom - top + LAYOUT_BAND_GAP
    place(bounds)  # now also lean towards lines going to other areas, staying inside the band
    cols = {}
    for p in points:
        cols.setdefault((layer[p], key[p]), []).append(p)
    columns = [sorted(c, key=lambda p: y[p]) for c in cols.values() if len(c) > 1]
    pos = {p: [round(x[p] - row_left[row_of[key[p]]]), round(y[p])] for p in points}
    return pos, columns, {p: half_h(p) for p in points}


def build_diagram(plant, route=None):
    nodes, edges = [], []
    used_kinds = set()
    areas = point_areas(plant)

    for nid, n in plant.nodes.items():
        kind = classify_node(nid, n["type"])
        used_kinds.add(kind)
        _, shape, fill, border, size = NODE_STYLES[kind]
        tip = [f"{nid}  ({n['type'] or NODE_STYLES[kind][0]})"]
        if n["desc"]:
            tip.append(n["desc"])
        if n.get("area"):
            tip.append(f"Área: {n['area']}")
        tip.append("Passagem: " + ("Sim" if n["pass"] else "Não"))
        if not n["active"]:
            tip.append("FORA DE OPERAÇÃO")
        node = {
            "id": nid, "label": nid, "shape": shape, "size": size,
            "color": {"background": fill, "border": border},
            "borderWidth": 2, "tip": "\n".join(tip), "group": "node", "area": areas.get(nid, ""),
            "font": {"size": 14 if kind != "junction" else 11, "face": "Arial"},
        }
        if shape in ("box", "database", "circle"):
            # these shapes hold the label inside; give them a sensible minimum size
            node["font"]["size"] = 13
            node["margin"] = 8
            node["widthConstraint"] = {"minimum": 64 if shape != "circle" else 40}
            if shape == "database":
                node["heightConstraint"] = {"minimum": 34}
                node["font"]["vadjust"] = 7
        if not n["active"]:
            node["color"] = {"background": "#EEEEEE", "border": "#E04848"}
            node["shapeProperties"] = {"borderDashes": [4, 3]}
        if route and nid in route["nodes"]:
            node["borderWidth"] = 4
            node["color"]["border"] = ROUTE_COLORS["edge"]
        nodes.append(node)

    for tag, v in plant.valves.items():
        seg = plant.segments[v["segment"]]
        code = valve_code(tag)
        vtype = v["type"].lower()
        is_check = vtype == "check" or code == "VRT"
        status = v["status"]
        fill, border, dashes = "#FFFFFF", "#404040", False
        if status == "locked open":
            border, dashes = "#3BA55C", True
        elif status in ("locked closed", "out of service"):
            fill, border = "#F8CBAD", "#E04848"
        if route:
            if tag in route["open"]:
                fill, border = ROUTE_COLORS["open"], "#1E5E33"
            elif tag in route["close"]:
                fill, border = ROUTE_COLORS["close"], "#7A1F1F"
            elif tag in route["endpoint"]:
                fill, border = ROUTE_COLORS["endpoint"], "#173E73"
        body = v.get("body") or VALVE_NAMES.get(code) or ("Retenção" if is_check else "Válvula")
        tip = [f"{tag}  –  {body}",
               f"Trecho {v['segment']} ({seg['from']} → {seg['to']}), posição {v['order']:g}",
               f"Acionamento: {'Automática' if vtype == 'automatic' else ('Passiva' if is_check else 'Manual')}"
               f"   Status: {STATUS_PT.get(status, status)}"]
        if v["notes"]:
            tip.append(v["notes"])
        if route and tag in route["close"]:
            sid, at = route["close"][tag]
            tip.append(f"FECHAR: isola {sid} em {at}")
        node = {
            "id": tag, "label": tag, "shape": "triangle" if is_check else "square",
            "size": 9 if is_check else 8, "group": "valve", "area": areas.get(tag, ""),
            "color": {"background": fill, "border": border}, "borderWidth": 2,
            "font": {"size": 11, "face": "Arial", "vadjust": -4},
            "tip": "\n".join(tip),
        }
        if dashes:
            node["shapeProperties"] = {"borderDashes": [3, 2]}
        nodes.append(node)

    for sid, s in plant.segments.items():
        chain = [s["from"], *s["valves"], s["to"]]
        bidir = s["bidir"] and not any(plant.valves[t]["type"].lower() == "check" for t in s["valves"])
        prods = ", ".join(sorted(s["products"])) if s["products"] else "qualquer"
        tip = (f"{sid}: {s['from']} → {s['to']}\n"
               f"Sentido: {'Bidirecional' if bidir else 'Unidirecional'}\n"
               f"Produtos: {prods}\nCusto: {s['cost']:g}"
               + ("" if s["active"] else "\nFORA DE OPERAÇÃO"))
        mid = (len(chain) - 2) // 2  # sub-edge that carries the segment label
        for i, (a, b) in enumerate(zip(chain, chain[1:])):
            on_route = bool(route) and (sid, frozenset((a, b))) in route["edges"]
            edge = {
                "id": f"{sid}#{i}", "from": a, "to": b, "tip": tip,
                "arrows": {"to": {"enabled": True, "scaleFactor": 0.6},
                           "from": {"enabled": bidir, "scaleFactor": 0.6}},
                "color": {"color": "#7F7F7F"}, "width": 2,
                "font": {"size": 11, "face": "Arial", "color": "#1F3864", "strokeWidth": 3, "align": "top"},
            }
            if i == mid:
                edge["label"] = sid
            if not s["active"]:
                edge["dashes"] = [6, 4]
                edge["color"] = {"color": "#E04848"}
            if s["products"] is not None:
                edge["width"] = 2.5
            if on_route:
                edge["color"] = {"color": ROUTE_COLORS["edge"]}
                edge["width"] = 6
            edges.append(edge)

    return {"nodes": nodes, "edges": edges, "kinds": sorted(used_kinds)}


def legend_html(kinds):
    rows = []
    for k in kinds:
        label, shape, fill, border, _ = NODE_STYLES[k]
        if shape == "triangleDown":
            style = f"border-top-color:{border}"
        else:
            style = f"background:{fill};border-color:{border}"
        rows.append(f'<div class="lg"><span class="sw {shape}" style="{style}"></span>{_html.escape(label)}</div>')
    rows += [
        '<div class="lg"><span class="sw square" style="background:#fff;border-color:#404040"></span>Válvula</div>',
        '<div class="lg"><span class="sw triangle"></span>Válvula de retenção</div>',
        '<div class="lg"><span class="sw square" style="background:#fff;border:2px dashed #3BA55C"></span>Travada aberta</div>',
        '<div class="lg"><span class="sw square" style="background:#F8CBAD;border-color:#E04848"></span>Travada fechada / fora de operação</div>',
        '<div class="lg"><span class="ar">──▶</span>Unidirecional</div>',
        '<div class="lg"><span class="ar">◀──▶</span>Bidirecional</div>',
        '<div class="lg"><span class="ln" style="border-top-width:3px"></span>Linha com restrição de produto</div>',
        '<div class="lg"><span class="ln" style="border-top:2px dashed #E04848"></span>Trecho fora de operação</div>',
        '<div class="lg"><svg width="26" height="16" style="flex:none"><path d="M0 11H8M18 11H26M13 1V16" '
        'stroke="#7F7F7F" stroke-width="2" fill="none"/><path d="M8 11A5 5 0 0 1 18 11" stroke="#7F7F7F" '
        'stroke-width="2" fill="none"/></svg>Cruzamento sem ligação</div>',
        '<div class="lg"><span class="sw" style="width:26px;border:1.5px solid #7FA7D6;background:#EAF2FB;'
        'border-radius:4px"></span>Área</div>',
    ]
    return "\n".join(rows)


def route_panel_html(route, title):
    if not route:
        return ""
    def tags(lst):
        return ", ".join(_html.escape(t) for t in lst) or "—"
    warn = "".join(f'<div class="warn">! {_html.escape(w)}</div>' for w in route["warnings"])
    return f"""
<div class="box">
  <h3>{_html.escape(title)}</h3>
  <div><b>Trechos:</b> {' → '.join(route['segments'])}</div>
  <div class="rv"><b style="color:{ROUTE_COLORS['open']}">■ Abrir ({len(route['open'])}):</b> {tags(route['open'])}</div>
  <div class="rv"><b style="color:{ROUTE_COLORS['close']}">■ Fechar ({len(route['close'])}):</b> {tags(sorted(route['close']))}</div>
  {f'<div class="rv"><b style="color:{ROUTE_COLORS["endpoint"]}">■ Extremidade:</b> {tags(route["endpoint"])}</div>' if route['endpoint'] else ''}
  <div class="lg"><span class="ln" style="border-top:5px solid {ROUTE_COLORS['edge']}"></span>Rota</div>
  {warn}
</div>"""


# Drawing helpers shared by this page and valve_editor.html: area boxes behind the diagram,
# lines drawn with right angles around the items, and a small jump where two unconnected lines cross.
DIAGRAM_JS = r"""
const VD = (() => {
  const AREA_COLORS = [["#EAF2FB", "#7FA7D6", "#2E5C8A"], ["#EDF6E9", "#8DBF7A", "#3D6B2C"], ["#FDF2E4", "#E0A860", "#8A5A12"],
    ["#F3EEFA", "#A88BD0", "#5B3F8A"], ["#FCEDEE", "#D98A90", "#8A2F37"], ["#E9F6F5", "#6FB8B2", "#1F6B65"],
    ["#F6F4E3", "#BDB15C", "#6B6215"], ["#F0F0F0", "#9A9A9A", "#4A4A4A"]];
  const GAP = 5;      // free space kept between a line and an item it doesn't connect to
  const TRACK = 10;   // distance between lines running side by side
  const STUB = 20;    // how far a line leaves an item before turning back (lines going right to left)
  const SHORT = 26;   // first / last pieces shorter than this are avoided

  // [half width, above centre, below centre] of an item with its label
  function extent(n) {
    if (n.bend) return [4, 4, 4];
    if (["box", "database"].includes(n.shape)) return [36, 24, 24];
    if (n.shape === "circle") return [26, 24, 24];
    const s = n.size || 10;
    return [Math.max(s + 2, 20), s + 2, s + 17];
  }
  function point(p) { return !p ? null : Array.isArray(p) ? { x: p[0], y: p[1] } : { x: p.x, y: p.y }; }

  // Pieces of a polyline as axis-aligned segments {h, c, lo, hi}: h = horizontal, c = its y (or x)
  function pieces(pts) {
    const out = [];
    for (let i = 1; i < pts.length; i++) {
      const a = pts[i - 1], b = pts[i];
      if (Math.abs(a.x - b.x) < 0.5 && Math.abs(a.y - b.y) < 0.5) continue;
      const h = Math.abs(a.y - b.y) <= Math.abs(a.x - b.x);
      out.push(h ? { h, c: (a.y + b.y) / 2, lo: Math.min(a.x, b.x), hi: Math.max(a.x, b.x) }
                 : { h, c: (a.x + b.x) / 2, lo: Math.min(a.y, b.y), hi: Math.max(a.y, b.y) });
    }
    return out;
  }
  function simplify(pts) {  // drop repeated points and corners that aren't corners
    const out = [];
    for (const p of pts) {
      const q = out[out.length - 1];
      if (q && Math.abs(q.x - p.x) < 0.5 && Math.abs(q.y - p.y) < 0.5) continue;
      out.push(p);
      while (out.length >= 3) {
        const [a, b, c] = out.slice(-3);
        if ((Math.abs(a.x - b.x) < 0.5 && Math.abs(b.x - c.x) < 0.5) || (Math.abs(a.y - b.y) < 0.5 && Math.abs(b.y - c.y) < 0.5)) out.splice(-2, 1);
        else break;
      }
    }
    return out;
  }

  // Right-angle routes for the edges: {edge id: [points]}. ortho_routes() in valve_visualizer.py
  // repeats this to rate layouts: change both together. Each line is tried straight, as an L,
  // as a Z through a gap between columns and as a detour along a free horizontal track; the route
  // that runs over no item, doesn't lie on top of another line and crosses the fewest wins.
  // keep: routes of other edges that stay as they are (only `edges` are routed).
  function orthoRoutes(nodes, edges, getPos, keep) {
    const P = {}, boxes = [], labelBelow = {};
    for (const n of nodes) {
      if (n.bend) continue;
      const p = point(getPos(n.id));
      if (!p) continue;
      P[n.id] = p;
      labelBelow[n.id] = !["box", "database", "circle"].includes(n.shape);
      const [w, up, down] = extent(n);
      boxes.push({ id: n.id, x1: p.x - w, x2: p.x + w, y1: p.y - up, y2: p.y + down });
    }
    // gaps between columns of items: where vertical runs go
    const xs = Object.values(P).map(p => p.x).sort((a, b) => a - b), cols = [];
    for (const x of xs) {
      if (cols.length && x - cols[cols.length - 1][1] <= 20) cols[cols.length - 1][1] = x;
      else cols.push([x, x]);
    }
    const channels = [];
    for (let i = 1; i < cols.length; i++) channels.push((cols[i - 1][1] + cols[i][0]) / 2);

    const placed = [];  // segments already drawn: {h, c, lo, hi, id, ends}
    const routes = {};
    for (const [id, pts] of Object.entries(keep || {})) {
      routes[id] = pts;
      for (const s of pieces(pts)) placed.push({ ...s, id, ends: pts.ends || [] });
    }
    const todo = edges.filter(e => P[e.from] && P[e.to] && e.from !== e.to);
    const span = e => Math.abs(P[e.to].x - P[e.from].x) + Math.abs(P[e.to].y - P[e.from].y);
    todo.sort((e1, e2) => span(e1) - span(e2) || (e1.id < e2.id ? -1 : 1));

    for (const e of todo) {
      const a = P[e.from], b = P[e.to], dx = b.x - a.x, dy = b.y - a.y;
      const lo = Math.min(a.x, b.x), hi = Math.max(a.x, b.x);
      const after = channels.find(c => c > a.x + 1), before = [...channels].reverse().find(c => c < b.x - 1);
      const x1 = after !== undefined ? after : a.x + STUB, x2 = before !== undefined ? before : b.x - STUB;
      // items around this line, and free horizontal tracks between them
      const near = boxes.filter(o => o.id !== e.from && o.id !== e.to && o.x2 > Math.min(lo, x2) - TRACK && o.x1 < Math.max(hi, x1) + TRACK);
      const bands = near.map(o => [o.y1 - GAP, o.y2 + GAP]).sort((p, q) => p[0] - q[0]), merged = [];
      for (const [t, u] of bands) {
        if (merged.length && t <= merged[merged.length - 1][1]) merged[merged.length - 1][1] = Math.max(merged[merged.length - 1][1], u);
        else merged.push([t, u]);
      }
      const tracks = [];
      if (merged.length) {
        tracks.push(merged[0][0] - TRACK, merged[merged.length - 1][1] + TRACK);
        for (let i = 1; i < merged.length; i++) if (merged[i][0] - merged[i - 1][1] >= 2 * TRACK) tracks.push((merged[i - 1][1] + merged[i][0]) / 2);
      }
      const mid = (a.y + b.y) / 2;
      tracks.sort((p, q) => Math.abs(p - mid) - Math.abs(q - mid));
      const ys = tracks.slice(0, 4);

      const options = [];
      if (Math.abs(dy) < 1 || Math.abs(dx) < 1) options.push([a, b]);
      else options.push([a, { x: b.x, y: a.y }, b], [a, { x: a.x, y: b.y }, b]);
      let ch = channels.filter(c => c > lo + 1 && c < hi - 1);
      if (!ch.length && hi - lo > 1) ch = [(lo + hi) / 2];
      if (ch.length > 3) ch = [ch[0], ch[Math.floor(ch.length / 2)], ch[ch.length - 1]];
      if (Math.abs(dy) >= 1)
        for (const c of ch) for (const k of [0, -1, 1, -2, 2]) {
          const xm = c + k * TRACK;
          options.push([a, { x: xm, y: a.y }, { x: xm, y: b.y }, b]);
        }
      for (const ym of ys) {
        options.push([a, { x: a.x, y: ym }, { x: b.x, y: ym }, b]);
        if (dx < 0 || x1 < x2) options.push([a, { x: x1, y: a.y }, { x: x1, y: ym }, { x: x2, y: ym }, { x: x2, y: b.y }, b]);
      }

      const ends = [e.from, e.to];
      let best = null;
      for (const raw of options) {
        const pts = simplify(raw), segs = pieces(pts), m = pts.length;
        let cost = 0.6 * (m - 2), length = 0;
        // leaving downwards / arriving from below runs through the item's tag label
        if (labelBelow[e.from] && Math.abs(pts[1].x - pts[0].x) < 0.5 && pts[1].y > pts[0].y) cost += 1.5;
        if (labelBelow[e.to] && Math.abs(pts[m - 2].x - pts[m - 1].x) < 0.5 && pts[m - 2].y > pts[m - 1].y) cost += 1.5;
        // a very short first / last piece looks like a stub and turns the arrowhead the wrong way
        const len = (p, q) => Math.abs(p.x - q.x) + Math.abs(p.y - q.y);
        if (m > 2 && len(pts[0], pts[1]) < SHORT) cost += 2;
        if (m > 2 && len(pts[m - 2], pts[m - 1]) < SHORT) cost += 2;
        for (const s of segs) {
          length += s.hi - s.lo;
          for (const o of near) {
            const hit = s.h ? s.c > o.y1 - GAP && s.c < o.y2 + GAP && s.hi > o.x1 - GAP && s.lo < o.x2 + GAP
                            : s.c > o.x1 - GAP && s.c < o.x2 + GAP && s.hi > o.y1 - GAP && s.lo < o.y2 + GAP;
            if (hit) cost += 100;
          }
          for (const r of placed) {
            const shares = r.ends.includes(e.from) || r.ends.includes(e.to);
            if (s.h === r.h) {
              if (Math.abs(s.c - r.c) < TRACK - 2 && Math.min(s.hi, r.hi) - Math.max(s.lo, r.lo) > 2) cost += shares ? 3 : 12;
            } else {
              const hz = s.h ? s : r, vt = s.h ? r : s;
              if (vt.c > hz.lo + 0.5 && vt.c < hz.hi - 0.5 && hz.c > vt.lo + 0.5 && hz.c < vt.hi - 0.5) cost += 1;
            }
          }
          if (best && cost >= best.cost) break;
        }
        cost += 0.004 * length;
        if (!best || cost < best.cost) best = { cost, pts };
      }
      best.pts.ends = ends;
      routes[e.id] = best.pts;
      for (const s of pieces(best.pts)) placed.push({ ...s, id: e.id, ends });
    }
    return routes;
  }

  // vis items for right-angle lines: each edge becomes pieces joined at invisible bend points.
  // Piece ids are the edge id, then "<edge id>~1", "~2"...; bend point ids "<edge id>~b1"...
  function orthoItems(nodes, edges, routes) {
    const bends = [], out = [];
    for (const e of edges) {
      const pts = routes[e.id];
      if (!pts || pts.length <= 2) { out.push({ ...e, smooth: { enabled: false } }); continue; }
      const ids = [e.from];
      for (let k = 1; k < pts.length - 1; k++) {
        const id = `${e.id}~b${k}`;
        ids.push(id);
        bends.push({ id, x: pts[k].x, y: pts[k].y, bend: true, shape: "dot", size: 0.5, fixed: true, physics: false,
          chosen: false, label: undefined, borderWidth: 0, color: { background: "rgba(0,0,0,0)", border: "rgba(0,0,0,0)" } });
      }
      ids.push(e.to);
      let labelAt = 0, longest = -1;  // the label goes on the longest horizontal piece
      for (let k = 0; k < pts.length - 1; k++) {
        const len = Math.abs(pts[k + 1].x - pts[k].x) * 2 + Math.abs(pts[k + 1].y - pts[k].y);
        if (len > longest) { longest = len; labelAt = k; }
      }
      const arrows = e.arrows || {};
      for (let k = 0; k < ids.length - 1; k++) {
        out.push({ ...e, id: k ? `${e.id}~${k}` : e.id, from: ids[k], to: ids[k + 1], smooth: { enabled: false },
          label: k === labelAt ? e.label : "",
          arrows: { to: { ...(arrows.to || {}), enabled: k === ids.length - 2 && !!(arrows.to && arrows.to.enabled) },
                    from: { ...(arrows.from || {}), enabled: k === 0 && !!(arrows.from && arrows.from.enabled) } } });
      }
    }
    return { nodes: bends, edges: out };
  }
  function isBend(id) { return typeof id === "string" && id.includes("~b"); }
  function edgeOfBend(id) { return id.slice(0, id.lastIndexOf("~b")); }

  // Keeps the right-angle lines of a network up to date. build() returns the vis items for new
  // data (positions from getPos); reroute(moved) re-draws after items were dragged: only the lines
  // of the moved items while dragging, all of them when moved is null.
  function orthoRouter(network, nodesDS, edgesDS) {
    let nodes = [], edges = [], routes = {};
    function push() {
      const items = orthoItems(nodes, edges, routes);
      const bendIds = new Set(items.nodes.map(n => n.id));
      nodesDS.remove(nodesDS.getIds().filter(id => isBend(id) && !bendIds.has(id)));
      nodesDS.update(items.nodes);
      const ids = new Set(items.edges.map(e => e.id));
      edgesDS.remove(edgesDS.getIds().filter(id => !ids.has(id)));
      edgesDS.update(items.edges);
    }
    return {
      build(n, e, getPos) {
        nodes = n.filter(x => !x.bend); edges = e;
        routes = orthoRoutes(nodes, edges, getPos);
        return orthoItems(nodes, edges, routes);
      },
      reroute(moved) {
        const pos = network.getPositions(nodes.map(n => n.id)), get = id => pos[id];
        if (!moved) routes = orthoRoutes(nodes, edges, get);
        else {
          const set = new Set(moved), mine = edges.filter(e => set.has(e.from) || set.has(e.to)), keep = {};
          for (const e of edges) if (!set.has(e.from) && !set.has(e.to) && routes[e.id]) keep[e.id] = routes[e.id];
          routes = orthoRoutes(nodes, mine, get, keep);
        }
        push();
      },
    };
  }

  function roundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath();
  }

  // Draw area boxes before the diagram and line jumps after it. opts: {nodes: DataSet, background}
  function install(network, opts) {
    const st = { areas: true, hops: true, boxes: [] };
    const bg = opts.background || "#fafafa";

    network.on("beforeDrawing", ctx => {
      st.boxes = [];
      if (!st.areas) return;
      const byArea = {}, body = network.body;
      for (const n of opts.nodes.get()) if (n.area) (byArea[n.area] = byArea[n.area] || []).push(n);
      const names = Object.keys(byArea).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      names.forEach((name, i) => {
        let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
        for (const n of byArea[name]) {
          const p = body.nodes[n.id];
          if (!p) continue;
          const [w, up, down] = extent(n);
          l = Math.min(l, p.x - w); t = Math.min(t, p.y - up); r = Math.max(r, p.x + w); b = Math.max(b, p.y + down);
        }
        if (l === Infinity) return;
        const [fill, line, text] = AREA_COLORS[i % AREA_COLORS.length];
        const box = { x: l - 22, y: t - 42, w: r - l + 44, h: b - t + 58, fill };
        st.boxes.push(box);
        ctx.save();
        roundRect(ctx, box.x, box.y, box.w, box.h, 12);
        ctx.fillStyle = fill; ctx.fill();
        ctx.lineWidth = 1.5; ctx.strokeStyle = line; ctx.stroke();
        ctx.fillStyle = text; ctx.font = "bold 16px Arial"; ctx.textBaseline = "top";
        ctx.fillText(/^\d/.test(name) ? "Área " + name : name, box.x + 12, box.y + 10);
        ctx.restore();
      });
    });

    network.on("afterDrawing", ctx => {
      if (!st.hops) return;
      const body = network.body, segs = [];
      for (const id of body.edgeIndices) {
        const e = body.edges[id];
        if (!e || !e.connected || !e.from || !e.to || e.options.hidden) continue;
        const o = e.options, color = (o.color && o.color.color) || "#7F7F7F", width = o.width || 2;
        let pts;
        if (o.smooth && o.smooth.enabled && e.edgeType && e.edgeType.getPoint) {
          pts = [];
          for (let i = 0; i <= 14; i++) pts.push(e.edgeType.getPoint(i / 14));
        } else pts = [{ x: e.from.x, y: e.from.y }, { x: e.to.x, y: e.to.y }];
        for (let i = 1; i < pts.length; i++) {
          const a = pts[i - 1], b = pts[i];
          segs.push({ id, a, b, color, width, ends: [e.fromId, e.toId],
            minx: Math.min(a.x, b.x), maxx: Math.max(a.x, b.x), miny: Math.min(a.y, b.y), maxy: Math.max(a.y, b.y) });
        }
      }
      const inside = (p, ids) => ids.some(id => {
        const bb = network.getBoundingBox(id);
        return bb && p.x > bb.left - 3 && p.x < bb.right + 3 && p.y > bb.top - 3 && p.y < bb.bottom + 3;
      });
      const hops = [];
      for (let i = 0; i < segs.length; i++) {
        const s = segs[i];
        for (let j = i + 1; j < segs.length; j++) {
          const q = segs[j];
          if (s.id === q.id || q.minx > s.maxx || q.maxx < s.minx || q.miny > s.maxy || q.maxy < s.miny) continue;
          if (s.ends.some(n => q.ends.includes(n))) continue;  // lines meeting at a node are connected
          const rx = s.b.x - s.a.x, ry = s.b.y - s.a.y, sx = q.b.x - q.a.x, sy = q.b.y - q.a.y;
          const den = rx * sy - ry * sx;
          if (Math.abs(den) < 1e-9) continue;
          const t = ((q.a.x - s.a.x) * sy - (q.a.y - s.a.y) * sx) / den;
          const u = ((q.a.x - s.a.x) * ry - (q.a.y - s.a.y) * rx) / den;
          if (t < 0 || t > 1 || u < 0 || u > 1) continue;
          const p = { x: s.a.x + t * rx, y: s.a.y + t * ry };
          if (inside(p, [...s.ends, ...q.ends])) continue;
          // the more horizontal line jumps over the other one
          const sh = Math.abs(rx) / Math.hypot(rx, ry), qh = Math.abs(sx) / Math.hypot(sx, sy);
          hops.push(sh >= qh ? { p, over: s, under: q } : { p, over: q, under: s });
        }
      }
      const bgAt = p => {
        for (let k = st.boxes.length - 1; k >= 0; k--) {
          const b = st.boxes[k];
          if (p.x >= b.x && p.x <= b.x + b.w && p.y >= b.y && p.y <= b.y + b.h) return b.fill;
        }
        return bg;
      };
      st.crossings = hops.length;
      const unit = s => { const dx = s.b.x - s.a.x, dy = s.b.y - s.a.y, l = Math.hypot(dx, dy); return { x: dx / l, y: dy / l }; };
      ctx.save();
      ctx.lineCap = "butt";
      for (const { p, over, under } of hops) {
        let u = unit(over);
        if (u.x < 0 || (u.x === 0 && u.y > 0)) u = { x: -u.x, y: -u.y };
        const v = unit(under), r = 7 + over.width;
        ctx.strokeStyle = bgAt(p); ctx.lineWidth = over.width + 4;  // gap in the line that jumps
        ctx.beginPath(); ctx.moveTo(p.x - r * u.x, p.y - r * u.y); ctx.lineTo(p.x + r * u.x, p.y + r * u.y); ctx.stroke();
        const k = r + 3;  // the other line stays continuous
        ctx.strokeStyle = under.color; ctx.lineWidth = under.width;
        ctx.beginPath(); ctx.moveTo(p.x - k * v.x, p.y - k * v.y); ctx.lineTo(p.x + k * v.x, p.y + k * v.y); ctx.stroke();
        const th = Math.atan2(u.y, u.x);
        ctx.strokeStyle = over.color; ctx.lineWidth = over.width;
        ctx.beginPath(); ctx.arc(p.x, p.y, r, th + Math.PI, th + 2 * Math.PI); ctx.stroke();
      }
      ctx.restore();
    });
    return st;
  }

  return { orthoRoutes, orthoItems, orthoRouter, isBend, edgeOfBend, install };
})();
"""

VIS_CDN = "https://cdn.jsdelivr.net/npm/vis-network@9.1.9/standalone/umd/vis-network.min.js"

PAGE = r"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
__LIB__
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; height: 100%; font-family: Arial, Helvetica, sans-serif; color: #1f1f1f; background: #fafafa; }
  #wrap { display: flex; height: 100%; }
  #side { width: 270px; min-width: 270px; overflow-y: auto; padding: 12px; border-right: 1px solid #ddd; background: #fff; font-size: 13px; }
  #main { flex: 1; display: flex; flex-direction: column; min-width: 0; }
  #bar { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; padding: 8px; border-bottom: 1px solid #ddd; background: #fff; }
  #net { flex: 1; min-height: 300px; }
  h2 { font-size: 16px; margin: 0 0 8px; } h3 { font-size: 14px; margin: 0 0 6px; }
  .box { border: 1px solid #e3e3e3; border-radius: 6px; padding: 8px; margin-bottom: 10px; }
  .lg { display: flex; align-items: center; gap: 6px; margin: 4px 0; }
  .sw { display: inline-block; width: 16px; height: 16px; border: 2px solid #404040; flex: none; }
  .sw.dot { width: 9px; height: 9px; border-radius: 50%; margin: 0 3.5px; }
  .sw.circle { border-radius: 50%; }
  .sw.database { border-radius: 4px / 6px; }
  .sw.diamond { transform: rotate(45deg) scale(.8); }
  .sw.hexagon { clip-path: polygon(25% 5%, 75% 5%, 100% 50%, 75% 95%, 25% 95%, 0 50%); border-width: 0; }
  .sw.triangle, .sw.triangleDown { width: 0; height: 0; border: 8px solid transparent; border-bottom: 14px solid #404040; border-top: 0; }
  .sw.triangleDown { border-bottom: 0; border-top: 14px solid #5E3B1C; }
  .ar { display: inline-block; width: 34px; color: #7F7F7F; font-size: 11px; letter-spacing: -1px; flex: none; }
  .rv { margin: 5px 0; }
  .ln { display: inline-block; width: 26px; border-top: 2px solid #7F7F7F; flex: none; }
  .warn { color: #a33; margin-top: 4px; }
  button, input { font: inherit; font-size: 13px; padding: 4px 8px; }
  button { cursor: pointer; border: 1px solid #bbb; background: #f4f4f4; border-radius: 4px; }
  button.on { background: #1F3864; color: #fff; border-color: #1F3864; }
  .vis-tooltip { white-space: pre-line; font-family: Arial, sans-serif !important; font-size: 12px !important; }
  #msg { color: #a33; font-size: 13px; }
</style></head>
<body><div id="wrap">
  <div id="side">
    <h2>__TITLE__</h2>
    __ROUTE__
    <div class="box"><h3>Legenda</h3>__LEGEND__</div>
    <div style="color:#777;font-size:12px">Passe o mouse sobre um item para ver detalhes. Arraste para reposicionar; role para zoom.<br>Onde duas linhas se cruzam sem ligação, uma delas faz um salto (⌒).<br>Gerado em __DATE__</div>
  </div>
  <div id="main">
    <div id="bar">
      <button id="bA" title="Cada área num bloco, fluxo da esquerda para a direita">Por área</button>
      <button id="bF">Livre</button>
      <button id="bFit">Ajustar à tela</button>
      <input id="q" list="ids" placeholder="Buscar tag…" style="width:150px">
      <datalist id="ids"></datalist>
      <button id="bQ">Localizar</button>
      <span id="msg"></span>
    </div>
    <div id="net"></div>
  </div>
</div>
<script>__DIAGRAM_JS__</script>
<script>
const DATA = __DATA__;
const START_LAYOUT = "__LAYOUT__";
function tipEl(t) { const d = document.createElement("div"); d.textContent = t; d.style.whiteSpace = "pre-line"; return d; }
DATA.nodes.forEach(n => { n.title = tipEl(n.tip); });
DATA.edges.forEach(e => { e.title = tipEl(e.tip); });
document.getElementById("ids").innerHTML = DATA.nodes.map(n => `<option value="${n.id}">`).join("");
let net = null, nodes = null, edges = null, view = null, router = null, pending = null;
function draw(layout) {
  const fixed = layout !== "free";
  if (net) net.destroy();
  nodes = new vis.DataSet(); edges = new vis.DataSet();
  net = new vis.Network(document.getElementById("net"), { nodes, edges }, {
    layout: { randomSeed: 7 },
    physics: fixed ? false : { barnesHut: { gravitationalConstant: -6000, springLength: 90, avoidOverlap: 0.3 }, stabilization: { iterations: 600 } },
    interaction: { hover: true, tooltipDelay: 120, navigationButtons: false, keyboard: true },
  });
  view = VD.install(net, { nodes, background: "#fafafa" });
  view.areas = fixed;
  if (fixed) {  // fixed positions, lines with right angles
    router = VD.orthoRouter(net, nodes, edges);
    const items = router.build(DATA.nodes, DATA.edges, id => { const n = DATA.nodes.find(m => m.id === id); return n && { x: n.x, y: n.y }; });
    nodes.add([...DATA.nodes, ...items.nodes]);
    edges.add(items.edges);
    net.on("dragging", e => {  // lines follow the items being dragged
      const moved = e.nodes.filter(id => !VD.isBend(id));
      if (!moved.length || pending) return;
      pending = requestAnimationFrame(() => { pending = null; router.reroute(moved); });
    });
    net.on("dragEnd", e => { if (e.nodes.some(id => !VD.isBend(id))) router.reroute(null); });
    net.fit();
  } else {
    nodes.add(DATA.nodes.map(n => { const o = { ...n }; delete o.x; delete o.y; return o; }));
    edges.add(DATA.edges.map(e => ({ ...e, smooth: { type: "continuous" } })));
    net.once("stabilizationIterationsDone", () => { net.setOptions({ physics: false }); net.fit(); });
  }
  document.getElementById("bA").classList.toggle("on", fixed);
  document.getElementById("bF").classList.toggle("on", !fixed);
}
function find() {
  const q = document.getElementById("q").value.trim().toUpperCase();
  const hit = DATA.nodes.find(n => n.id.toUpperCase() === q) || DATA.nodes.find(n => n.id.toUpperCase().includes(q));
  document.getElementById("msg").textContent = hit || !q ? "" : "Tag não encontrado";
  if (hit) { net.selectNodes([hit.id]); net.focus(hit.id, { scale: 1.4, animation: { duration: 500 } }); }
}
document.getElementById("bA").onclick = () => draw("area");
document.getElementById("bF").onclick = () => draw("free");
document.getElementById("bFit").onclick = () => net.fit({ animation: true });
document.getElementById("bQ").onclick = find;
document.getElementById("q").addEventListener("keydown", e => { if (e.key === "Enter") find(); });
draw(START_LAYOUT);
</script></body></html>"""


def render_html(plant, route=None, title="Mapa de válvulas", layout="area", lib_js=None, positions=None):
    """layout: "area" (fixed positions, grouped by area) or "free" (force layout).
    positions: {tag: [x, y]} to use instead of the automatic area layout (e.g. saved by valve_editor.py)."""
    import datetime as _dt
    data = build_diagram(plant, route)
    auto = area_layout(plant)
    for node in data["nodes"]:
        x, y = (positions or {}).get(node["id"]) or auto.get(node["id"], (0, 0))
        node["x"], node["y"] = x, y
    lib = (f"<script>{lib_js}</script>" if lib_js else f'<script src="{VIS_CDN}"></script>')
    route_title = ""
    if route:
        loop = route["path"][0] == route["path"][-1]
        route_title = (f"Recirculação {route['path'][0]}" if loop else f"Rota {route['path'][0]} → {route['path'][-1]}") + (f" ({route['product']})" if route.get("product") else "")
    page = PAGE
    for key, val in {
        "__DIAGRAM_JS__": DIAGRAM_JS,
        "__LIB__": lib,
        "__ROUTE__": route_panel_html(route, route_title),
        "__LEGEND__": legend_html(data["kinds"]),
        "__DATE__": _dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "__LAYOUT__": "free" if layout == "free" else "area",
        "__TITLE__": _html.escape(title),
    }.items():
        page = page.replace(key, val)
    data_js = json.dumps({"nodes": data["nodes"], "edges": data["edges"]}, ensure_ascii=False).replace("</", "<\\/")
    return page.replace("__DATA__", data_js)
# --- END VISUALIZER CORE ---


def saved_positions(path):
    """Diagram positions valve_editor.py keeps in the hidden 'Layout do Diagrama' sheet."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    name = next((n for n in ("Layout do Diagrama", "Diagram Layout") if n in wb.sheetnames), None)
    out = {}
    if name:
        for row in wb[name].iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 3 and row[0] and isinstance(row[1], (int, float)) and isinstance(row[2], (int, float)):
                out[str(row[0])] = [row[1], row[2]]
    wb.close()
    return out


def main():
    import argparse
    import os
    import sys
    import webbrowser
    try:
        import networkx as nx
        from valve_router import Plant
    except ImportError as e:
        sys.exit(f"{e}\nPut valve_router.py in the same folder and run: pip install openpyxl networkx")

    ap = argparse.ArgumentParser(description="Draw the valve network as an interactive HTML diagram.")
    ap.add_argument("workbook")
    ap.add_argument("-o", "--output", default="valve_diagram.html")
    ap.add_argument("--from", dest="start", help="highlight the best route from this point")
    ap.add_argument("--to", dest="end")
    ap.add_argument("--product")
    ap.add_argument("--avoid", default="")
    ap.add_argument("--valve-penalty", type=float, default=0.0)
    ap.add_argument("--layout", choices=["area", "saved", "free", "flow"], default="area",
                    help="area = one block per area, flow left to right, few crossings (default); "
                         "saved = positions saved by valve_editor.py, area layout for the rest; "
                         "free = force layout. 'flow' is the old name of 'area'.")
    ap.add_argument("--title", default="Mapa de válvulas")
    ap.add_argument("--offline-lib", metavar="vis-network.min.js",
                    help="embed a local copy of vis-network so the page works without internet")
    ap.add_argument("--open", action="store_true", help="open the diagram in the browser")
    args = ap.parse_args()

    plant = Plant(args.workbook)
    if plant.errors:
        for e in plant.errors:
            print("ERROR: " + e)
        sys.exit("Fix the database errors first (python valve_router.py <file> --check).")

    route = None
    if args.start or args.end:
        if not (args.start and args.end):
            sys.exit("Use --from and --to together.")
        avoid = [a.strip() for a in args.avoid.split(",") if a.strip()]
        G = plant.build_graph(args.product, avoid, args.valve_penalty, endpoints=(args.start, args.end))
        found = plant.find_routes(G, args.start, args.end, 1)
        if not found:
            sys.exit(f"No valid route from {args.start} to {args.end}. Run valve_router.py for the reason.")
        path = found[0]
        route = route_overlay(plant, G, path)
        route["product"] = args.product

    lib_js = None
    if args.offline_lib:
        with open(args.offline_lib, encoding="utf-8") as f:
            lib_js = f.read()

    positions = None
    if args.layout == "saved":
        positions = saved_positions(args.workbook)
        if not positions:
            print("No saved positions in the workbook (sheet 'Layout do Diagrama'); using the area layout.")

    page = render_html(plant, route, args.title, args.layout, lib_js, positions)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(page)
    print(f"Diagram saved to {os.path.abspath(args.output)}")
    if route:
        print(f"Route: {' -> '.join(route['path'])}")
        print(f"Open {len(route['open'])} valve(s), close {len(route['close'])} valve(s).")
    if args.open:
        webbrowser.open("file://" + os.path.abspath(args.output))


if __name__ == "__main__":
    main()

# Offline use: download https://cdn.jsdelivr.net/npm/vis-network@9.1.9/standalone/umd/vis-network.min.js
# once, save it next to this script, and run with --offline-lib vis-network.min.js
