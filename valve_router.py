#!/usr/bin/env python3
"""
valve_router.py - find the best route between two points in a piping network
and work out which valves must be OPEN and CLOSED for that operation.

Reads the workbook produced from the valve_database.xlsx template
(Portuguese sheets: Nós, Trechos, Válvulas, Produtos, Matriz de Válvulas;
the older English template with Nodes, Segments, Valves, Products, Valve Matrix also works).

Usage
    python valve_router.py valve_database.xlsx                     # interactive
    python valve_router.py valve_database.xlsx --from TQ-101 --to RE-202 --product Leite
    python valve_router.py valve_database.xlsx --from TQ-101 --to RE-202 --product Leite \
           --alternatives 3 --avoid TC-101,VES-104 --save "Leite TQ-101 para RE-202"
    python valve_router.py valve_database.xlsx --check

Requirements: pip install openpyxl networkx
"""
import argparse
import datetime as dt
import sys

try:
    import networkx as nx
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    sys.exit("Missing packages. Run:  pip install openpyxl networkx")

# Valve statuses that block flow / prevent isolation (canonical, after translation)
BLOCKING_STATUS = {"out of service", "locked closed"}
NO_ISOLATION_STATUS = {"locked open", "out of service"}
# Node types whose default "pass-through" is No (English and Portuguese)
NON_PASS_TYPES = {"tank", "drain", "reactor", "mixer", "cip",
                  "tanque", "dreno", "reator", "misturador", "central cip"}

# The workbook can be in Portuguese or English. Sheets and columns are found by any of these names.
SHEETS = {
    "nodes": ["Nós", "Nos", "Nodes"],
    "segments": ["Trechos", "Segments"],
    "valves": ["Válvulas", "Valvulas", "Valves"],
    "products": ["Produtos", "Products"],
    "matrix": ["Matriz de Válvulas", "Matriz de Valvulas", "Valve Matrix"],
}
COLUMNS = {
    "node_id": ["Tag do Nó", "Tag do No", "Node ID"],
    "node_type": ["Tipo de Nó", "Tipo de No", "Node Type"],
    "desc": ["Descrição", "Descricao", "Description"],
    "pass": ["Permite Passagem", "Pass-through Allowed"],
    "status": ["Status", "Situação", "Situacao"],
    "notes": ["Observações", "Observacoes", "Notes"],
    "product": ["Produto", "Product"],
    "seg_id": ["Tag do Trecho", "Segment ID"],
    "from": ["Origem", "From"],
    "to": ["Destino", "To"],
    "direction": ["Sentido", "Direction"],
    "products": ["Produtos Permitidos", "Allowed Products"],
    "cost": ["Custo", "Cost"],
    "valve_tag": ["Tag da Válvula", "Tag da Valvula", "Valve Tag"],
    "segment": ["Trecho", "Segment"],
    "order": ["Posição no Trecho", "Posicao no Trecho", "Order on Segment"],
    "body": ["Tipo de Corpo"],
    "actuation": ["Acionamento", "Valve Type"],
    "area": ["Área", "Area"],
}
# Cell values -> canonical English values used internally
VALUE_MAP = {
    "sim": "yes", "não": "no", "nao": "no",
    "em operação": "in service", "em operacao": "in service",
    "fora de operação": "out of service", "fora de operacao": "out of service",
    "travada fechada": "locked closed", "travada aberta": "locked open",
    "unidirecional": "one-way", "bidirecional": "bidirectional",
    "automática": "automatic", "automatica": "automatic", "retenção": "check", "retencao": "check",
}
CHECK_BODIES = {"retenção", "retencao", "vrt", "check"}


def norm(v, default=""):
    """Lower-case canonical value (Portuguese values translated to English)."""
    t = text(v, default).lower()
    return VALUE_MAP.get(t, t)


def find_sheet(wb, key, required=True):
    for name in SHEETS[key]:
        if name in wb.sheetnames:
            return name
    if required:
        sys.exit(f"Sheet '{SHEETS[key][0]}' not found in workbook (also tried: {', '.join(SHEETS[key][1:])}).")
    return None


# --------------------------------------------------------------------------- #
#  Reading the workbook
# --------------------------------------------------------------------------- #
def read_sheet(wb, key, required, optional=()):
    """Return list of dicts keyed by canonical column keys. Rows with an empty first column are skipped."""
    name = find_sheet(wb, key)
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    col = {}
    for k in list(required) + list(optional):
        for alias in COLUMNS[k]:
            if alias.lower() in headers:
                col[k] = headers.index(alias.lower())
                break
    missing = [COLUMNS[k][0] for k in required if k not in col]
    if missing:
        sys.exit(f"Sheet '{name}' is missing column(s): {', '.join(missing)}")
    out = []
    for i, row in enumerate(rows[1:], start=2):
        rec = {k: (row[c].strip() if isinstance(row[c], str) else row[c]) if c < len(row) else None
               for k, c in col.items()}
        if rec.get(required[0]) in (None, ""):
            continue
        rec["_row"] = i
        out.append(rec)
    return name, out


def text(v, default=""):
    return default if v is None else str(v).strip()


def area_text(v):
    """Area cell -> text, so 100, 100.0 and '100' are the same area."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return text(v)


# Columns read from each sheet: (required, optional). The first required column identifies a row.
TABLE_SPECS = {
    "nodes": (["node_id", "node_type"], ["desc", "pass", "status", "area"]),
    "products": (["product"], []),
    "segments": (["seg_id", "from", "to", "direction"], ["products", "cost", "status"]),
    "valves": (["valve_tag", "segment", "order"], ["body", "actuation", "status", "notes"]),
}


def read_tables(wb):
    """{key: (sheet name, rows)} for every table the Plant needs."""
    return {key: read_sheet(wb, key, req, opt) for key, (req, opt) in TABLE_SPECS.items()}


class Plant:
    def __init__(self, path, tables=None):
        """Reads the workbook at path. Pass tables (same shape as read_tables returns) to
        build the plant from data held in memory instead, e.g. unsaved edits in valve_editor.py."""
        self.path = path
        if tables is None:
            tables = read_tables(load_workbook(path, data_only=True))
        self.errors, self.warnings = [], []
        self.portuguese = tables["valves"][0] != "Valves"

        # Nodes
        self.nodes = {}
        sh, rows = tables["nodes"]
        for r in rows:
            nid = text(r["node_id"])
            if nid in self.nodes:
                self.errors.append(f"{sh} row {r['_row']}: duplicate node tag '{nid}'")
            ntype = text(r.get("node_type")).lower()
            pt = norm(r.get("pass"))
            passthrough = (pt == "yes") if pt in ("yes", "no") else (ntype not in NON_PASS_TYPES)
            self.nodes[nid] = {
                "type": text(r.get("node_type")),
                "desc": text(r.get("desc")),
                "pass": passthrough,
                "active": norm(r.get("status"), "in service") != "out of service",
                "area": area_text(r.get("area")),
            }

        # Products
        _, rows = tables["products"]
        self.products = {text(r["product"]) for r in rows}

        # Segments
        self.segments = {}
        sh, rows = tables["segments"]
        for r in rows:
            sid = text(r["seg_id"])
            if sid in self.segments:
                self.errors.append(f"{sh} row {r['_row']}: duplicate segment tag '{sid}'")
            frm, to = text(r["from"]), text(r["to"])
            for end in (frm, to):
                if end not in self.nodes:
                    self.errors.append(f"{sh} row {r['_row']} ({sid}): unknown node '{end}'")
            direction = norm(r.get("direction"))
            if direction not in ("one-way", "bidirectional"):
                self.warnings.append(f"Segment {sid}: direction '{r.get('direction')}' not recognised, treated as one-way")
                direction = "one-way"
            prods_raw = text(r.get("products"))
            prods = {p.strip() for p in prods_raw.split(",") if p.strip()} if prods_raw else None
            if prods:
                unknown = prods - self.products
                if unknown:
                    self.warnings.append(f"Segment {sid}: product(s) not in the products sheet: {', '.join(sorted(unknown))}")
            cost = r.get("cost")
            try:
                cost = float(cost) if cost not in (None, "") else 1.0
            except (TypeError, ValueError):
                self.warnings.append(f"Segment {sid}: cost '{cost}' is not a number, using 1")
                cost = 1.0
            self.segments[sid] = {
                "from": frm, "to": to,
                "bidir": direction == "bidirectional",
                "products": prods,
                "cost": cost,
                "active": norm(r.get("status"), "in service") != "out of service",
                "valves": [],
            }

        # Valves
        self.valves = {}
        sh, rows = tables["valves"]
        for r in rows:
            tag = text(r["valve_tag"])
            if tag in self.valves:
                self.errors.append(f"{sh} row {r['_row']}: duplicate valve tag '{tag}'")
            if tag in self.nodes:
                self.errors.append(f"Valve tag '{tag}' is also used as a node tag")
            sid = text(r["segment"])
            if sid not in self.segments:
                self.errors.append(f"{sh} row {r['_row']} ({tag}): unknown segment '{sid}'")
                continue
            try:
                order = float(r.get("order"))
            except (TypeError, ValueError):
                self.warnings.append(f"Valve {tag}: missing position on segment, placed last")
                order = 999
            body = text(r.get("body"))
            act = norm(r.get("actuation"), "manual")
            # canonical type used by the router: Check / Manual / Automatic / Control
            if body.lower().split(" ")[0] in CHECK_BODIES or act == "check":
                vtype = "Check"
            elif act in ("automatic", "control"):
                vtype = act.capitalize()
            else:
                vtype = "Manual"
            self.valves[tag] = {
                "segment": sid,
                "order": order,
                "type": vtype,
                "body": body,
                "status": norm(r.get("status"), "in service"),
                "notes": text(r.get("notes")),
            }
            self.segments[sid]["valves"].append(tag)

        for sid, s in self.segments.items():
            s["valves"].sort(key=lambda t: self.valves[t]["order"])
            orders = [self.valves[t]["order"] for t in s["valves"]]
            if len(orders) != len(set(orders)):
                self.errors.append(f"Segment {sid}: two valves share the same position on the segment")

    # The physical chain of points along a segment, From -> valves -> To
    def chain(self, sid):
        s = self.segments[sid]
        mid = s["valves"] if s["valves"] else [f"~{sid}"]  # hidden point keeps parallel lines distinct
        return [s["from"], *mid, s["to"]]

    def is_valve(self, p):
        return p in self.valves

    def is_hidden(self, p):
        return p.startswith("~")

    def points(self):
        return sorted(self.nodes) + sorted(self.valves)

    # ------------------------------------------------------------------- #
    #  Graph of allowed flow for one operation
    # ------------------------------------------------------------------- #
    def build_graph(self, product=None, avoid=(), valve_penalty=0.0, endpoints=()):
        G = nx.DiGraph()
        avoid = set(avoid)
        for sid, s in self.segments.items():
            if not s["active"] or sid in avoid:
                continue
            if product and s["products"] is not None and product not in s["products"]:
                continue
            ch = self.chain(sid)
            # any blocked valve on this segment cuts the line at that valve
            has_check = any(self.valves[v]["type"].lower() == "check" for v in s["valves"])
            step = s["cost"] / (len(ch) - 1)
            for a, b in zip(ch, ch[1:]):
                if self._blocked(a, avoid) or self._blocked(b, avoid):
                    continue
                w_ab = step + (valve_penalty if self.is_valve(b) else 0)
                w_ba = step + (valve_penalty if self.is_valve(a) else 0)
                G.add_edge(a, b, segment=sid, weight=w_ab)
                if s["bidir"] and not has_check:
                    G.add_edge(b, a, segment=sid, weight=w_ba)
        # nodes that can't be passed through (tanks etc.) are only allowed as endpoints
        for nid, n in self.nodes.items():
            if nid in G and nid not in endpoints and (not n["pass"] or not n["active"] or nid in avoid):
                G.remove_node(nid)
        # Loop / recirculation (start = end): split the point in two. Routes leave from the
        # real point and arrive at a "return" copy that receives all of its incoming lines.
        if len(endpoints) == 2 and endpoints[0] == endpoints[1] and endpoints[0] in G:
            s0, ret = endpoints[0], self.return_point(endpoints[0])
            for u in list(G.predecessors(s0)):
                G.add_edge(u, ret, **G[u][s0])
                G.remove_edge(u, s0)
        return G

    @staticmethod
    def return_point(nid):
        return f"~return~{nid}"

    def find_routes(self, G, start, end, k=1):
        """Up to k best routes (lowest cost first). start == end finds recirculation loops.
        Returns [] when there is no route. Paths use real point names only."""
        target = self.return_point(start) if start == end else end
        if start not in G or target not in G:
            return []
        paths = []
        try:
            for p in nx.shortest_simple_paths(G, start, target, weight="weight"):
                # a loop must not leave and come back through the same piece of pipe
                if start == end and frozenset(p[:2]) == frozenset((p[-2], start)):
                    continue
                paths.append(p)
                if len(paths) >= k:
                    break
        except nx.NetworkXNoPath:
            return []
        if start == end:
            for u in list(G.predecessors(target)):  # so segment/cost lookups work on real names
                G.add_edge(u, start, **G[u][target])
            paths = [p[:-1] + [start] for p in paths]
        return paths

    def _blocked(self, p, avoid):
        if p in avoid:
            return True
        if self.is_valve(p):
            return self.valves[p]["status"] in BLOCKING_STATUS
        if p in self.nodes:
            return not self.nodes[p]["active"]
        return False

    # ------------------------------------------------------------------- #
    #  Valve states for a route
    # ------------------------------------------------------------------- #
    def valve_states(self, path):
        start, end = path[0], path[-1]
        route_edges = {frozenset(e) for e in zip(path, path[1:])}
        on_route = set(path)

        open_v, endpoint_v = [], []
        for p in path:
            if self.is_valve(p):
                (endpoint_v if p in (start, end) else open_v).append(p)

        close_v, warnings = {}, []
        for node in dict.fromkeys(path):  # a loop route visits its start/end point twice
            if self.is_valve(node) or self.is_hidden(node):
                continue
            for sid, s in self.segments.items():
                ch = self.chain(sid)
                for direction in ((ch, s["from"]), (ch[::-1], s["to"])):
                    seq, end_node = direction
                    if end_node != node:
                        continue
                    if frozenset((seq[0], seq[1])) in route_edges:
                        continue  # this side of the segment is part of the route
                    found = None
                    passed_check = []
                    for p in seq[1:-1]:
                        if not self.is_valve(p):
                            continue
                        v = self.valves[p]
                        if v["type"].lower() == "check" or v["status"] in NO_ISOLATION_STATUS:
                            passed_check.append(p)
                            continue
                        found = p
                        break
                    if found and found not in on_route:
                        close_v.setdefault(found, (sid, node))
                    elif not found:
                        other = seq[-1]
                        extra = f" (only {', '.join(passed_check)}, which can't isolate)" if passed_check else ""
                        warnings.append(
                            f"No isolating valve on {sid} between {node} and {other}{extra}: "
                            f"product can reach {other}.")
        return open_v, close_v, endpoint_v, warnings

    def path_segments(self, G, path):
        segs = []
        for a, b in zip(path, path[1:]):
            sid = G[a][b]["segment"]
            if not segs or segs[-1] != sid:
                segs.append(sid)
        return segs

    def path_cost(self, G, path):
        return sum(G[a][b]["weight"] for a, b in zip(path, path[1:]))

    # ------------------------------------------------------------------- #
    #  Check the database
    # ------------------------------------------------------------------- #
    def check(self):
        issues = list(self.errors)
        connected = set()
        for s in self.segments.values():
            connected.update((s["from"], s["to"]))
        for nid in self.nodes:
            if nid not in connected:
                self.warnings.append(f"Node {nid} is not connected to any segment")
        # valve numbers must be unique whatever the body type (VES-101 and VBO-101 can't coexist)
        by_number = {}
        for tag in self.valves:
            num = tag.rsplit("-", 1)[-1]
            if num.isdigit():
                by_number.setdefault(int(num), []).append(tag)
        for num, tags in sorted(by_number.items()):
            if len(tags) > 1:
                self.warnings.append(f"Valve number {num} is used by more than one valve: {', '.join(sorted(tags))}")
        for nid, n in self.nodes.items():
            if n["type"].lower() in ("junction", "manifold", "derivação", "derivacao", "coletor"):
                degree = sum(1 for s in self.segments.values() if nid in (s["from"], s["to"]))
                if degree < 3:
                    self.warnings.append(f"Junction {nid} only has {degree} segment(s); a split/merge normally has 3+")
        return issues, self.warnings


# --------------------------------------------------------------------------- #
#  Output
# --------------------------------------------------------------------------- #
def describe_valve(plant, tag):
    v = plant.valves[tag]
    t = v["type"].lower()
    flag = ""
    if t == "manual":
        flag = "  <-- MANUAL, operator action"
    elif t == "check":
        flag = "  (check valve, passive)"
    kind = v.get("body") or v["type"]
    return f"{tag:<10} on {v['segment']:<8} [{kind}]{flag}"


def print_route(plant, G, path, label):
    segs = plant.path_segments(G, path)
    visible = [p for p in path if not plant.is_hidden(p)]
    open_v, close_v, endpoint_v, warns = plant.valve_states(path)

    print(f"\n=== {label}  (cost {plant.path_cost(G, path):g}) ===")
    print("Path    : " + " -> ".join(visible))
    print("Segments: " + " -> ".join(segs))
    print(f"\nOPEN ({len(open_v)}):")
    for t in open_v:
        print("  " + describe_valve(plant, t))
    print(f"\nCLOSE ({len(close_v)}):  (isolates branches leaving the route)")
    for t, (sid, node) in sorted(close_v.items()):
        print("  " + describe_valve(plant, t) + f"   isolates {sid} at {node}")
    if endpoint_v:
        print("\nENDPOINT valve(s), state depends on the operation: " + ", ".join(endpoint_v))
    if warns:
        print("\nWARNINGS:")
        for w in warns:
            print("  ! " + w)
    return open_v, close_v, endpoint_v, segs


def save_to_matrix(plant, name, product, start, end, segs, open_v, close_v, endpoint_v):
    wb = load_workbook(plant.path)  # formulas preserved
    sheet = find_sheet(wb, "matrix", required=False)
    if sheet is None:
        sheet = "Matriz de Válvulas" if plant.portuguese else "Valve Matrix"
        wb.create_sheet(sheet)
    ws = wb[sheet]
    if plant.portuguese:
        fixed = ["Operação", "Produto", "Origem", "Destino", "Rota (trechos)", "Salvo em"]
    else:
        fixed = ["Operation", "Product", "From", "To", "Route (segments)", "Saved On"]
    for i, h in enumerate(fixed, start=1):
        ws.cell(row=1, column=i, value=h)

    # make sure every valve has a column, in alphabetical order after existing ones
    existing = {ws.cell(row=1, column=c).value: c for c in range(len(fixed) + 1, ws.max_column + 1)
                if ws.cell(row=1, column=c).value}
    col = max(existing.values(), default=len(fixed)) + 1
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", start_color="1F3864")
    for tag in sorted(plant.valves):
        if tag not in existing:
            c = ws.cell(row=1, column=col, value=tag)
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(horizontal="center", textRotation=90)
            ws.column_dimensions[c.column_letter].width = 6
            existing[tag] = col
            col += 1

    # replace a row with the same operation name, otherwise append after the last used row
    row = None
    last = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=3).value:  # rows with a From value are operations
            last = r
            if ws.cell(row=r, column=1).value == name:
                row = r
    if row is None:
        row = max(last + 1, 3)  # row 2 holds the legend in the template

    values = [name, product or ("Qualquer" if plant.portuguese else "Any"), start, end, " > ".join(segs), dt.datetime.now().strftime("%Y-%m-%d %H:%M")]
    for i, v in enumerate(values, start=1):
        ws.cell(row=row, column=i, value=v).font = Font(name="Arial")
    fills = {"O": "C6EFCE", "C": "FFC7CE", "MO": "FFEB9C", "MC": "FFEB9C", "E": "DDEBF7"}
    for tag, c in existing.items():
        manual = plant.valves[tag]["type"].lower() == "manual"
        if tag in open_v:
            state = "MO" if manual else "O"
        elif tag in close_v:
            state = "MC" if manual else "C"
        elif tag in endpoint_v:
            state = "E"
        else:
            state = None
        cell = ws.cell(row=row, column=c, value=state)
        cell.font = Font(name="Arial", bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill("solid", start_color=fills[state]) if state else PatternFill(fill_type=None)

    try:
        wb.save(plant.path)
    except PermissionError:
        sys.exit("Could not save: close the workbook in Excel and try again.")
    print(f"\nSaved operation '{name}' to the '{sheet}' sheet (row {row}).")


def no_route_reason(plant, start, end, product, avoid, valve_penalty):
    """Why there is no route, found by relaxing the restrictions one at a time. Returns lines of text."""
    lines = [f"No valid route from {start} to {end}" + (f" for product '{product}'" if product else "") + "."]
    if product:
        G = plant.build_graph(None, avoid, valve_penalty, endpoints=(start, end))
        found = plant.find_routes(G, start, end, 1)
        if found:
            path = found[0]
            bad = [sid for sid in plant.path_segments(G, path)
                   if plant.segments[sid]["products"] is not None and product not in plant.segments[sid]["products"]]
            lines.append(f"  A route exists if product restrictions are ignored; it is blocked for '{product}' on: "
                         + ", ".join(f"{b} (allows {', '.join(sorted(plant.segments[b]['products']))})" for b in bad))
            return lines
    if avoid:
        G = plant.build_graph(product, (), valve_penalty, endpoints=(start, end))
        if plant.find_routes(G, start, end, 1):
            lines.append("  A route exists without the --avoid list; something you avoided is required.")
            return lines
    lines.append("  Check line directions (one-way segments, check valves), pass-through settings on tanks,")
    lines.append("  and items marked Out of service / Locked closed.")
    return lines


def explain_no_route(plant, start, end, product, avoid, valve_penalty):
    """Say why there is no route by relaxing the restrictions one at a time."""
    print("\n" + "\n".join(no_route_reason(plant, start, end, product, avoid, valve_penalty)))


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def ask(prompt, options=None, allow_blank=False):
    while True:
        ans = input(prompt).strip()
        if not ans and allow_blank:
            return None
        if options is None or ans in options:
            return ans
        match = [o for o in options if o.lower() == ans.lower()]
        if match:
            return match[0]
        print(f"  '{ans}' not found. Try again.")


def main():
    ap = argparse.ArgumentParser(description="Find valve routes in a piping network.")
    ap.add_argument("workbook")
    ap.add_argument("--from", dest="start")
    ap.add_argument("--to", dest="end")
    ap.add_argument("--product", help="product name; segments that don't allow it are skipped")
    ap.add_argument("--avoid", default="", help="comma-separated nodes, valves or segments to avoid")
    ap.add_argument("--alternatives", type=int, default=1, help="show the N best routes (default 1)")
    ap.add_argument("--valve-penalty", type=float, default=0.0,
                    help="extra cost per valve passed, to prefer routes with fewer valves")
    ap.add_argument("--save", metavar="OPERATION_NAME", help="write the best route to the Valve Matrix sheet")
    ap.add_argument("--check", action="store_true", help="only validate the database")
    args = ap.parse_args()

    plant = Plant(args.workbook)
    errors, warnings = plant.check()
    if args.check or errors:
        for e in errors:
            print("ERROR  : " + e)
        for w in warnings:
            print("WARNING: " + w)
        if errors:
            sys.exit(f"\n{len(errors)} error(s) in the database. Fix them before routing.")
        print(f"\nDatabase OK: {len(plant.nodes)} nodes, {len(plant.segments)} segments, {len(plant.valves)} valves.")
        return

    points = plant.points()
    start, end, product = args.start, args.end, args.product
    if not start or not end:
        print("Nodes : " + ", ".join(sorted(plant.nodes)))
        print("Valves: " + ", ".join(sorted(plant.valves)))
        start = start or ask("Start point: ", points)
        end = end or ask("End point  : ", points)
        if product is None and plant.products:
            print("Products: " + ", ".join(sorted(plant.products)))
            product = ask("Product (blank = ignore product restrictions): ", sorted(plant.products), allow_blank=True)

    for p in (start, end):
        if p not in plant.nodes and p not in plant.valves:
            sys.exit(f"'{p}' is not a known node or valve.")
    if product and plant.products and product not in plant.products:
        sys.exit(f"Product '{product}' is not in the Products sheet.")

    avoid = [a.strip() for a in args.avoid.split(",") if a.strip()]
    G = plant.build_graph(product, avoid, args.valve_penalty, endpoints=(start, end))

    paths = plant.find_routes(G, start, end, max(1, args.alternatives))
    if not paths:
        explain_no_route(plant, start, end, product, avoid, args.valve_penalty)
        sys.exit(1)

    header = f"{start} -> {end}" + ("  (recirculation loop)" if start == end else "") \
        + (f"  |  product: {product}" if product else "")
    print("\n" + header)
    results = []
    for i, path in enumerate(paths, start=1):
        label = "BEST ROUTE" if i == 1 else f"ALTERNATIVE {i - 1}"
        results.append(print_route(plant, G, path, label))

    if args.save:
        open_v, close_v, endpoint_v, segs = results[0]
        save_to_matrix(plant, args.save, product, start, end, segs, open_v, close_v, endpoint_v)


if __name__ == "__main__":
    main()
