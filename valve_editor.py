#!/usr/bin/env python3
"""
valve_editor.py - edit the valve database on an interactive diagram: add and connect
equipment, segments and valves, and see the checks and routes update as you edit.

The editor opens in your browser, but everything runs on this computer. The workbook
stays the file format: saving writes the same sheets (Nós, Trechos, Válvulas, Produtos)
and leaves the Excel formulas, drop-down lists and the Matriz de Válvulas untouched, so
valve_router.py, valve_visualizer.py and the spreadsheet checks keep working.

Usage
    python valve_editor.py valve_database.xlsx
    python valve_editor.py                          # asks which workbook to open
    python valve_editor.py valve_database.xlsx --port 8800 --no-browser

Needs valve_router.py, valve_visualizer.py and valve_editor.html in the same folder.
pip install openpyxl networkx
Works without internet if vis-network.min.js is saved next to this script (see the end of this file).
"""
import argparse
import copy
import json
import os
import re
import shutil
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

try:
    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator
    import valve_router as vr
    import valve_visualizer as vz
except ImportError as e:
    sys.exit(f"{e}\nPut valve_router.py and valve_visualizer.py in this folder and run: pip install openpyxl networkx")

# Columns the editor reads and writes, as (model field, column key). Column keys are looked up
# through valve_router.COLUMNS plus the extra columns below; the first field identifies a row.
EXTRA_COLUMNS = {
    "normal": ["Posição Normal", "Posicao Normal", "Normal Position"],
}
ALIASES = {**vr.COLUMNS, **EXTRA_COLUMNS}
TABLE_FIELDS = {
    "nodes": [("tag", "node_id"), ("type", "node_type"), ("desc", "desc"), ("area", "area"),
              ("pass", "pass"), ("status", "status"), ("notes", "notes")],
    "segments": [("tag", "seg_id"), ("from", "from"), ("to", "to"), ("direction", "direction"),
                 ("products", "products"), ("cost", "cost"), ("status", "status"), ("notes", "notes")],
    "valves": [("tag", "valve_tag"), ("segment", "segment"), ("order", "order"), ("body", "body"),
               ("actuation", "actuation"), ("status", "status"), ("normal", "normal"), ("notes", "notes")],
    "products": [("name", "product"), ("notes", "notes")],
}
NUMBER_FIELDS = {"cost", "order"}
LAYOUT_SHEETS = ["Layout do Diagrama", "Diagram Layout"]
CODE_SHEETS = ["Códigos", "Codigos", "Codes"]

# Drop-down values used when the workbook has no data validation for a column
DEFAULT_OPTIONS = {
    True: {  # Portuguese workbook
        "pass": ["Sim", "Não"], "node_status": ["Em operação", "Fora de operação"],
        "direction": ["Unidirecional", "Bidirecional"], "seg_status": ["Em operação", "Fora de operação"],
        "actuation": ["Manual", "Automática"], "normal": ["Aberta", "Fechada"],
        "valve_status": ["Em operação", "Fora de operação", "Travada fechada", "Travada aberta"],
        "node_type": ["Tanque", "Reator", "Misturador", "Bomba", "Trocador de Calor", "Filtro",
                      "Derivação", "Coletor", "Dreno", "Central CIP", "Equipamento"],
        "body": list(vz.VALVE_NAMES.values()),
    },
    False: {  # English workbook
        "pass": ["Yes", "No"], "node_status": ["In service", "Out of service"],
        "direction": ["One-way", "Bidirectional"], "seg_status": ["In service", "Out of service"],
        "actuation": ["Manual", "Automatic"], "normal": ["Open", "Closed"],
        "valve_status": ["In service", "Out of service", "Locked closed", "Locked open"],
        "node_type": ["Tank", "Reactor", "Mixer", "Pump", "Heat exchanger", "Filter",
                      "Junction", "Manifold", "Drain", "CIP", "Equipment"],
        "body": [],
    },
}
# (table, column key) -> option list name
OPTION_COLUMNS = {
    ("nodes", "node_type"): "node_type", ("nodes", "pass"): "pass", ("nodes", "status"): "node_status",
    ("segments", "direction"): "direction", ("segments", "status"): "seg_status",
    ("valves", "body"): "body", ("valves", "actuation"): "actuation",
    ("valves", "status"): "valve_status", ("valves", "normal"): "normal",
}
TAG_RE = re.compile(r"[A-Z0-9]+(-[A-Z0-9]+)*")


class EditorError(Exception):
    pass


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def clean(v):
    """Cell value -> JSON-friendly model value ('' for empty, whole floats as int)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float)):
        return v
    return str(v).strip()


def to_number(v):
    if v in ("", None):
        return None
    try:
        f = float(str(v).replace(",", "."))
    except ValueError:
        return v  # keep what was typed; the router reports it
    return int(f) if f.is_integer() else f


def column_map(ws, fields):
    """{model field: column number or None} from the header row."""
    headers = {str(c.value).strip().lower(): c.column for c in ws[1] if c.value is not None}
    cols = {}
    for field, key in fields:
        cols[field] = next((headers[a.lower()] for a in ALIASES[key] if a.lower() in headers), None)
    return cols


def list_options(wb, ws, col):
    """Values of the drop-down list (data validation) on a column, or None."""
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list" or not dv.formula1:
            continue
        if not any(r.min_col <= col <= r.max_col for r in dv.sqref.ranges):
            continue
        f = dv.formula1.strip()
        if f.startswith('"'):
            return [s.strip() for s in f.strip('"').split(",") if s.strip()]
        if "!" in f:
            sheet, rng = f.rsplit("!", 1)
            sheet = sheet.strip("'")
            if sheet in wb.sheetnames:
                cells = wb[sheet][rng.replace("$", "")]
                cells = cells if isinstance(cells, tuple) else ((cells,),)
                return [str(c.value).strip() for row in cells for c in (row if isinstance(row, tuple) else (row,))
                        if c.value not in (None, "")]
    return None


def read_codes(wb):
    """Tag codes from the Códigos sheet: node type -> code / default pass, valve body -> code / actuation."""
    codes = {"node": {}, "valve": {}}
    name = next((n for n in CODE_SHEETS if n in wb.sheetnames), None)
    if name is None:
        return codes
    ws = wb[name]
    headers = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]

    def col(*names):
        return next((i for i, h in enumerate(headers) if h in names), None)

    ntype, npass = col("tipo de nó", "tipo de no", "node type"), col("passagem padrão", "passagem padrao", "default pass-through")
    vbody, vact = col("tipo de corpo", "body type"), col("acionamento padrão", "acionamento padrao", "default actuation")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if ntype and row[ntype] and row[ntype - 1]:
            codes["node"][str(row[ntype]).strip()] = {
                "code": str(row[ntype - 1]).strip(),
                "pass": str(row[npass]).strip() if npass is not None and row[npass] else ""}
        if vbody and row[vbody] and row[vbody - 1]:
            codes["valve"][str(row[vbody]).strip()] = {
                "code": str(row[vbody - 1]).strip(),
                "actuation": str(row[vact]).strip() if vact is not None and row[vact] else ""}
    return codes


# --------------------------------------------------------------------------- #
#  Workbook <-> model
# --------------------------------------------------------------------------- #
def load_model(path):
    """Read the workbook into the editor model plus the info the page needs (drop-downs, codes)."""
    wb = load_workbook(path)  # formulas as written; data cells are plain values
    sheets = {key: vr.find_sheet(wb, key) for key in TABLE_FIELDS}
    portuguese = sheets["valves"] != "Valves"
    model = {key: [] for key in TABLE_FIELDS}
    options = dict(DEFAULT_OPTIONS[portuguese])

    for key, fields in TABLE_FIELDS.items():
        ws = wb[sheets[key]]
        cols = column_map(ws, fields)
        id_field = fields[0][0]
        if cols[id_field] is None:
            raise EditorError(f"Sheet '{ws.title}' has no '{ALIASES[fields[0][1]][0]}' column")
        managed = {c for c in cols.values() if c}
        for field, colkey in fields:
            name = OPTION_COLUMNS.get((key, colkey))
            if name and cols[field]:
                found = list_options(wb, ws, cols[field])
                if found:
                    options[name] = found
        for r in range(2, ws.max_row + 1):
            ident = ws.cell(r, cols[id_field]).value
            if ident in (None, "") or is_formula(ident):
                continue
            rec = {field: clean(ws.cell(r, c).value) if c else "" for field, c in cols.items()}
            if key == "segments":
                rec["products"] = [p.strip() for p in str(rec["products"]).split(",") if p.strip()]
            # other typed-in columns travel with their row, so reordering rows doesn't mix them up
            extra = {}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if c not in managed and v is not None and not is_formula(v):
                    extra[str(c)] = v if isinstance(v, (str, int, float, bool)) else str(v)
            if extra:
                rec["_extra"] = extra
            model[key].append(rec)

    positions = {}
    layout = next((n for n in LAYOUT_SHEETS if n in wb.sheetnames), None)
    if layout:
        for row in wb[layout].iter_rows(min_row=2, values_only=True):
            if row and row[0] and isinstance(row[1], (int, float)) and isinstance(row[2], (int, float)):
                positions[str(row[0])] = [row[1], row[2]]
    model["positions"] = positions

    codes = read_codes(wb)
    if codes["node"]:
        options["node_type"] = options.get("node_type") or list(codes["node"])
    info = {"file": Path(path).name, "sheets": sheets, "portuguese": portuguese,
            "options": options, "codes": codes}
    return model, info


def plant_tables(model, sheets):
    """Editor model -> the tables valve_router.Plant reads (as if read from the workbook)."""
    def rows(key, convert):
        id_field = TABLE_FIELDS[key][0][0]
        return (sheets[key], [dict(convert(item), _row=i)
                              for i, item in enumerate(model[key], start=2) if item.get(id_field)])

    return {
        "nodes": rows("nodes", lambda n: {
            "node_id": n["tag"], "node_type": n["type"], "desc": n["desc"], "pass": n["pass"], "status": n["status"],
            "area": n["area"]}),
        "products": rows("products", lambda p: {"product": p["name"]}),
        "segments": rows("segments", lambda s: {
            "seg_id": s["tag"], "from": s["from"], "to": s["to"], "direction": s["direction"],
            "products": ", ".join(s["products"]) or None, "cost": to_number(s["cost"]), "status": s["status"]}),
        "valves": rows("valves", lambda v: {
            "valve_tag": v["tag"], "segment": v["segment"], "order": to_number(v["order"]), "body": v["body"],
            "actuation": v["actuation"], "status": v["status"], "notes": v["notes"]}),
    }


def write_table(ws, records, fields, notes):
    """Write records from row 2 down, touching only data cells; formula columns stay as they are."""
    cols = column_map(ws, fields)
    id_col = cols[fields[0][0]]
    template_last = ws.max_row
    owned = [r for r in range(2, ws.max_row + 1) if ws.cell(r, id_col).value not in (None, "")]
    data_cols = {c for c in cols.values() if c}
    for rec in records:
        data_cols.update(int(c) for c in rec.get("_extra", {}))
    for r in owned:  # clear the rows that held data, then rewrite them in the new order
        for c in data_cols:
            cell = ws.cell(r, c)
            if not is_formula(cell.value):
                cell.value = None

    for i, rec in enumerate(records):
        r = i + 2
        if r > template_last:
            extend_row(ws, r)
        for field, c in cols.items():
            if not c:
                continue
            value = rec.get(field, "")
            if field == "products":
                value = ", ".join(value)
            elif field in NUMBER_FIELDS:
                value = to_number(value)
            ws.cell(r, c).value = None if value == "" else value
        for c, value in rec.get("_extra", {}).items():
            if not is_formula(ws.cell(r, int(c)).value):
                ws.cell(r, int(c)).value = value
    if len(records) + 1 > template_last and template_last > 2:
        notes.append(f"'{ws.title}' now has {len(records)} rows, more than the {template_last - 1} the template "
                     "was prepared for: the Excel check columns and drop-down lists don't cover the extra rows.")


def extend_row(ws, r):
    """Give a row past the end of the template the styles and formulas of row 2."""
    for c in range(1, ws.max_column + 1):
        src, dst = ws.cell(2, c), ws.cell(r, c)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if is_formula(src.value):
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)


def save_model(path, model):
    """Write the model into the workbook at path. Returns notes about anything worth telling the user."""
    wb = load_workbook(path)
    notes = []
    for key, fields in TABLE_FIELDS.items():
        write_table(wb[vr.find_sheet(wb, key)], model[key], fields, notes)

    # node and valve positions on the diagram, in a hidden sheet the scripts ignore
    name = next((n for n in LAYOUT_SHEETS if n in wb.sheetnames), None)
    if name is None:
        name = LAYOUT_SHEETS[0] if vr.find_sheet(wb, "valves") != "Valves" else LAYOUT_SHEETS[1]
        ws = wb.create_sheet(name)
        ws.sheet_state = "hidden"
    ws = wb[name]
    ws.delete_rows(1, ws.max_row)
    ws.append(["Tag", "X", "Y"])
    tags = {n["tag"] for n in model["nodes"]} | {v["tag"] for v in model["valves"]}
    for tag, (x, y) in sorted(model["positions"].items()):
        if tag in tags:
            ws.append([tag, round(x), round(y)])

    wb.calculation.fullCalcOnLoad = True  # Excel recalculates the check columns when the file is opened
    tmp = path.with_name(f".{path.stem}.saving{path.suffix}")
    try:
        wb.save(tmp)
        shutil.copymode(path, tmp)
        os.replace(tmp, path)
    except PermissionError:
        tmp.unlink(missing_ok=True)
        raise EditorError("Não foi possível salvar: feche a planilha no Excel e tente de novo.")
    return notes


# --------------------------------------------------------------------------- #
#  Checks, diagram and routes for the model being edited
# --------------------------------------------------------------------------- #
def editor_checks(model, codes):
    """The checks the spreadsheet's Verificação columns make, for data entered in the editor."""
    issues = []

    def add(level, tag, msg):
        issues.append({"level": level, "tag": tag, "msg": f"{tag}: {msg}"})

    for n in model["nodes"]:
        tag = n["tag"]
        if not TAG_RE.fullmatch(tag):
            add("warning", tag, "use só letras maiúsculas, números e hífen no tag")
        if not n["type"]:
            add("warning", tag, "falta o tipo de nó")
        code = codes["node"].get(n["type"], {}).get("code")
        if code and f"-{code}-" not in f"-{tag}-":
            add("warning", tag, f"tag não contém {code} (tipo {n['type']})")
    for s in model["segments"]:
        if s["from"] and s["from"] == s["to"]:
            add("error", s["tag"], "origem = destino")
    for v in model["valves"]:
        tag = v["tag"]
        if not TAG_RE.fullmatch(tag):
            add("warning", tag, "use só letras maiúsculas, números e hífen no tag")
        if not v["body"]:
            add("warning", tag, "falta o tipo de corpo")
        code = codes["valve"].get(v["body"], {}).get("code")
        if code and f"-{code}-" not in f"-{tag}-":
            add("warning", tag, f"tag não contém {code} (corpo {v['body']})")
    return issues


def tagged(level, msg, known):
    """Router message -> issue, pointing at the first tag it mentions so the page can select it."""
    tag = next((t for t in re.findall(r"[\w~-]+", msg) if t in known), None)
    return {"level": level, "tag": tag, "msg": msg}


def analyze(state, model):
    plant = vr.Plant(state.path, plant_tables(model, state.info["sheets"]))
    errors, warnings = plant.check()
    known = {n["tag"] for n in model["nodes"]} | {s["tag"] for s in model["segments"]} | {v["tag"] for v in model["valves"]}
    issues = ([tagged("error", e, known) for e in errors] + [tagged("warning", w, known) for w in warnings]
              + editor_checks(model, state.info["codes"]))
    issues.sort(key=lambda i: i["level"] != "error")
    diagram = vz.build_diagram(plant)
    return {"nodes": diagram["nodes"], "edges": diagram["edges"],
            "legend": vz.legend_html(diagram["kinds"]), "issues": issues}


def route_args(req):
    avoid = [a.strip() for a in req.get("avoid", []) if a.strip()]
    k = max(1, min(int(req.get("alternatives") or 1), 10))
    return req["start"], req["end"], (req.get("product") or None), avoid, float(req.get("penalty") or 0), k


def compute_routes(plant, start, end, product, avoid, penalty, k):
    if plant.errors:
        return {"error": "Corrija os erros da aba Verificação antes de calcular rotas.", "details": plant.errors}
    for p in (start, end):
        if p not in plant.nodes and p not in plant.valves:
            return {"error": f"'{p}' não é um nó nem uma válvula."}
    if product and plant.products and product not in plant.products:
        return {"error": f"O produto '{product}' não está na lista de produtos."}
    G = plant.build_graph(product, avoid, penalty, endpoints=(start, end))
    paths = plant.find_routes(G, start, end, k)
    if not paths:
        return {"none": vr.no_route_reason(plant, start, end, product, avoid, penalty)}
    manual = {t for t, v in plant.valves.items() if v["type"].lower() == "manual"}
    routes = []
    for i, path in enumerate(paths):
        overlay = vz.route_overlay(plant, G, path)
        diagram = vz.build_diagram(plant, overlay)
        routes.append({
            "label": "Melhor rota" if i == 0 else f"Alternativa {i}",
            "cost": plant.path_cost(G, path), "path": overlay["path"], "segments": overlay["segments"],
            "open": [[t, t in manual] for t in overlay["open"]],
            "close": [[t, t in manual, sid, at] for t, (sid, at) in sorted(overlay["close"].items())],
            "endpoint": overlay["endpoint"], "warnings": overlay["warnings"],
            "nodes": diagram["nodes"], "edges": diagram["edges"],
        })
    return {"routes": routes}


def save_operation(state, req):
    """Write a route to the Matriz de Válvulas, from the saved workbook (the page saves first)."""
    plant = vr.Plant(state.path)
    if plant.errors:
        return {"error": "Corrija os erros da aba Verificação antes de salvar operações.", "details": plant.errors}
    start, end, product, avoid, penalty, _ = route_args(req)
    index = int(req.get("index") or 0)
    G = plant.build_graph(product, avoid, penalty, endpoints=(start, end))
    paths = plant.find_routes(G, start, end, index + 1)
    if len(paths) <= index:
        return {"error": "Essa rota não existe no arquivo salvo. Calcule a rota de novo."}
    path = paths[index]
    open_v, close_v, endpoint_v, _ = plant.valve_states(path)
    try:
        vr.save_to_matrix(plant, req["name"], product, start, end, plant.path_segments(G, path),
                          open_v, close_v, endpoint_v)
    except SystemExit as exc:  # save_to_matrix exits when Excel has the file open
        return {"error": "Não foi possível salvar: feche a planilha no Excel e tente de novo." if exc.code else "Erro ao salvar."}
    state.mtime = state.path.stat().st_mtime_ns
    return {"ok": True}


# --------------------------------------------------------------------------- #
#  Web server
# --------------------------------------------------------------------------- #
class EditorState:
    def __init__(self, path):
        self.path = path
        self.lock = threading.Lock()
        self.mtime = None
        self.info = None
        self.backup = None

    def load(self):
        model, self.info = load_model(self.path)
        self.mtime = self.path.stat().st_mtime_ns
        return model


class Handler(BaseHTTPRequestHandler):
    state = None
    port = None

    def log_message(self, fmt, *args):  # keep the console quiet; errors are printed below
        pass

    def send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path, ctype):
        body = path.read_bytes()
        if path.name == "valve_editor.html":
            local = (HERE / "vis-network.min.js").exists()
            body = body.replace(b"__VIS_SRC__", (b"/vis-network.min.js" if local else vz.VIS_CDN.encode()))
            body = body.replace(b"__DIAGRAM_JS__", vz.DIAGRAM_JS.encode())
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def host_ok(self):
        # only pages served by this editor may call it (blocks other websites and DNS rebinding)
        return self.headers.get("Host") in (f"127.0.0.1:{self.port}", f"localhost:{self.port}")

    def do_GET(self):
        if not self.host_ok():
            return self.send_error(403)
        st = self.state
        try:
            if self.path == "/":
                return self.send_file(HERE / "valve_editor.html", "text/html; charset=utf-8")
            if self.path == "/vis-network.min.js" and (HERE / "vis-network.min.js").exists():
                return self.send_file(HERE / "vis-network.min.js", "text/javascript")
            if self.path == "/api/model":
                with st.lock:
                    model = st.load()
                return self.send_json({"model": model, "info": st.info})
        except (EditorError, OSError, SystemExit) as exc:
            return self.send_json({"error": str(exc)}, 500)
        self.send_error(404)

    def do_POST(self):
        if not self.host_ok() or self.headers.get("X-Valve-Editor") != "1":
            return self.send_error(403)
        st = self.state
        try:
            length = int(self.headers.get("Content-Length") or 0)
            req = json.loads(self.rfile.read(length) or b"{}")
            if self.path == "/api/analyze":
                return self.send_json(analyze(st, req["model"]))
            if self.path == "/api/layout":
                plant = vr.Plant(st.path, plant_tables(req["model"], st.info["sheets"]))
                return self.send_json({"positions": vz.area_layout(plant)})
            if self.path == "/api/route":
                plant = vr.Plant(st.path, plant_tables(req["model"], st.info["sheets"]))
                return self.send_json(compute_routes(plant, *route_args(req)))
            if self.path == "/api/save":
                with st.lock:
                    if not req.get("force") and st.path.stat().st_mtime_ns != st.mtime:
                        return self.send_json({"conflict": True})
                    backup = None
                    if st.backup is None:  # keep the file as it was before this session's first save
                        st.backup = st.path.with_name(f"{st.path.stem}.backup{st.path.suffix}")
                        shutil.copy2(st.path, st.backup)
                        backup = st.backup.name
                    notes = save_model(st.path, req["model"])
                    st.mtime = st.path.stat().st_mtime_ns
                return self.send_json({"ok": True, "backup": backup, "notes": notes})
            if self.path == "/api/save_operation":
                with st.lock:
                    return self.send_json(save_operation(st, req))
        except EditorError as exc:
            return self.send_json({"error": str(exc)}, 500)
        except Exception as exc:  # report to the page instead of dropping the connection
            import traceback
            traceback.print_exc()
            return self.send_json({"error": f"Erro interno: {exc}"}, 500)
        self.send_error(404)


def choose_workbook():
    """No file given: ask with a file dialog, or use valve_database.xlsx from this folder."""
    try:
        import tkinter
        from tkinter import filedialog
        root = tkinter.Tk()
        root.withdraw()
        name = filedialog.askopenfilename(title="Abrir banco de dados de válvulas",
                                          filetypes=[("Planilhas Excel", "*.xlsx")], initialdir=os.getcwd())
        root.destroy()
        if name:
            return name
    except Exception:
        pass
    for candidate in (Path.cwd() / "valve_database.xlsx", HERE / "valve_database.xlsx"):
        if candidate.exists():
            return str(candidate)
    sys.exit("Usage: python valve_editor.py valve_database.xlsx")


def main():
    ap = argparse.ArgumentParser(description="Edit the valve database on an interactive diagram.")
    ap.add_argument("workbook", nargs="?")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true", help="don't open the browser automatically")
    args = ap.parse_args()

    path = Path(args.workbook or choose_workbook()).resolve()
    if not path.exists():
        sys.exit(f"File not found: {path}")
    state = EditorState(path)
    try:
        model = state.load()
    except EditorError as exc:
        sys.exit(f"Could not read {path.name}: {exc}")
    print(f"{path.name}: {len(model['nodes'])} nodes, {len(model['segments'])} segments, "
          f"{len(model['valves'])} valves, {len(model['products'])} products")

    server = None
    for port in range(args.port, args.port + 20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    if server is None:
        sys.exit(f"No free port between {args.port} and {args.port + 19}.")
    Handler.state, Handler.port = state, port
    url = f"http://127.0.0.1:{port}/"
    print(f"Editor running at {url}  (Ctrl+C to stop)")
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nEditor stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()

# Offline use: download https://cdn.jsdelivr.net/npm/vis-network@9.1.9/standalone/umd/vis-network.min.js
# once and save it next to this script; the editor then loads it from there instead of the internet.
