# Valve Router

Tools for a piping network kept in an Excel workbook: find the best route between two
points, work out which valves to open and close for it, and see and edit the network on
an interactive diagram.

| File | What it does |
|---|---|
| `valve_router.py` | Command line. Finds the best route(s) between two points and lists the valves to **open** and **close**; can save the operation to the *Matriz de Válvulas* sheet. Also checks the workbook. |
| `valve_visualizer.py` | Draws the network as an interactive HTML page, optionally with a route highlighted. |
| `valve_editor.py` + `valve_editor.html` | Interactive editor in the browser: add and connect equipment, segments and valves, check the data and calculate routes as you edit, then save back to the workbook. |
| `valvulas_m.xlsx` | Example workbook. |

Everything runs on your own computer; the browser pages only load the drawing library
(vis-network) from the internet, and that can be avoided too (see [Offline use](#offline-use)).

## Installation

Python 3.9 or newer.

```bash
pip install -r requirements.txt
```

## Quick start

```bash
python valve_router.py valvulas_m.xlsx --check
```

```bash
python valve_router.py valvulas_m.xlsx --from TQ-1 --to RF-02
```

```bash
python valve_visualizer.py valvulas_m.xlsx --open
```

```bash
python valve_editor.py valvulas_m.xlsx
```

## The workbook

The workbook can be in Portuguese (sheets below) or in English (`Nodes`, `Segments`,
`Valves`, `Products`, `Valve Matrix`). Fill in the yellow cells; the grey *Verificação*
columns are Excel formulas that check the data. The *LEIAME* sheet has the full filling
instructions. Choose values such as *Unidirecional* or *Em operação* from the drop-down
lists so the scripts recognise them.

**Nós** – one row per tank, reactor, mixer, pump, heat exchanger, filter, junction,
manifold or drain.

| Column | Meaning |
|---|---|
| Tag do Nó | Unique tag, e.g. `TQ-101` |
| Tipo de Nó | Tanque, Reator, Misturador, Bomba, Trocador de Calor, Filtro, Derivação, Coletor, Dreno, Central CIP, Equipamento |
| Descrição | Free text |
| Área | Plant area; the diagram groups items by area |
| Permite Passagem | *Sim*: a route may pass through it. *Não*: only start or end. Blank: by type (tanks, reactors, mixers, drains, CIP = Não) |
| Status | *Fora de operação* blocks routes |

Every point where lines split or join must be a junction node (*Derivação*, `DV`), so that
each segment connects exactly two nodes.

**Trechos** – one row per piece of pipe between two nodes.

| Column | Meaning |
|---|---|
| Tag do Trecho | Unique tag, e.g. `TR-001` |
| Origem / Destino | Node tags; Origem → Destino is the drawn direction |
| Sentido | *Unidirecional* (flow only Origem → Destino) or *Bidirecional* |
| Produtos Permitidos | Blank = any product; otherwise names from *Produtos*, separated by commas |
| Custo | Relative cost (length, preference); blank = 1. The router picks the cheapest valid route |
| Status | *Fora de operação* blocks routes |

**Válvulas** – one row per valve.

| Column | Meaning |
|---|---|
| Tag da Válvula | Unique tag, e.g. `VES-101`. A valve number is never reused, whatever the body type |
| Trecho | Segment the valve is on |
| Posição no Trecho | 1 = closest to the segment's Origem, 2 = next, … |
| Tipo de Corpo | Esfera, Borboleta, Gaveta, Globo, Diafragma, Macho, Agulha, Retenção, Três vias |
| Acionamento | Manual or Automática. Check valves (*Retenção*) only let flow pass Origem → Destino |
| Status | *Fora de operação* / *Travada fechada* block routes; *Travada aberta* lets flow pass but is not used to isolate |
| Posição Normal | Aberta / Fechada |

**Produtos** – the products and fluids of the plant (avoid commas in names).

**Códigos** – tag codes per node type and valve body; feed the drop-down lists and the checks.

**Matriz de Válvulas** – filled by the scripts when you save an operation. Do not edit by hand.
`O` / `C` = open / close, `MO` / `MC` = manual valve to open / close, `E` = endpoint.

**Layout do Diagrama** (hidden) – diagram positions saved by the editor.

Tag format: `[Area]-[Type]-[Number]`, e.g. `100-TQ-101`, or `TQ-101` when the number
already starts with the area digit. No accents, no spaces.

## valve_router.py

```bash
python valve_router.py valvulas_m.xlsx
```

Without `--from` / `--to` it asks for the start point, end point and product.

| Option | |
|---|---|
| `--from TAG --to TAG` | Start and end (nodes or valves). The same tag twice finds a recirculation loop |
| `--product NAME` | Only use segments that allow this product |
| `--avoid TAG,TAG` | Nodes, valves or segments the route must not use |
| `--alternatives N` | Show the N best routes |
| `--valve-penalty X` | Extra cost per valve passed, to prefer routes with fewer valves |
| `--save "NAME"` | Write the best route to *Matriz de Válvulas* as operation NAME |
| `--check` | Only validate the workbook |

Example:

```bash
python valve_router.py valvulas_m.xlsx --from TQ-1 --to RF-02 --product Água --alternatives 3 --save "Água TQ-1 para RF-02"
```

For each route it lists the valves to open, the valves to close to isolate every branch
leaving the route, valves at the endpoints, and warnings where a branch has no valve that
can isolate it. When there is no route it explains why (product restriction, an avoided
item, line directions or items out of service).

## valve_visualizer.py

```bash
python valve_visualizer.py valvulas_m.xlsx -o diagrama.html --open
```

| Option | |
|---|---|
| `-o FILE` | Output page (default `valve_diagram.html`) |
| `--from TAG --to TAG` | Highlight the best route: valves to open in green, to close in red |
| `--product`, `--avoid`, `--valve-penalty` | As in `valve_router.py` |
| `--layout area` | Default. Automatic layout, see below |
| `--layout saved` | Use the positions saved by the editor (automatic layout for anything not placed) |
| `--layout free` | Force-directed layout |
| `--title TEXT` | Page title |
| `--offline-lib vis-network.min.js` | Embed the drawing library so the page works without internet |
| `--open` | Open the page in the browser |

**Automatic layout.** Each area (column *Área*) gets its own labelled box, and flow runs
from left to right. The order of the items is chosen to keep line crossings low, and
items in line with each other are placed exactly level. Lines are drawn with right angles
and routed around the items. Where two lines cross without being connected, one of them
jumps over the other with a small arc. Several arrangements are tried (a few seconds for a
large plant) and the one with the fewest crossings for its size is kept.

## valve_editor.py

```bash
python valve_editor.py valvulas_m.xlsx
```

Opens the editor in the browser (at `http://127.0.0.1:8765/`; `--port` to change it,
`--no-browser` to not open it automatically). Without a file name it asks for one, or uses
`valve_database.xlsx` from the current folder.

- **＋ Equipamento** (`N`): click an empty spot to place a node.
- **＋ Trecho** (`T`): drag from one node to another, in the direction of flow.
- **＋ Válvula** (`V`): click on a segment's line where the valve goes.
- Click an item to edit it in the side panel; drag to move it. `Delete` deletes,
  `Ctrl+Z` / `Ctrl+Y` undo / redo, `Ctrl+S` saves.
- **Reorganizar** applies the automatic layout described above.
- **Verificação** lists the problems in the data; **Rota** calculates routes and saves
  operations to *Matriz de Válvulas*.

Saving writes the data sheets and the diagram positions back to the workbook and leaves
the formulas, drop-down lists and *Matriz de Válvulas* as they are. Before the first save
in a session the editor keeps a copy of the file as `<name>.backup.xlsx`. Close the
workbook in Excel before saving. If the file was changed outside the editor in the
meantime, the editor asks before overwriting it.

## Offline use

The diagram pages load vis-network from the internet. On a computer without internet,
download it once from
https://cdn.jsdelivr.net/npm/vis-network@9.1.9/standalone/umd/vis-network.min.js and save
it next to the scripts. The editor then uses the local copy automatically; for the
visualizer pass `--offline-lib vis-network.min.js`.
