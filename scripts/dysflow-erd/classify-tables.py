"""Classify every table of the phase 0 reference copy by the applications that use it.

Inputs, all in one directory: ownership.tsv (inventory-data.ps1 -Mode Ownership), inventory.tsv
(inventory-data.ps1), and fronts.tsv, queries.tsv, bases.tsv and src/<repo>/src (collect-usage.ps1).
Output: classification.tsv (one row per table) and doubts.tsv (names that can still mean two tables).
Usage: python classify-tables.py <dir> [<output dir>]
"""

import collections
import csv
import os
import re
import sys
from pathlib import Path

D = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else D
os.makedirs(OUT, exist_ok=True)
S = lambda b: b.replace(".accdb", "")


def tsv(name):
    return list(
        csv.DictReader(open(os.path.join(D, name), encoding="utf-8"), delimiter="\t")
    )


APP_BACKENDS = {
    "Lanzadera": ["Lanzadera_Datos"],
    "Expedientes": ["Expedientes_datos"],
    "Condor": ["condor_datos"],
    "Seguridad": ["Seguridad_datos"],
    "HPS": ["HPST"],
    "Solicitudes HPS": ["Solicitudes_HPS_datos"],
    "Gestión de Riesgos": ["Gestion_Riesgos_Datos"],
    "No Conformidades": ["NoConformidades_Datos"],
    "Registros": ["Registro_Ent_Salida_Datos", "Registro_Datos"],
    "Brass": ["Gestion_Brass_Gestion_Datos"],
    "AGEDYS Técnico": ["AGEDYS_DATOS"],
    "AGEDYS ECO": ["AGEDYS_DATOS"],
    "Contratos": ["GestionContratos_datos"],
}
OWNERS = collections.defaultdict(list)
for app, bks in APP_BACKENDS.items():
    for b in bks:
        OWNERS[b].append(app)
OUT_OF_SCOPE = {
    "AGEDO20_Datos",
    "Tareas_datos1",
    "SICA_datos",
    "Control_Cambios_datos",
    "Correos_datos",
    "GestionContratos_datos",
}

own = tsv("ownership.tsv")
inv = {
    (S(r["backend"]), r["table"]): (r["rows"], r["pk"])
    for r in tsv("inventory.tsv")
    if r["kind"] == "table"
}
fronts = tsv("fronts.tsv")
queries = tsv("queries.tsv")
bases = [
    l.split("\t")
    for l in Path(D, "bases.tsv").read_text(encoding="utf-8").splitlines()
    if l.strip()
]

local = collections.defaultdict(set)  # name -> backends where it is a local table
blink = {}  # (backend, name) -> (backend, table)
for r in own:
    k = r["name"].lower()
    if r["kind"] == "local":
        local[k].add(S(r["backend"]))
    else:
        blink[(S(r["backend"]), k)] = (S(r["target_backend"]), r["source_table"])
flocal = collections.defaultdict(set)  # app -> front local names
flink = {}  # (app, name) -> (backend, table)
front_tables = []
for r in fronts:
    if r["kind"] == "local":
        flocal[r["app"]].add(r["name"].lower())
        front_tables.append((r["app"], r["name"]))
    elif r["kind"] == "link":
        flink[(r["app"], r["name"].lower())] = (
            S(r["target_backend"]),
            r["source_table"],
        )

names = (
    set(local)
    | {k for (_, k) in blink}
    | {k for v in flocal.values() for k in v}
    | {k for (_, k) in flink}
)
canon = {}
for r in own:
    canon.setdefault(r["name"].lower(), r["name"])
for r in fronts:
    canon.setdefault(r["name"].lower(), r["name"])
SPACED = sorted((n for n in names if " " in n), key=len, reverse=True)
TOKEN = re.compile(r"[A-Za-z0-9_]+")


def names_in(line):
    low = line.lower()
    hits = {t for t in TOKEN.findall(low) if t in names}
    for n in SPACED:
        if n in low:
            hits.add(n)
    return hits


def strip_vba_comment(line):
    s = line.lstrip()
    if s.startswith("'") or s.lower().startswith("rem "):
        return ""
    out, inq = [], False
    for ch in line:
        if ch == '"':
            inq = not inq
        elif ch == "'" and not inq:
            break
        out.append(ch)
    return "".join(out)


SKIP_FORM = re.compile(
    r"^\s*(Name|Caption|ControlTipText|StatusBarText|Tag|Description)\s*=",
    re.IGNORECASE,
)


def used_names(src_dir):
    found = collections.defaultdict(set)  # name -> {'código', 'formulario'}
    for root, _, files in os.walk(src_dir):
        for fn in files:
            low = fn.lower()
            if (
                re.match(r"^(test|modbattery|battery)", low)
                or "_test" in low
                or "test_" in low
            ):
                continue
            p = os.path.join(root, fn)
            try:
                text = Path(p).read_text(encoding="utf-8", errors="replace")
            except OSError:
                continue
            if low.endswith((".form.txt", ".report.txt")):
                kind = "formulario"
                lines = [l for l in text.splitlines() if not SKIP_FORM.match(l)]
            elif low.endswith((".bas", ".cls", ".sql")):
                kind = "código"
                lines = [strip_vba_comment(l) for l in text.splitlines()]
            else:
                continue
            for l in lines:
                for k in names_in(l):
                    found[k].add(kind)
    return found


# Ambiguous references resolved by reading the code that opens them (which connection is used).
# An empty list means the match is not a table use (text, SQL alias, or a table missing where it is opened).
OVERRIDES = {
    ("Lanzadera", "tbherramientadocayuda"): (
        [],
        "getdb() sobre Lanzadera_Datos, donde no existe: ayuda rota (Lanzadera#24)",
    ),
    ("Registros", "tbherramientadocayuda"): (
        [],
        "Dame() usa CurrentDb y el front no tiene la tabla: ayuda rota",
    ),
    ("Seguridad", "tbcorreosenviados"): (
        [("Correos_datos", "TbCorreosEnviados")],
        "abre con getdbCorreo",
    ),
    ("Gestión de Riesgos", "entradas"): (
        [],
        "palabra de un texto del formulario, no una tabla",
    ),
    ("Gestión de Riesgos", "salidas"): (
        [],
        "palabra de un texto del formulario, no una tabla",
    ),
    ("Expedientes", "tbexpedientes_1"): (
        [],
        "alias en un SQL (TbExpedientes AS TbExpedientes_1)",
    ),
    ("AGEDYS Técnico", "tbexpedientesanualidades"): (
        [("Expedientes_datos", "TbExpedientesAnualidades")],
        "CurrentDb: vínculo del front a Expedientes_datos",
    ),
    ("AGEDYS Técnico", "tbherramientadocayuda"): (
        [("front:AGEDYS Técnico", "TbHerramientaDocAyuda")],
        "Dame() usa CurrentDb: la copia del front",
    ),
    ("AGEDYS Técnico", "tbstgparametrosparaformulario"): (
        [("front:AGEDYS Técnico", "TbSTGParametrosParaFormulario")],
        "CurrentDb: la copia del front",
    ),
    ("Registros", "tbusuarios"): (
        [],
        "RowSource con CurrentDb y el front no vincula la tabla: combo roto (REG-G5)",
    ),
    ("AGEDYS ECO", "tbauxexpedientesvivos"): (
        [("front:AGEDYS ECO", "TbAuxExpedientesVivos")],
        "CurrentDb: la copia del front",
    ),
}


def resolve(app, k):
    """All physical tables a name can mean for an application, and whether it is ambiguous."""
    if (app, k) in OVERRIDES:
        return (
            set(OVERRIDES[(app, k)][0]),
            "resuelto en el código: " + OVERRIDES[(app, k)][1],
            False,
        )
    cands = set()
    if k in flocal[app]:
        cands.add(("front:" + app, canon.get(k, k)))
    if (app, k) in flink:
        cands.add(flink[(app, k)])
    for b in APP_BACKENDS[app]:
        if b in local[k]:
            cands.add((b, canon.get(k, k)))
        if (b, k) in blink:
            cands.add(blink[(b, k)])
    if cands:
        return cands, "propio contexto", len(cands) > 1
    others = sorted(local[k])
    if len(others) == 1:
        return {(others[0], canon.get(k, k))}, "apertura directa", False
    if len(others) > 1:
        return {(b, canon.get(k, k)) for b in others}, "apertura directa", True
    return set(), "", False


uses = collections.defaultdict(
    lambda: collections.defaultdict(set)
)  # (backend|front, table lower) -> app -> evidence
doubts = []
for app, repo, ref in bases:
    src = os.path.join(D, "src", repo, "src")
    for k, kinds in used_names(src).items():
        cands, how, amb = resolve(app, k)
        for c in cands:
            uses[(c[0], c[1].lower())][app] |= {f"{x} ({how})" for x in kinds}
        if amb:
            doubts.append(
                (
                    app,
                    canon.get(k, k),
                    "el nombre puede referirse a: "
                    + ", ".join(sorted(f"{a}.{t}" for a, t in cands)),
                )
            )

QREF = re.compile(
    r"(?:FROM|JOIN|INTO|UPDATE)\s+(\[[^\]]+\]|[A-Za-z0-9_]+)", re.IGNORECASE
)
for q in queries:
    kind, who = q["owner"].split(":", 1)
    tabs = {t.strip("[]").lower() for t in QREF.findall(q["sql"])}
    if kind == "front":
        apps = [who]
    elif kind == "data":
        apps = ["Seguridad"]
    else:
        apps = OWNERS.get(who, [])
    for app in apps:
        for k in tabs:
            if kind == "backend":
                b = who
                cands = (
                    {(b, canon.get(k, k))}
                    if b in local[k]
                    else ({blink[(b, k)]} if (b, k) in blink else set())
                )
            else:
                cands, _, _ = resolve(app, k)
            label = (
                "catálogo de datos (Tb0TipoRegistros)"
                if kind == "data"
                else f"consulta guardada {q['query']} ({kind})"
            )
            for c in cands:
                uses[(c[0], c[1].lower())][app].add(label)

# AGEDO and Tareas stay alive in Access (decision 14) but their fronts are not analysed: the links of
# their backends are the only evidence of what they consume.
for r in own:
    if r["kind"] == "link" and S(r["backend"]) in ("AGEDO20_Datos", "Tareas_datos1"):
        consumer = "AGEDO" if S(r["backend"]) == "AGEDO20_Datos" else "Tareas"
        uses[(S(r["target_backend"]), r["source_table"].lower())][consumer].add(
            "vínculo de su backend (consumidor permanente, decisión 14)"
        )

rows = []
for r in own:
    if r["kind"] != "local":
        continue
    b, t = S(r["backend"]), r["name"]
    u = uses.get((b, t.lower()), {})
    users = set(u)
    ownapps = set(OWNERS.get(b, []))
    n, pk = inv.get((b, t), ("", ""))
    if b in OUT_OF_SCOPE and b != "GestionContratos_datos":
        cls = "Consumida por el alcance" if users else "Sin consumidores del alcance"
    elif users & ownapps:
        cls = "Propia y compartida" if users - ownapps else "Propia"
    elif len(users) == 1:
        cls = "Mal ubicada"
    elif users:
        cls = "Mal ubicada con dueño por decidir"
    else:
        cls = "Huérfana"
    ev = "; ".join(f"{a}: " + ", ".join(sorted(v)) for a, v in sorted(u.items()))
    rows.append(
        {
            "ubicacion": b,
            "tabla": t,
            "filas": n,
            "clave": pk,
            "clase": cls,
            "usada_por": ", ".join(sorted(users)),
            "evidencia": ev,
        }
    )
for app, t in front_tables:
    u = uses.get(("front:" + app, t.lower()), {})
    cls = "Usada por su front" if u else "Sin uso en el código del front"
    ev = "; ".join(f"{a}: " + ", ".join(sorted(v)) for a, v in sorted(u.items()))
    rows.append(
        {
            "ubicacion": "front " + app,
            "tabla": t,
            "filas": "",
            "clave": "",
            "clase": cls,
            "usada_por": ", ".join(sorted(u)),
            "evidencia": ev,
        }
    )

with open(
    os.path.join(OUT, "classification.tsv"), "w", encoding="utf-8", newline=""
) as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0]), delimiter="\t")
    w.writeheader()
    w.writerows(rows)
seen = set()
with open(os.path.join(OUT, "doubts.tsv"), "w", encoding="utf-8", newline="") as f:
    f.write("aplicacion\tnombre\tmotivo\n")
    for d in doubts:
        if d not in seen:
            seen.add(d)
            f.write("\t".join(d) + "\n")
c = collections.Counter(r["clase"] for r in rows)
print(len(rows), "tables;", dict(c), ";", len(seen), "doubts")
