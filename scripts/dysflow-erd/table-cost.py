"""Code cost of changing each table and column (phase 0 of the coexistence epic).

Input: the directory written by collect-usage.ps1 (bases.tsv and src/<repo>/src), its
columns.tsv (inventory-data.ps1 -Mode Columns) and the classification.tsv written by
classify-tables.py. Output, in the output directory:

- table-cost.tsv: per backend table with consumers, the code lines and files that name it in
  the applications that use it.
- column-cost.tsv: per column of those tables, the code lines that name it as a field
  (!Col, .Col, [Col]) or inside SQL text, counted only in files that name its table.
- column-risks.tsv: columns whose type or rules do not move as they are to PostgreSQL.

Usage: python table-cost.py <input dir> <classification.tsv> [<output dir>]
"""

import collections
import csv
import os
import re
import sys
from pathlib import Path

D, CLASSIFICATION = sys.argv[1], sys.argv[2]
OUT = sys.argv[3] if len(sys.argv) > 3 else D

# DAO field types that need a decision before the move (Field.Type values).
TYPE_RISK = {11: "objeto OLE", 101: "datos adjuntos"}
ACCESS_FN = re.compile(r"\b(?:date|now|time|nz|iif|format|dlookup|dmax|left|right|mid)\s*\(", re.I)


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


def is_test(fn):
    low = fn.lower()
    return bool(re.match(r"^(test|modbattery|battery)", low)) or "_test" in low or "test_" in low


DATA_SOURCE = re.compile(r"^\s*(RecordSource|RowSource|ControlSource)\s*=", re.IGNORECASE)


def source_lines(src):
    """(relative file, line) for the code and data-source properties of an application."""
    for root, _, files in os.walk(src):
        for fn in sorted(files):
            low = fn.lower()
            if is_test(fn) or not low.endswith((".bas", ".cls", ".form.txt", ".report.txt")):
                continue
            p = os.path.join(root, fn)
            rel = os.path.relpath(p, src).replace(os.sep, "/")
            text = Path(p).read_text(encoding="utf-8", errors="replace").splitlines()
            in_code = not low.endswith((".form.txt", ".report.txt"))
            in_source = False  # inside a data-source property, which Access splits over lines
            for ln in text:
                if ln.strip().startswith(("CodeBehindForm", "CodeBehindReport")):
                    in_code = True
                    continue
                if in_code:
                    ln = strip_vba_comment(ln)
                elif DATA_SOURCE.match(ln):
                    in_source = True
                elif in_source and ln.lstrip().startswith('"'):
                    pass  # continuation line of the same property value
                else:
                    in_source = False
                    continue
                if ln.strip():
                    yield rel, ln


def tsv(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f, delimiter="\t"))


classes = {}
for r in csv.reader(open(CLASSIFICATION, encoding="utf-8", newline=""), delimiter="\t"):
    if len(r) < 6 or r[0] == "backend":
        continue
    consumers = [c.strip() for c in r[5].split(",") if c.strip()]
    if consumers:
        classes[(r[0], r[1])] = (r[4], consumers)

# The names each application uses for each table (its own name or a link name), as resolved by
# classify-tables.py: counting the physical name would miss link aliases and count the same name
# twice when two backends have a table with that name.
names_of = collections.defaultdict(lambda: collections.defaultdict(set))  # table -> app -> names
for r in tsv(os.path.join(os.path.dirname(os.path.abspath(CLASSIFICATION)), "resolved.tsv")):
    names_of[(r["backend"], r["tabla"].lower())][r["aplicacion"]].add(r["nombre"].lower())

columns = collections.defaultdict(list)
for r in tsv(os.path.join(D, "columns.tsv")):
    columns[(r["backend"].replace(".accdb", ""), r["table"])].append(r)

bases = [
    line.split("\t")
    for line in Path(D, "bases.tsv").read_text(encoding="utf-8").splitlines()
    if line.strip()
]
code = {app: list(source_lines(os.path.join(D, "src", repo, "src"))) for app, repo, _ref in bases}
TOKEN = re.compile(r"[a-z0-9_à-ÿ]+")
FIELD = re.compile(r"[!.]\[?([a-z0-9_à-ÿ ]+?)\]?(?![a-z0-9_à-ÿ])|\[([^\]]+)\]")
# Per (app, file): in how many lines each identifier appears at all, and as a field or inside
# SQL text. One pass over the code; the costs below are sums of these counters.
any_count = collections.defaultdict(collections.Counter)
col_count = collections.defaultdict(collections.Counter)
for app, lines in code.items():
    for f, ln in lines:
        low = ln.lower()
        key = (app, f)
        any_count[key].update(set(TOKEN.findall(low)))
        named = {a or b for a, b in FIELD.findall(low)}
        for s in low.split('"')[1::2]:
            named.update(TOKEN.findall(s))
            named.update(m.strip() for m in re.findall(r"\[([^\]]+)\]", s))
        col_count[key].update({n.strip() for n in named if n.strip()})
files_of = collections.defaultdict(list)  # app -> its file keys
for key in any_count:
    files_of[key[0]].append(key)

table_rows, column_rows = [], []
for (backend, table), (cls, consumers) in sorted(classes.items()):
    if (backend, table) not in columns:
        continue  # front tables and tables outside the reference backends
    lines, files = 0, []
    for app, names in names_of[(backend, table.lower())].items():
        for key in files_of.get(app, []):
            # names with spaces only appear between brackets
            n = sum((any_count if " " not in k else col_count)[key][k] for k in names)
            if n:
                lines += n
                files.append(key)
    table_rows.append((backend, table, cls, ", ".join(consumers), lines, len(files)))
    for col in columns[(backend, table)]:
        name = col["column"].lower()
        n = sum(col_count[key][name] for key in files)
        column_rows.append((backend, table, col["column"], col["type"], n))

risks = []
for (backend, table), cols in sorted(columns.items()):
    if (backend, table) not in classes:
        continue
    for c in cols:
        kind = TYPE_RISK.get(int(c["type"]))
        if c["complex"].lower() == "true" and not kind:
            kind = "multivalor"
        if kind:
            risks.append((backend, table, c["column"], kind, ""))
        for what in ("default", "validation"):
            v = c[what]
            if v and ACCESS_FN.search(v):
                label = "valor por defecto" if what == "default" else "regla de validación"
                risks.append((backend, table, c["column"], label + " con función de Access", v))

os.makedirs(OUT, exist_ok=True)
for name, header, rows in (
    ("table-cost.tsv", ["backend", "table", "class", "consumers", "lines", "files"], table_rows),
    ("column-cost.tsv", ["backend", "table", "column", "type", "lines"], column_rows),
    ("column-risks.tsv", ["backend", "table", "column", "risk", "value"], risks),
):
    with open(os.path.join(OUT, name), "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter="\t", lineterminator="\n")
        w.writerow(header)
        w.writerows(rows)
print(len(table_rows), "tables;", len(column_rows), "columns;", len(risks), "column risks")
