"""Inventory of VBA risk patterns per application front (phase 0 of the coexistence epic).

Input: the directory written by collect-usage.ps1 (bases.tsv and src/<repo>/src).
Output: patterns.tsv (one row per hit) and summary.tsv (hits and files per application and
pattern). The patterns are the ones the epic lists for every front: code that would not survive
the move of the data to PostgreSQL behind ODBC.

Usage: python risk-patterns.py <input dir> [<output dir>]
"""

import collections
import csv
import os
import re
import sys
from pathlib import Path

D = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else D

# Pattern id -> (description, test on one logical line of code without comments).
# String literals are where SQL lives, so most tests look inside them.


def strings_of(line):
    """Text inside the double-quoted literals of a VBA line ("" is an escaped quote)."""
    out, cur, inq, i = [], [], False, 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if inq and i + 1 < len(line) and line[i + 1] == '"':
                cur.append('"')
                i += 2
                continue
            if inq:
                out.append("".join(cur))
                cur = []
            inq = not inq
        elif inq:
            cur.append(ch)
        i += 1
    return out


LIKE_STAR = re.compile(r"\blike\s*'?\*|\blike\s*'[^']*\*", re.IGNORECASE)
DATE_FN = re.compile(r"\b(?:date|now|time)\s*\(\s*\)", re.IGNORECASE)
SEEK = re.compile(r"\.seek\b|\bdbopentable\b", re.IGNORECASE)
OPEN_RS = re.compile(r"\bopenrecordset\s*\(", re.IGNORECASE)
SEE_CHANGES = re.compile(r"\bdbseechanges\b", re.IGNORECASE)
SQL_WORD = re.compile(r"\b(?:select|update|insert|delete|where|from)\b", re.IGNORECASE)


# Max + 1 identifier generators are detected per procedure: one that reads a maximum (SQL Max(
# or DMax) and adds 1 is a generator. The hit points at the procedure, once.
PROC_START = re.compile(
    r"^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?"
    r"(?:function|sub|property\s+(?:get|let|set))\s+([a-z0-9_]+)",
    re.IGNORECASE,
)
PROC_END = re.compile(r"^\s*end\s+(?:function|sub|property)\b", re.IGNORECASE)
DMAX = re.compile(r"\bdmax\s*\(", re.IGNORECASE)
PLUS_ONE = re.compile(r"\+\s*1\b(?!\.)")
# Recordsets that are edited (.AddNew or .Edit, directly or inside a With block) and were
# opened without dbSeeChanges: the subset of sin_dbseechanges that fails on identity tables.
SET_RS = re.compile(r"^\s*set\s+([a-z0-9_]+)\s*=.*\bopenrecordset\s*\(", re.IGNORECASE)
WITH = re.compile(r"^\s*with\s+([a-z0-9_]+)\s*$", re.IGNORECASE)
EDITS = re.compile(r"(?:\b([a-z0-9_]+)|(?<![a-z0-9_]))\.(?:addnew|edit)\b", re.IGNORECASE)
# A loop counter (x = x + 1) is not the maximum plus one.
COUNTER = re.compile(r"^\s*([a-z0-9_.]+)\s*=\s*\1\s*\+\s*1\s*$", re.IGNORECASE)


def like_star(line, strs):
    # Access wildcards (* and ?) in LIKE: PostgreSQL uses % and _.
    joined = " ".join(strs)
    return bool(LIKE_STAR.search(joined)) or (
        bool(re.search(r"\blike\b", joined, re.IGNORECASE)) and "*" in joined
    )


NUMBER_MASK = re.compile(r"[#0][#0,.]*[,0][#0,.]*|#,#")


def date_hash(line, strs):
    # Date literals delimited with # inside SQL text; number masks ("#,##0.00") are not dates.
    return any("#" in NUMBER_MASK.sub("", s) for s in strs) and (
        any(SQL_WORD.search(s) for s in strs) or bool(re.search(r"\bformat\s*\(", line, re.I))
    )


def date_in_sql(line, strs):
    # Date(), Now() or Time() evaluated by the engine inside the SQL text.
    return any(DATE_FN.search(s) for s in strs)


def seek_table(line, strs):
    # Seek and table-type recordsets do not exist on ODBC tables.
    return bool(SEEK.search(line))


def no_see_changes(line, strs):
    # Recordsets opened without dbSeeChanges fail on SQL Server/PostgreSQL identity tables.
    return bool(OPEN_RS.search(line)) and not SEE_CHANGES.search(line)


PATTERNS = [
    ("like_comodin", "LIKE con comodines de Access (* ?)", like_star),
    ("fecha_almohadilla", "Literal de fecha con # en SQL", date_hash),
    ("fecha_en_sql", "Date(), Now() o Time() dentro del SQL", date_in_sql),
    ("seek_tabla", "Seek o dbOpenTable", seek_table),
    ("sin_dbseechanges", "OpenRecordset sin dbSeeChanges", no_see_changes),
]


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


def logical_lines(lines):
    """Join VBA continuation lines (ending in ' _') and keep the first physical line number."""
    buf, start = [], None
    for n, ln in enumerate(lines, 1):
        code = strip_vba_comment(ln).rstrip()
        if start is None:
            start = n
        if code.endswith(" _"):
            buf.append(code[:-2])
            continue
        buf.append(code)
        yield start, " ".join(buf)
        buf, start = [], None
    if buf:
        yield start, " ".join(buf)


def code_lines(path):
    text = Path(path).read_text(encoding="utf-8", errors="replace").splitlines()
    low = path.lower()
    if low.endswith((".form.txt", ".report.txt")):
        for i, ln in enumerate(text):
            if ln.strip().startswith(("CodeBehindForm", "CodeBehindReport")):
                return text[i + 1 :], i + 1
        return [], 0
    return text, 0


def is_test(fn):
    low = fn.lower()
    return bool(re.match(r"^(test|modbattery|battery)", low)) or "_test" in low or "test_" in low


bases = [
    line.split("\t")
    for line in Path(D, "bases.tsv").read_text(encoding="utf-8").splitlines()
    if line.strip()
]
hits = []
modules_compare_db = collections.Counter()
modules_total = collections.Counter()
for app, repo, _ref in bases:
    src = os.path.join(D, "src", repo, "src")
    for root, _, files in os.walk(src):
        for fn in sorted(files):
            low = fn.lower()
            if is_test(fn) or not low.endswith((".bas", ".cls", ".form.txt", ".report.txt")):
                continue
            p = os.path.join(root, fn)
            rel = os.path.relpath(p, src).replace(os.sep, "/")
            lines, offset = code_lines(p)
            if not lines:
                continue
            modules_total[app] += 1
            if any(ln.strip().lower() == "option compare database" for ln in lines[:15]):
                modules_compare_db[app] += 1
            proc = None  # [name, first line, reads a maximum, adds 1]
            for n, line in logical_lines(lines):
                if not line.strip():
                    continue
                strs = strings_of(line)
                m = PROC_START.match(line)
                if m:
                    proc = [m.group(1), n + offset, False, False]
                    opened, reported, with_var = {}, set(), None
                elif PROC_END.match(line):
                    if proc and proc[2] and proc[3]:
                        hits.append((app, "max_mas_uno", rel, proc[1], "procedimiento " + proc[0]))
                    proc = None
                elif proc:
                    if any("max(" in s.lower() for s in strs) or DMAX.search(line):
                        proc[2] = True
                    if PLUS_ONE.search(line) and not COUNTER.match(line):
                        proc[3] = True
                    o = SET_RS.match(line)
                    if o and not SEE_CHANGES.search(line):
                        opened[o.group(1).lower()] = (n + offset, line.strip()[:200])
                    w = WITH.match(line)
                    if w:
                        with_var = w.group(1).lower()
                    e = EDITS.search(line)
                    if e:
                        var = (e.group(1) or with_var or "").lower()
                        if var in opened and var not in reported:
                            reported.add(var)
                            ln0, code0 = opened[var]
                            hits.append((app, "sin_dbseechanges_edita", rel, ln0, code0))
                for pid, _desc, test in PATTERNS:
                    if test(line, strs):
                        hits.append((app, pid, rel, n + offset, line.strip()[:200]))

os.makedirs(OUT, exist_ok=True)
with open(os.path.join(OUT, "patterns.tsv"), "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, delimiter="\t", lineterminator="\n")
    w.writerow(["app", "pattern", "file", "line", "code"])
    w.writerows(hits)

count = collections.Counter((h[0], h[1]) for h in hits)
files = collections.defaultdict(set)
for h in hits:
    files[(h[0], h[1])].add(h[2])
with open(os.path.join(OUT, "summary.tsv"), "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f, delimiter="\t", lineterminator="\n")
    w.writerow(["app", "pattern", "hits", "files"])
    for app, _repo, _ref in bases:
        for pid in ["max_mas_uno"] + [p[0] for p in PATTERNS] + ["sin_dbseechanges_edita"]:
            w.writerow([app, pid, count[(app, pid)], len(files[(app, pid)])])
        w.writerow([app, "option_compare_database", modules_compare_db[app], modules_total[app]])
print(len(hits), "hits;", dict(collections.Counter(h[1] for h in hits)))
