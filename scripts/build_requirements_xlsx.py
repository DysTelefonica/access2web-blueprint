#!/usr/bin/env python3
"""Build the Calidad requirements workbook from `docs/03-aplicaciones/<app>/requisitos.yaml`.

The schema and the writing rules live in `docs/07-migracion/inventario-requisitos.md`
(issue #1494). The workbook is written with the standard library only, so the
generator adds no dependency to the repository.

Exit codes:
    0  workbook written
    1  the source breaks the schema (one line per problem on stderr)
    2  the source cannot be read as YAML
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

import yaml

TOP_FIELDS = ("aplicacion", "prefijo", "repositorio", "rama", "commit", "formularios", "features")
FEATURE_FIELDS = ("id", "area", "nombre", "descripcion", "formulario", "tipo", "resultado")
FORM_STATES = ("cubierto", "no aplicable")
FEATURE_DECISIONS = "Mantener,Retirar"
RULE_DECISIONS = "Mantener,Cambiar,Quitar"
HEADERS = (
    "ID",
    "Área",
    "Funcionalidad / regla",
    "Descripción",
    "Formulario",
    "Botón o campo",
    "Quién",
    "Resultado",
    "Decisión",
    "Comentario",
)
WIDTHS = (14, 18, 34, 60, 28, 24, 18, 50, 14, 40)

# Cell styles, indexes into cellXfs of STYLES.
HEADER, FEATURE, RULE = 1, 2, 3


def missing(item: dict[str, Any], fields: tuple[str, ...]) -> list[str]:
    return [f for f in fields if not item.get(f)]


def validate_forms(forms: list[dict[str, Any]], used: set[str]) -> list[str]:
    errors = []
    for form in forms:
        name = form.get("interno", "?")
        state = form.get("estado")
        if state not in FORM_STATES:
            errors.append(f"{name}: estado '{state}' no válido (use {' o '.join(FORM_STATES)})")
        elif state == "no aplicable" and not form.get("motivo"):
            errors.append(f"{name}: 'no aplicable' sin motivo")
        elif state == "cubierto" and name not in used:
            errors.append(f"{name}: marcado como cubierto pero ninguna funcionalidad lo usa")
    return errors


def validate_feature(feature: dict[str, Any]) -> list[str]:
    fid = feature.get("id", "?")
    errors = [f"{fid}: falta el campo '{f}'" for f in missing(feature, FEATURE_FIELDS)]
    if not (feature.get("tecnico") or {}).get("evidencia"):
        errors.append(f"{fid}: sin evidencia")
    for rule in feature.get("reglas") or []:
        rid = rule.get("id", "?")
        errors += [f"{rid}: falta el campo '{f}'" for f in missing(rule, ("id", "descripcion"))]
        if not rule.get("evidencia"):
            errors.append(f"{rid}: sin evidencia")
    return errors


def validate(data: Any) -> list[str]:
    if not isinstance(data, dict):
        return ["el fichero no contiene un objeto YAML"]
    errors = [f"falta el campo '{f}'" for f in missing(data, TOP_FIELDS)]
    if errors:
        return errors
    seen: set[str] = set()
    used: set[str] = set()
    for feature in data["features"]:
        errors += validate_feature(feature)
        used.add((feature.get("tecnico") or {}).get("formulario", ""))
        for item_id in [feature.get("id")] + [r.get("id") for r in feature.get("reglas") or []]:
            if item_id in seen:
                errors.append(f"identificador repetido: {item_id}")
            seen.add(item_id)
    return errors + validate_forms(data["formularios"], used)


def requirement_rows(data: dict[str, Any]) -> list[tuple[int, list[str], str]]:
    """One row per feature followed by one row per rule: (style, cells, decision list)."""
    rows = []
    for f in data["features"]:
        roles = ", ".join(f.get("roles") or [])
        cells = [f["id"], f["area"], f["nombre"], f["descripcion"], f["formulario"]]
        cells += [f.get("control") or "", roles, f["resultado"]]
        rows.append((FEATURE, cells, FEATURE_DECISIONS))
        for r in f.get("reglas") or []:
            rows.append((RULE, [r["id"], f["area"], "Regla", r["descripcion"]], RULE_DECISIONS))
    return rows


def cell(ref: str, value: str, style: int) -> str:
    if not value:
        return f'<c r="{ref}" s="{style}"/>'
    text = escape(str(value))
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def sheet_xml(
    rows: list[tuple[int, list[str]]], widths: tuple[int, ...], validations: str = ""
) -> str:
    cols = "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, 1)
    )
    body = []
    for n, (style, cells) in enumerate(rows, 1):
        refs = "".join(cell(f"{chr(65 + i)}{n}", v, style) for i, v in enumerate(cells))
        body.append(f'<row r="{n}">{refs}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        f"<cols>{cols}</cols><sheetData>{''.join(body)}</sheetData>{validations}</worksheet>"
    )


def requirements_sheet(data: dict[str, Any]) -> str:
    rows = requirement_rows(data)
    table: list[tuple[int, list[str]]] = [(HEADER, list(HEADERS))]
    table += [(style, cells + [""] * (len(HEADERS) - len(cells))) for style, cells, _ in rows]
    checks = "".join(
        f'<dataValidation type="list" allowBlank="1" sqref="I{n}">'
        f'<formula1>"{choices}"</formula1></dataValidation>'
        for n, (_, _, choices) in enumerate(rows, 2)
    )
    return sheet_xml(
        table, WIDTHS, f'<dataValidations count="{len(rows)}">{checks}</dataValidations>'
    )


def forms_sheet(data: dict[str, Any]) -> str:
    table: list[tuple[int, list[str]]] = [
        (HEADER, ["Formulario interno", "Título", "Estado", "Motivo"])
    ]
    for f in data["formularios"]:
        table.append(
            (RULE, [f["interno"], f.get("titulo") or "", f["estado"], f.get("motivo") or ""])
        )
    return sheet_xml(table, (40, 40, 16, 60))


def origin_sheet(data: dict[str, Any]) -> str:
    forms = data["formularios"]
    covered = sum(1 for f in forms if f["estado"] == "cubierto")
    rules = sum(len(f.get("reglas") or []) for f in data["features"])
    table: list[tuple[int, list[str]]] = [
        (HEADER, ["Campo", "Valor"]),
        (RULE, ["Aplicación", data["aplicacion"]]),
        (RULE, ["Repositorio", data["repositorio"]]),
        (RULE, ["Rama", data["rama"]]),
        (RULE, ["Commit", data["commit"]]),
        (
            RULE,
            [
                "Formularios",
                f"{len(forms)} ({covered} cubiertos, {len(forms) - covered} no aplicables)",
            ],
        ),
        (RULE, ["Funcionalidades", str(len(data["features"]))]),
        (RULE, ["Reglas", str(rules)]),
    ]
    return sheet_xml(table, (20, 60))


STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="4"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill></fills>'
    '<borders count="1"><border/></borders>'
    '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
    '<cellXfs count="4"><xf/>'
    '<xf fontId="1" fillId="2" applyFont="1" applyFill="1"/>'
    '<xf fontId="1" fillId="3" applyFont="1" applyFill="1" applyAlignment="1">'
    '<alignment wrapText="1" vertical="top"/></xf>'
    '<xf applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)


XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
OFFICE = "application/vnd.openxmlformats-officedocument.spreadsheetml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS_OPEN = (
    f'{XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
)


def relationship(rid: int, kind: str, target: str) -> str:
    return f'<Relationship Id="rId{rid}" Type="{REL}/{kind}" Target="{target}"/>'


def override(part: str, content_type: str) -> str:
    return f'<Override PartName="{part}" ContentType="{OFFICE}.{content_type}+xml"/>'


def write_workbook(path: Path, sheets: list[tuple[str, str]]) -> None:
    numbered = list(enumerate(sheets, 1))
    content_types = (
        f'{XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + override("/xl/workbook.xml", "sheet.main")
        + override("/xl/styles.xml", "styles")
        + "".join(override(f"/xl/worksheets/sheet{i}.xml", "worksheet") for i, _ in numbered)
        + "</Types>"
    )
    workbook = (
        f'{XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'xmlns:r="{REL}"><sheets>'
        + "".join(f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, (n, _) in numbered)
        + "</sheets></workbook>"
    )
    workbook_rels = (
        RELS_OPEN
        + "".join(relationship(i, "worksheet", f"worksheets/sheet{i}.xml") for i, _ in numbered)
        + relationship(len(sheets) + 1, "styles", "styles.xml")
        + "</Relationships>"
    )
    files = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": RELS_OPEN
        + relationship(1, "officeDocument", "xl/workbook.xml")
        + "</Relationships>",
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/styles.xml": STYLES,
    }
    files.update({f"xl/worksheets/sheet{i}.xml": xml for i, (_, xml) in numbered})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("source", type=Path, help="docs/03-aplicaciones/<app>/requisitos.yaml")
    parser.add_argument(
        "-o", "--output", type=Path, required=True, help="workbook to write (.xlsx)"
    )
    args = parser.parse_args(argv)
    # The messages carry accents; a Windows console defaults to cp1252.
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    try:
        data = yaml.safe_load(args.source.read_text(encoding="utf-8"))
    except (OSError, yaml.YAMLError) as exc:
        print(f"{args.source}: no se puede leer: {exc}", file=sys.stderr)
        return 2
    errors = validate(data)
    if errors:
        for error in errors:
            print(f"{args.source}: {error}", file=sys.stderr)
        return 1
    write_workbook(
        args.output,
        [
            ("Requisitos", requirements_sheet(data)),
            ("Formularios", forms_sheet(data)),
            ("Origen", origin_sheet(data)),
        ],
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
