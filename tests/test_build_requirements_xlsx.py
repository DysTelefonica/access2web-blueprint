"""The requirements inventory generator (issue #1494).

`scripts/build_requirements_xlsx.py` turns `docs/03-aplicaciones/<app>/requisitos.yaml`
into the Excel that Calidad fills in. The tests drive the CLI and read the
workbook back as XML, so they pin what Calidad sees, not how it is built.
"""

from __future__ import annotations

import copy
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any

import yaml

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "build_requirements_xlsx.py"

VALID: dict[str, Any] = {
    "aplicacion": "Gestión de Riesgos",
    "prefijo": "RIE",
    "repositorio": "DysTelefonica/GESTION_RIESGOS",
    "rama": "release/2026-008.2.rc1",
    "commit": "99ffa81df3f56a80e6f713e3afc0fb73fd9b3b25",
    "formularios": [
        {"interno": "Form_FormEdicion", "titulo": "Edición", "estado": "cubierto", "motivo": ""},
        {
            "interno": "Form_Aux",
            "titulo": "",
            "estado": "no aplicable",
            "motivo": "Subformulario sin comportamiento propio",
        },
    ],
    "features": [
        {
            "id": "RIE-F001",
            "area": "Ediciones",
            "nombre": "Publicar edición",
            "descripcion": "Deja la edición publicada.",
            "formulario": "Edición",
            "control": "Publicar edición",
            "tipo": "botón",
            "roles": ["Calidad"],
            "resultado": "La edición queda publicada.",
            "tecnico": {
                "formulario": "Form_FormEdicion",
                "control": "ComandoPublicar",
                "evidencia": ["src/forms/Form_FormEdicion.cls:120"],
            },
            "reglas": [
                {
                    "id": "RIE-F001-R01",
                    "descripcion": "Todos los riesgos tienen responsable.",
                    "evidencia": ["src/classes/Edicion.cls:340"],
                }
            ],
        }
    ],
}


def run(tmp_path: Path, data: Any) -> tuple[subprocess.CompletedProcess[str], Path]:
    source = tmp_path / "requisitos.yaml"
    source.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    out = tmp_path / "out.xlsx"
    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(source), "-o", str(out)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    return result, out


def sheet_rows(book: Path, name: str) -> list[list[str]]:
    with zipfile.ZipFile(book) as zf:
        workbook = zf.read("xl/workbook.xml").decode("utf-8")
        names = re.findall(r'<sheet name="([^"]+)"', workbook)
        xml = zf.read(f"xl/worksheets/sheet{names.index(name) + 1}.xml").decode("utf-8")
    rows = []
    for row in re.findall(r"<row[^>]*>(.*?)</row>", xml, re.S):
        rows.append([unescape(t) for t in re.findall(r"<t[^>]*>(.*?)</t>", row, re.S)])
    return rows


def unescape(text: str) -> str:
    return (
        text.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&amp;", "&")
    )


def test_valid_source_builds_the_three_sheets(tmp_path: Path) -> None:
    result, out = run(tmp_path, VALID)

    assert result.returncode == 0, result.stderr
    assert result.stderr == ""
    with zipfile.ZipFile(out) as zf:
        workbook = zf.read("xl/workbook.xml").decode("utf-8")
    assert re.findall(r'<sheet name="([^"]+)"', workbook) == ["Requisitos", "Formularios", "Origen"]


def test_feature_row_is_followed_by_its_rules(tmp_path: Path) -> None:
    _, out = run(tmp_path, VALID)

    rows = sheet_rows(out, "Requisitos")
    assert rows[0] == [
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
    ]
    assert rows[1] == [
        "RIE-F001",
        "Ediciones",
        "Publicar edición",
        "Deja la edición publicada.",
        "Edición",
        "Publicar edición",
        "Calidad",
        "La edición queda publicada.",
    ]
    assert rows[2] == [
        "RIE-F001-R01",
        "Ediciones",
        "Regla",
        "Todos los riesgos tienen responsable.",
    ]


def test_decision_lists_differ_for_features_and_rules(tmp_path: Path) -> None:
    _, out = run(tmp_path, VALID)

    with zipfile.ZipFile(out) as zf:
        xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
    validations = dict(re.findall(r'sqref="([^"]+)"><formula1>"([^"]+)"</formula1>', xml))
    assert validations == {"I2": "Mantener,Retirar", "I3": "Mantener,Cambiar,Quitar"}


def test_technical_names_stay_out_of_the_quality_sheet(tmp_path: Path) -> None:
    _, out = run(tmp_path, VALID)

    flat = " ".join(" ".join(r) for r in sheet_rows(out, "Requisitos"))
    assert "Form_FormEdicion" not in flat
    assert "ComandoPublicar" not in flat
    assert ".cls:" not in flat


def test_forms_and_origin_sheets_report_coverage(tmp_path: Path) -> None:
    _, out = run(tmp_path, VALID)

    assert sheet_rows(out, "Formularios")[1:] == [
        ["Form_FormEdicion", "Edición", "cubierto"],
        ["Form_Aux", "no aplicable", "Subformulario sin comportamiento propio"],
    ]
    origin = dict((r[0], r[1]) for r in sheet_rows(out, "Origen") if len(r) == 2)
    assert origin["Commit"] == VALID["commit"]
    assert origin["Formularios"] == "2 (1 cubiertos, 1 no aplicables)"
    assert origin["Funcionalidades"] == "1"
    assert origin["Reglas"] == "1"


def assert_rejected(tmp_path: Path, data: Any, message: str) -> None:
    result, out = run(tmp_path, data)
    assert result.returncode == 1
    assert message in result.stderr
    assert result.stdout == ""
    assert not out.exists()


def test_missing_required_field_is_rejected(tmp_path: Path) -> None:
    data = copy.deepcopy(VALID)
    del data["features"][0]["descripcion"]
    assert_rejected(tmp_path, data, "RIE-F001: falta el campo 'descripcion'")


def test_duplicate_id_is_rejected(tmp_path: Path) -> None:
    data = copy.deepcopy(VALID)
    data["features"][0]["reglas"][0]["id"] = "RIE-F001"
    assert_rejected(tmp_path, data, "identificador repetido: RIE-F001")


def test_rule_without_evidence_is_rejected(tmp_path: Path) -> None:
    data = copy.deepcopy(VALID)
    data["features"][0]["reglas"][0]["evidencia"] = []
    assert_rejected(tmp_path, data, "RIE-F001-R01: sin evidencia")


def test_not_applicable_form_without_reason_is_rejected(tmp_path: Path) -> None:
    data = copy.deepcopy(VALID)
    data["formularios"][1]["motivo"] = ""
    assert_rejected(tmp_path, data, "Form_Aux: 'no aplicable' sin motivo")


def test_unknown_form_state_is_rejected(tmp_path: Path) -> None:
    data = copy.deepcopy(VALID)
    data["formularios"][0]["estado"] = "pendiente"
    assert_rejected(tmp_path, data, "Form_FormEdicion: estado 'pendiente' no válido")


def test_covered_form_without_feature_is_rejected(tmp_path: Path) -> None:
    data = copy.deepcopy(VALID)
    data["features"][0]["tecnico"]["formulario"] = "Form_Otro"
    assert_rejected(
        tmp_path, data, "Form_FormEdicion: marcado como cubierto pero ninguna funcionalidad lo usa"
    )


def test_unreadable_yaml_exits_2(tmp_path: Path) -> None:
    source = tmp_path / "requisitos.yaml"
    source.write_text("features: [unclosed", encoding="utf-8")
    result = subprocess.run(
        [sys.executable, str(SCRIPT), str(source), "-o", str(tmp_path / "out.xlsx")],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    assert result.returncode == 2
    assert "no se puede leer" in result.stderr
