#!/usr/bin/env python3
"""Cada clave del host contract que leen los assets esta declarada en el schema (DysTelefonica/team-skills#956).

``check_host_drift.validate_contract`` deriva su conjunto ``known`` de las
``properties`` de ``host-contract.schema.json`` y falla cerrado (exit 2) ante
cualquier clave de primer nivel que no este ahi. Ese ``known`` es la UNICA
fuente de verdad, y por eso el schema es el sitio donde una clave opt-in nueva
tiene que declararse.

El hueco es facil de dejar: la clave se implementa y se documenta en SKILL.md,
con sus tests, y el gate la rechaza. Pasó con ``post_merge_dispatch``
(#946) y volví a pasar con ``governed_path`` (#909) y ``tracker_branch``
(#333): las tres están implementadas y documentadas, y ninguna declarada
en el schema.

Este test no hardcodea claves: recorre el AST de los assets que consumen el
host contract, recoge las claves de primer nivel que se LEEN de el, y
exige que esten todas declaradas. Asi, la proxima clave opt-in que se
anada al contrato queda cubierta por el invariante, no por un test escrito
a mano para ella.
"""
import ast
import json
import unittest
from pathlib import Path

# tests/ -> host-readback/ -> assets/
ASSETS = Path(__file__).resolve().parents[2]
SCHEMA = ASSETS / "host-readback" / "host-contract.schema.json"

# Assets que reciben el host contract como parametro `contract`, mas
# `validate_contract` (que lo recibe como `data` y define el vocabulario que el
# schema declara). No se audita `check_workflow_policy.py`: su parametro `data`
# es el POLICY file (HR-39), no el contrato del host.
CONSUMERS = (
    ASSETS / "host-readback" / "check_host_drift.py",
    ASSETS / "runner-controls" / "check_push_compliance.py",
    ASSETS / "runner-controls" / "check_governed_merge.py",
)

# Funciones cuyo parametro `data` SI es el host contract. El resto de las
# funciones con un parametro `data` de este repositorio reciben otra cosa
# (respuestas de la API de GitHub, el policy file, el manifiesto de gobierno).
DATA_IS_CONTRACT = {"validate_contract"}


def _contract_param_names(func: ast.FunctionDef) -> set[str]:
    """Nombres de los parametros por los que esta funcion recibe el contrato."""
    args = [a.arg for a in (*func.args.posonlyargs, *func.args.args, *func.args.kwonlyargs)]
    names = {a for a in args if a == "contract"}
    if func.name in DATA_IS_CONTRACT and "data" in args:
        names.add("data")
    return names


def _derived_keys(tree: ast.AST) -> set[str]:
    """Claves que el asset ESCRIBE en el dict del contrato al normalizarlo.

    `check_host_drift` anade `contract["runner_label"]` para poder comparar; no
    es una clave que el consumer declare, asi que no puede exigirse en el
    schema. Se detecta por la asignacion, no por una lista escrita a mano.
    """
    derived: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if (
                    isinstance(target, ast.Subscript)
                    and isinstance(target.value, ast.Name)
                    and target.value.id == "contract"
                    and isinstance(target.slice, ast.Constant)
                    and isinstance(target.slice.value, str)
                ):
                    derived.add(target.slice.value)
    return derived


def _read_keys(func: ast.FunctionDef, params: set[str]) -> set[str]:
    """Claves de primer nivel que la funcion LEE del contrato."""
    found: set[str] = set()
    for node in ast.walk(func):
        # contract.get("clave")
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and node.func.attr == "get"
            and isinstance(node.func.value, ast.Name)
            and node.func.value.id in params
            and node.args
            and isinstance(node.args[0], ast.Constant)
            and isinstance(node.args[0].value, str)
        ):
            found.add(node.args[0].value)
        # contract["clave"] en lectura (no en escritura)
        if (
            isinstance(node, ast.Subscript)
            and isinstance(node.value, ast.Name)
            and node.value.id in params
            and isinstance(node.slice, ast.Constant)
            and isinstance(node.slice.value, str)
        ):
            found.add(node.slice.value)
    return found


def _contract_keys_read() -> set[str]:
    """Union de las claves de primer nivel leidas del host contract."""
    keys: set[str] = set()
    for path in CONSUMERS:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        derived = _derived_keys(tree)
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef):
                params = _contract_param_names(node)
                if params:
                    keys |= _read_keys(node, params)
        keys -= derived
    return keys


class ContractKeysAreDeclaredTests(unittest.TestCase):
    def test_every_contract_key_read_by_the_assets_is_declared_in_the_schema(self):
        # RED si un asset lee una clave que el schema no declara: el gate
        # rechaza con 'contract has unknown keys' (exit 2) y el consumer no
        # puede declarar la opt-in que el canon documenta.
        declared = set(json.loads(SCHEMA.read_text(encoding="utf-8"))["properties"])
        read = _contract_keys_read()
        self.assertTrue(read, "el recorrido no encontro ninguna clave; el test estaria roto")
        missing = sorted(read - declared)
        self.assertEqual(
            missing,
            [],
            "claves que los assets leen del host contract y el schema no declara: "
            + ", ".join(missing)
            + " (DysTelefonica/team-skills#956). Declararlas en host-contract.schema.json"
            " o dejar de leerlas. Sin la declaracion, validate_contract rechaza el"
            " contrato con exit 2 y el consumer no puede usar la opt-in.",
        )

    def test_the_governed_path_and_tracker_branch_keys_are_declared(self):
        # Fixity explicita de las dos claves que #956 cierra. El test de arriba
        # ya las cubre por derivacion; este las nombra para que el fallo diga
        # que la regresion es de estas dos y no de una derivacion que se rompio.
        declared = set(json.loads(SCHEMA.read_text(encoding="utf-8"))["properties"])
        for key in ("governed_path", "tracker_branch"):
            with self.subTest(key=key):
                self.assertIn(key, declared)

    def test_derived_normalisation_keys_are_not_demanded(self):
        # Guarda contra un falso positivo del recorrido: `runner_label` se
        # escribe en el contrato al normalizarlo, no lo declara el consumer.
        declared = set(json.loads(SCHEMA.read_text(encoding="utf-8"))["properties"])
        self.assertNotIn("runner_label", _contract_keys_read())
        self.assertNotIn("runner_label", declared)


if __name__ == "__main__":
    unittest.main()
