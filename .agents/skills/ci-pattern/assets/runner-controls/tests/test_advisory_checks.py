#!/usr/bin/env python3
# ci-pattern asset tests -- los checks advisory declarados por el consumer (#1762)
"""`advisory_checks` en `.github/host-contract.json`: los checks que no bloquean.

Sin la clave, el comportamiento es el de siempre. Un check requerido nunca
queda tapado por un patron advisory. Una declaracion invalida falla cerrado.
"""
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from advisory_checks import (  # noqa: E402
    AdvisoryContractError,
    advisory_patterns,
    is_advisory,
    read_advisory_patterns,
    read_contract,
    split_runs,
)


def run(name: str, conclusion: str | None = "success") -> dict:
    return {"id": abs(hash(name)) % 10000, "name": name, "status": "completed",
            "conclusion": conclusion}


class AdvisoryChecksTests(unittest.TestCase):
    """La declaracion se lee del host-contract y tiene una forma valida o falla."""

    def test_sin_clave_no_hay_patrones(self) -> None:
        self.assertEqual(advisory_patterns({"contract_version": 1}), ())

    def test_una_lista_de_globs_se_lee(self) -> None:
        self.assertEqual(advisory_patterns({"advisory_checks": ["*-gh", "docs-gh"]}),
                         ("*-gh", "docs-gh"))

    def test_una_forma_invalida_falla_cerrado(self) -> None:
        for value in ("*-gh", 3, [""], [" "], [None], [3], {"a": 1}):
            with self.subTest(value=value):
                with self.assertRaises(AdvisoryContractError):
                    advisory_patterns({"advisory_checks": value})

    def test_un_fichero_que_no_existe_no_declara_nada(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(read_advisory_patterns(Path(d) / "host-contract.json"), ())

    def test_un_fichero_ilegible_falla_cerrado(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "host-contract.json"
            path.write_text("{not json", encoding="utf-8")
            with self.assertRaises(AdvisoryContractError):
                read_advisory_patterns(path)

    def test_un_contrato_que_no_es_utf8_falla_cerrado(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "host-contract.json"
            path.write_bytes(b'{"advisory_checks": ["\xff"]}')
            with self.assertRaises(AdvisoryContractError):
                read_advisory_patterns(path)

    def test_el_nombre_casa_con_fnmatch(self) -> None:
        self.assertTrue(is_advisory("docs-gh", ("*-gh",)))
        self.assertFalse(is_advisory("docs", ("*-gh",)))

    def test_un_advisory_no_se_juzga(self) -> None:
        blocking, advisory = split_runs([run("unit"), run("docs-gh", "failure")],
                                        ("*-gh",), required=())
        self.assertEqual([r["name"] for r in blocking], ["unit"])
        self.assertEqual([r["name"] for r in advisory], ["docs-gh"])

    def test_un_requerido_nunca_queda_tapado(self) -> None:
        blocking, advisory = split_runs([run("unit", "failure")], ("unit",), required={"unit"})
        self.assertEqual([r["name"] for r in blocking], ["unit"])
        self.assertEqual(advisory, [])


class RequiredChecksFailClosedTests(unittest.TestCase):
    """#1795: un `required_checks` mal formado falla cerrado; sin la clave, no hay requeridos.

    La regla es la de `check_push_compliance.required_checks()`: lista, cada entrada un
    objeto con `name` cadena no vacia. Una entrada rota no se descarta en silencio: si se
    descartara, un patron advisory podria tapar ese requerido."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def read(self, required: object) -> object:
        path = self.root / "host-contract.json"
        path.write_text(json.dumps({"contract_version": 1, "advisory_checks": ["*-gh"],
                                    "required_checks": required}), encoding="utf-8")
        return read_contract(path)

    def test_clave_ausente_significa_sin_requeridos(self) -> None:
        path = self.root / "host-contract.json"
        path.write_text(json.dumps({"contract_version": 1, "advisory_checks": ["*-gh"]}),
                        encoding="utf-8")
        self.assertEqual(read_contract(path).required, ())

    def test_contrato_valido_no_cambia(self) -> None:
        declared = self.read([{"name": "unit", "class": "documented-only"},
                              {"name": "pr-gate"}])
        self.assertEqual(declared.required, ("unit", "pr-gate"))
        self.assertEqual(declared.patterns, ("*-gh",))

    def test_entrada_que_no_es_objeto_falla_cerrado(self) -> None:
        with self.assertRaises(AdvisoryContractError) as ctx:
            self.read(["unit"])
        self.assertIn("required_checks[0]", str(ctx.exception))

    def test_entrada_sin_name_falla_cerrado(self) -> None:
        with self.assertRaises(AdvisoryContractError) as ctx:
            self.read([{"class": "documented-only"}])
        self.assertIn("required_checks[0]", str(ctx.exception))

    def test_entrada_con_name_vacio_falla_cerrado(self) -> None:
        with self.assertRaises(AdvisoryContractError) as ctx:
            self.read([{"name": "unit"}, {"name": ""}])
        self.assertIn("required_checks[1]", str(ctx.exception))

    def test_required_checks_no_lista_falla_cerrado(self) -> None:
        with self.assertRaises(AdvisoryContractError) as ctx:
            self.read({"name": "unit"})
        self.assertIn("required_checks", str(ctx.exception))

    def test_required_checks_null_falla_cerrado(self) -> None:
        # `null` no es una lista: la misma regla que check_push_compliance (que usa .get con []).
        with self.assertRaises(AdvisoryContractError) as ctx:
            self.read(None)
        self.assertIn("required_checks", str(ctx.exception))

    def test_nombre_solo_con_espacios_falla_cerrado(self) -> None:
        with self.assertRaises(AdvisoryContractError) as ctx:
            self.read([{"name": "   "}])
        self.assertIn("required_checks[0]", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
