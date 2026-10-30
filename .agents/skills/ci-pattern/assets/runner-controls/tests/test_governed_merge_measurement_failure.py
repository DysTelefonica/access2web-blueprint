#!/usr/bin/env python3
# ci-pattern asset tests — un fallo de medicion no interrumpe el trabajo post-merge (#1279)
"""#1279: el fallo de la medicion del arbol integrado no puede saltarse el
cierre de las issues enlazadas ni el dispatch post-merge.

El merge ya esta hecho cuando la medicion corre (#1265). Si la lectura del
arbol falla —un 5xx, un timeout, un 404— el asset mide lo que puede, guarda
la causa, sigue con el cierre y el dispatch, y sale con 2 al final nombrando
la causa. Antes de #1279 el `FailClosedError` subia hasta `main()` y el cierre
y el dispatch no se ejecutaban nunca: el canal de evidencia post-merge
desaparecia justo cuando mas falta hace."""
import json
import tempfile
import unittest
from pathlib import Path

import _github_stub
from test_governed_merge import run_gate
from test_governed_merge_close import ref
from test_governed_merge_post_dispatch import contract_with, with_yaml, yaml_with_inputs

DISPATCH_WORKFLOW = "push-compliance"


class MeasurementFailureTests(unittest.TestCase):
    """#1279: un fallo de medicion no interrumpe el trabajo post-merge."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def state(self, *, measurement_fails: bool) -> dict:
        """PR de punta de cadena que cierra #305, con dispatch post-merge.

        Con ``measurement_fails`` la lectura del arbol del commit de merge
        responde 404: la medicion no se puede hacer, pero el merge ya esta.
        """
        state = with_yaml(DISPATCH_WORKFLOW, yaml_with_inputs(["before", "after"]))
        state["closing_refs"] = [ref(305)]
        if measurement_fails:
            del state["commits"][_github_stub.MERGED_SHA]
        return state

    def run_merged(self, *, measurement_fails: bool, summary: str = "") -> object:
        return run_gate(self.root, contract_with([DISPATCH_WORKFLOW]),
                        self.state(measurement_fails=measurement_fails), summary=summary)

    def test_el_fallo_de_medicion_no_impide_cerrar_la_issue_ni_despachar(self) -> None:
        # Criterio 1: el merge esta hecho; la issue enlazada se cierra y el
        # dispatch post-merge se dispara aunque la medicion haya fallado.
        proc = self.run_merged(measurement_fails=True)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        stored = self.stored()
        self.assertTrue(stored["merges"])
        self.assertEqual(stored["closed_issues"],
                         [{"number": 305, "state": "closed", "state_reason": "completed"}])
        self.assertEqual([item["name"] for item in stored["dispatches"]],
                         [f"{DISPATCH_WORKFLOW}.yml"])

    def test_el_fallo_de_medicion_sale_con_2_nombrando_la_causa(self) -> None:
        # Criterio 2: el codigo es 2 y el mensaje nombra la causa real.
        proc = self.run_merged(measurement_fails=True)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("was not measured", proc.stderr)
        self.assertIn("404", proc.stderr)

    def test_el_resumen_dice_not_measured_y_no_afirma_una_coincidencia(self) -> None:
        # Criterio 3: la causa viaja al resumen del job y el resumen nunca
        # afirma que el arbol coincide cuando no se ha medido.
        summary = self.root / "summary.md"
        proc = self.run_merged(measurement_fails=True, summary=str(summary))
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        text = summary.read_text(encoding="utf-8")
        self.assertIn("not measured:", text)
        self.assertIn("404", text)
        self.assertNotIn("the integrated tree is the tree CI reviewed", text)
        self.assertNotIn("the integrated tree is NOT the tree CI reviewed", text)

    def test_el_cierre_fallido_sigue_dando_3_aunque_la_medicion_fallara(self) -> None:
        # Criterio 4: la precedencia de 3 y 4 sobre el 2 no cambia.
        state = self.state(measurement_fails=True)
        state["close_error"] = True
        proc = run_gate(self.root, contract_with([DISPATCH_WORKFLOW]), state)
        self.assertEqual(proc.returncode, 3, proc.stdout + proc.stderr)
        self.assertIn("MERGED but not closed", proc.stderr)

    def test_el_dispatch_fallido_sigue_dando_4_aunque_la_medicion_fallara(self) -> None:
        state = self.state(measurement_fails=True)
        state["dispatch_error"] = True
        proc = run_gate(self.root, contract_with([DISPATCH_WORKFLOW]), state)
        self.assertEqual(proc.returncode, 4, proc.stdout + proc.stderr)
        self.assertIn("MERGED but not dispatched", proc.stderr)

    def test_la_medicion_sigue_despues_del_merge_y_antes_del_cierre_y_el_dispatch(self) -> None:
        # Criterio 5: la medicion conserva su sitio, para que la evidencia se
        # publique antes de que un paso posterior pueda fallar.
        proc = self.run_merged(measurement_fails=True)
        out = proc.stdout
        self.assertLess(out.index("INTEGRATED TREE"), out.index("CLOSED: #305"))
        self.assertLess(out.index("CLOSED: #305"), out.index("DISPATCHED:"))

    def test_un_resumen_inalcanzable_tampoco_interrumpe_el_trabajo_post_merge(self) -> None:
        # Criterio 6: ni la lectura ni la escritura del resumen pueden
        # interrumpir el cierre y el dispatch.
        summary = str(self.root / "no" / "summary.md")
        proc = self.run_merged(measurement_fails=False, summary=summary)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("GITHUB_STEP_SUMMARY", proc.stderr)
        stored = self.stored()
        self.assertEqual(stored["closed_issues"],
                         [{"number": 305, "state": "closed", "state_reason": "completed"}])
        self.assertEqual([item["name"] for item in stored["dispatches"]],
                         [f"{DISPATCH_WORKFLOW}.yml"])

    def test_sin_fallo_de_medicion_el_comportamiento_no_cambia(self) -> None:
        # Criterio 7: con la medicion en verde, todo sigue igual: 0, la issue
        # cerrada y el dispatch disparado.
        proc = self.run_merged(measurement_fails=False)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        stored = self.stored()
        self.assertEqual(stored["closed_issues"],
                         [{"number": 305, "state": "closed", "state_reason": "completed"}])
        self.assertEqual([item["name"] for item in stored["dispatches"]],
                         [f"{DISPATCH_WORKFLOW}.yml"])


if __name__ == "__main__":
    unittest.main()
