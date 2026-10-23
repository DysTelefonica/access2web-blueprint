#!/usr/bin/env python3
# ci-pattern asset tests — plantilla de notas de release (#1015)
"""Contrato de la plantilla de notas (HR-19, HR-21).

Las notas son prosa, pero el canon fija qué secciones tienen: son un contrato
que el workflow publica tal cual y que una release sin ellas no se puede
revisar. Este test fija las seis secciones y lo que la tabla de Procedencia
tiene que poder decir, sin juzgar el texto que un mantenedor escriba en ellas.

Ejecuta sin red y sin consumer.
"""
import re
import unittest
from pathlib import Path

TEMPLATE = Path(__file__).resolve().parent.parent / "RELEASE-NOTES.md.tmpl"
SECCIONES = ("Resumen", "Procedencia", "Qué cambia", "Cambios incompatibles",
             "Cómo actualizar", "Qué se arregla")


class TestReleaseNotesTemplate(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = TEMPLATE.read_text(encoding="utf-8")

    def test_declara_las_seis_secciones_del_canon(self):
        encabezados = set(re.findall(r"^#{1,3} (.+?)\s*$", self.text, re.MULTILINE))
        faltan = [s for s in SECCIONES if s not in encabezados]
        self.assertEqual(faltan, [], f"secciones que el canon exige y la plantilla no trae: {faltan}")

    def test_el_resumen_distingue_el_hotfix_de_la_release(self):
        resumen = self.text.split("## Procedencia")[0]
        self.assertIn("Hotfix", resumen)

    def test_la_procedencia_declara_version_sha_y_ejecucion_antes_y_despues(self):
        tabla = self.text.split("## Procedencia")[1].split("##")[0]
        for columna in ("| Anterior |", "| Nueva |"):
            self.assertIn(columna, tabla)
        self.assertIn("SHA", tabla)
        self.assertIn("Ejecución", tabla)

    def test_que_se_arregla_enlaza_la_causa_y_no_la_encierra(self):
        seccion = self.text.split("## Qué se arregla")[1]
        self.assertIn("post-mortem", seccion)
        self.assertIn("issue", seccion)
        # HR-21: la causa raíz vive en su documento; las notas solo la enlazan.
        self.assertIn("solo el", seccion)


if __name__ == "__main__":
    unittest.main()
