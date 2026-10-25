#!/usr/bin/env python3
# ci-pattern asset tests — plantilla de notas de release (#1015, #1148)
"""Contrato de la plantilla de notas (HR-19, HR-21).

Las notas son prosa, pero el canon fija qué secciones tienen y en qué orden: son
un contrato que el workflow publica tal cual y que una release sin ellas no se
puede revisar. La estructura es la del informe de Cadete v1.0.1 (#1148): una
release de un web project sale igual en cualquier repositorio. Este test fija
las secciones, su orden, las tres líneas fijas de cada corrección, las cinco
categorías del registro de cierres y lo que la tabla de Procedencia tiene que
poder decir, sin juzgar el texto que un mantenedor escriba en ellas.

Ejecuta sin red y sin consumer.
"""
import re
import unittest
from pathlib import Path

ASSETS = Path(__file__).resolve().parent.parent
TEMPLATE = ASSETS / "RELEASE-NOTES.md.tmpl"
RUNBOOK = ASSETS.parent.parent / "references" / "release-runbook.md"
RELEASE_WORKFLOW = ASSETS / ".github" / "workflows" / "release.yml.tmpl"

# Orden y nombres exactos del informe de Cadete v1.0.1 (#1148).
SECCIONES = ("Resumen", "Procedencia", "Qué se corrige", "Cambios incompatibles",
             "Cómo actualizar / desplegar", "Cifras", "Registro de cierres")
LINEAS_FIJAS = ("**Qué fallaba.**", "**Qué cambia.**", "**Qué lo prueba.**")
CATEGORIAS_CIERRE = ("Cerradas con evidencia", "Corregidas, sin test que lo fije",
                     "No cerradas por esta versión", "Diferido", "No abordado")
FILAS_PROCEDENCIA = ("Release anterior", "Objetivo", "CI de `main` sobre ese SHA",
                     "Gate `check_release_tag.py`", "PR incluidas")


def _secciones(text):
    """Encabezados de segundo nivel, en orden de aparición."""
    return re.findall(r"^## (.+?)\s*$", text, re.MULTILINE)


def _bloque(text, titulo):
    """Texto entre un encabezado `## titulo` y el siguiente `##`."""
    partes = text.split(f"## {titulo}\n", 1)
    if len(partes) < 2:
        return ""
    return re.split(r"^## ", partes[1], maxsplit=1, flags=re.MULTILINE)[0]


class TestReleaseNotesTemplate(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = TEMPLATE.read_text(encoding="utf-8")

    def test_declara_las_siete_secciones_en_el_orden_de_cadete(self):
        self.assertEqual(_secciones(self.text), list(SECCIONES))

    def test_quitar_o_renombrar_una_seccion_rompe_el_contrato(self):
        for quitada in SECCIONES:
            with self.subTest(seccion=quitada):
                sin = self.text.replace(f"## {quitada}\n", "## Otra cosa\n", 1)
                self.assertNotEqual(_secciones(sin), list(SECCIONES))

    def test_el_resumen_distingue_el_hotfix_de_la_release(self):
        self.assertIn("Hotfix", _bloque(self.text, "Resumen"))

    def test_la_procedencia_declara_las_filas_del_informe(self):
        tabla = _bloque(self.text, "Procedencia")
        self.assertIn("| Concepto | Valor |", tabla)
        for fila in FILAS_PROCEDENCIA:
            with self.subTest(fila=fila):
                self.assertIn(f"| {fila} |", tabla)
        self.assertIn("SHA", tabla)

    def test_cada_correccion_lleva_las_tres_lineas_fijas(self):
        corrige = _bloque(self.text, "Qué se corrige")
        self.assertTrue(re.search(r"^### .+$", corrige, re.MULTILINE), "falta al menos una subsección por PR o issue")
        for linea in LINEAS_FIJAS:
            with self.subTest(linea=linea):
                self.assertIn(f"- {linea}", corrige)

    def test_que_se_corrige_enlaza_la_causa_del_hotfix_y_no_la_encierra(self):
        corrige = _bloque(self.text, "Qué se corrige")
        # HR-21: la causa raíz vive en su documento; las notas solo la enlazan.
        self.assertIn("post-mortem", corrige)
        self.assertIn("solo", corrige)

    def test_el_registro_de_cierres_tiene_las_cinco_categorias_en_orden(self):
        registro = _bloque(self.text, "Registro de cierres")
        encontradas = re.findall(r"^### (.+?)\s*$", registro, re.MULTILINE)
        self.assertEqual(encontradas, list(CATEGORIAS_CIERRE))

    def test_cifras_declara_tests_cobertura_y_duracion(self):
        cifras = _bloque(self.text, "Cifras")
        for marca in ("Tests:", "Cobertura de líneas:", "Duración del CI"):
            with self.subTest(marca=marca):
                self.assertIn(marca, cifras)

    def test_como_actualizar_incluye_el_rollback(self):
        self.assertIn("Rollback:", _bloque(self.text, "Cómo actualizar / desplegar"))

    def test_runbook_y_workflow_apuntan_a_la_misma_ruta(self):
        # La convención vigente es la raíz `RELEASE-<TAG>.md`: el workflow la lee con
        # `--notes-file` y el runbook la describe. Si alguna cambia, este test falla.
        runbook = RUNBOOK.read_text(encoding="utf-8")
        workflow = RELEASE_WORKFLOW.read_text(encoding="utf-8")
        self.assertIn("RELEASE-<TAG>.md", runbook)
        self.assertIn('RELEASE-$GITHUB_REF_NAME.md', workflow)

    def test_runbook_describe_las_secciones_de_la_plantilla(self):
        runbook = RUNBOOK.read_text(encoding="utf-8")
        for seccion in SECCIONES:
            with self.subTest(seccion=seccion):
                self.assertIn(seccion, runbook)


if __name__ == "__main__":
    unittest.main()
