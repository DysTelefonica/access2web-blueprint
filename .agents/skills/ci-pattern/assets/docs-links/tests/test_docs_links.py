#!/usr/bin/env python3
# test_docs_links.py — enlaces locales de la documentación (HR-66, #712)
"""Enlaces locales, anclas, navegación hub<->spokes, retirados y baseline
shrink-only. Solo stdlib."""

import contextlib
import importlib.machinery
import importlib.util
import io
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
CLI_PATH = ASSET.parent / "bin" / "ci-pattern"


def _load(name, path, loader=None):
    spec = importlib.util.spec_from_file_location(name, path, loader=loader)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


GATE = _load("check_docs_links", ASSET / "check_docs_links.py")
CLI = _load("ci_pattern_cli_links", CLI_PATH,
            importlib.machinery.SourceFileLoader("ci_pattern_cli_links", str(CLI_PATH)))
HUB = "docs/CODEBASE-GUIDE.md"
FILES = {
    "README.md": "# P\n\nVea la [guía](docs/CODEBASE-GUIDE.md#páginas).\n",
    HUB: "# Guía\n\n## Páginas\n\n[Mapa](codebase/map.md)\n",
    "docs/codebase/map.md": "# Mapa\n\n[Volver](../CODEBASE-GUIDE.md)\n",
}
CONFIG = {"hub": HUB, "spokes": "docs/codebase"}


def _findings(files=None, config=None):
    with tempfile.TemporaryDirectory() as d:
        for rel, text in {**FILES, **(files or {})}.items():
            path = Path(d) / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        return GATE.check(Path(d), {**CONFIG, **(config or {})})


class DocsLinksTests(unittest.TestCase):
    def test_conforme_da_verde(self):
        self.assertEqual(_findings(), [])

    def test_destino_inexistente_nombra_fichero_linea_y_destino(self):
        found = _findings({"README.md": "# P\n\n[x](docs/ghost.md)\n"})
        self.assertTrue(any("README.md:3" in f and "docs/ghost.md" in f for f in found), found)

    def test_ancla_inexistente_en_destino_markdown_es_roja(self):
        found = _findings({"README.md": "# P\n\n[x](docs/CODEBASE-GUIDE.md#nada)\n"})
        self.assertTrue(any("#nada" in f for f in found), found)

    def test_remotos_mailto_y_anclas_de_la_misma_pagina_no_cuentan(self):
        text = "# P\n\n[a](https://e.com/x) [b](mailto:a@b.c) [c](#p) [d](tel:123)\n"
        self.assertEqual(_findings({"README.md": text}), [])

    def test_enlaces_dentro_de_un_bloque_cercado_no_cuentan(self):
        self.assertEqual(_findings({"README.md": "# P\n\n```\n[x](ghost.md)\n```\n"}), [])

    def test_spoke_sin_enlace_de_vuelta_al_hub_es_rojo(self):
        found = _findings({"docs/codebase/map.md": "# Mapa\n"})
        self.assertTrue(any("map.md" in f and "hub" in f for f in found), found)

    def test_spoke_no_enlazado_desde_el_hub_es_rojo(self):
        found = _findings({HUB: "# Guía\n\n## Páginas\n\nsin enlaces\n"})
        self.assertTrue(any("map.md" in f and "hub" in f for f in found), found)

    def test_enlaces_de_un_bloque_anidado_o_de_tildes_no_cuentan_y_los_de_despues_si(self):
        nested = "# P\n\n````md\n```sh\n```\n[x](ghost.md)\n````\n\n~~~\n[y](ghost.md)\n~~~\n"
        self.assertEqual(_findings({"README.md": nested}), [])
        after = "# P\n\n````md\n```sh\n```\n````\n\n[x](ghost.md)\n"
        self.assertTrue(any("README.md:8" in f for f in _findings({"README.md": after})))

    def test_fichero_retirado_presente_es_rojo(self):
        found = _findings(config={"retired": {"files": ["README.md"], "phrases": []}})
        self.assertTrue(any("README.md" in f and "retirado" in f for f in found), found)

    def test_frase_retirada_es_roja(self):
        retired = {"files": [], "phrases": [{"path": "README.md", "phrase": "Vea la"}]}
        found = _findings(config={"retired": retired})
        self.assertTrue(any("README.md" in f and "Vea la" in f for f in found), found)

    def test_baseline_tolera_la_deuda_y_su_entrada_obsoleta_es_stale(self):
        broken = {"README.md": "# P\n\n[x](ghost.md)\n"}
        debt = {"baseline": [{"file": "README.md", "target": "ghost.md"}]}
        self.assertEqual(_findings(broken, debt), [])
        found = _findings(None, debt)
        self.assertTrue(any("STALE" in f and "ghost.md" in f for f in found), found)

    def test_cero_ficheros_recorridos_falla_cerrado(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertTrue(any("ningún" in f for f in GATE.check(Path(d), {})))

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        buf = io.StringIO()
        with tempfile.TemporaryDirectory() as d, contextlib.redirect_stderr(buf):
            self.assertEqual(CLI._check_docs_links(Path(d), {}), [])
        self.assertIn("HR-66: docs_links no declarado", buf.getvalue())


if __name__ == "__main__":
    unittest.main()
