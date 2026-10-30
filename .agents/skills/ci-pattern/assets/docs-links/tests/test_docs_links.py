#!/usr/bin/env python3
# test_docs_links.py — enlaces locales de la documentación (HR-66, #712)
"""Enlaces locales, anclas, navegación hub<->spokes, retirados, baseline
shrink-only y alcanzabilidad de la documentación (HR-80, #1379). Solo stdlib."""

import contextlib
import importlib.machinery
import importlib.util
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
CLI_PATH = ASSET.parent / "bin" / "ci-pattern"
JOB = ASSET / "run_docs_links.py"
CONTRACT = ".github/ci-pattern-adoption.json"


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


def _job_repo(decl=None, files=None):
    """Repo temporal con el contrato de adopción y `decl` en `docs_links`."""
    tmp = tempfile.TemporaryDirectory()
    repo = Path(tmp.name)
    for rel, text in {**FILES, **(files or {})}.items():
        path = repo / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
    (repo / CONTRACT).parent.mkdir(parents=True, exist_ok=True)
    (repo / CONTRACT).write_text(json.dumps(
        {"docs_links": dict(CONFIG, job="docs-links", **(decl or {}))}), encoding="utf-8")
    return tmp, repo


def _job(repo, *args):
    proc = subprocess.run([sys.executable, str(JOB), str(repo), *args],
                          capture_output=True, text=True, timeout=60, check=False)
    return proc.returncode, proc.stdout, proc.stderr


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

    def test_el_job_lee_el_contrato_de_adopcion_sin_fichero_en_la_raiz(self):
        tmp, repo = _job_repo()
        with tmp:
            code, out, _ = _job(repo)
        self.assertEqual(code, 0, out)
        self.assertIn("docs-links: OK", out)

    def test_el_contrato_manda_y_no_el_fichero_de_la_raiz(self):
        # El fichero de raiz declara otro hub: si el job lo leyera, seria rojo.
        tmp, repo = _job_repo(files={"docs-links.json": json.dumps({"hub": "docs/ghost.md",
                                                                    "spokes": "docs/codebase"})})
        with tmp:
            code, out, _ = _job(repo)
        self.assertEqual(code, 0, out)
        self.assertIn("docs-links: OK", out)

    def test_config_obsoleto_aun_funciona_y_avisa_por_stderr(self):
        tmp, repo = _job_repo(files={"viejo.json": json.dumps(CONFIG)})
        with tmp:
            code, out, err = _job(repo, "--config", "viejo.json")
        self.assertEqual(code, 0, out)
        self.assertIn("docs-links: OK", out)
        self.assertIn("obsoleto", err)

    def test_config_obsoleto_inexistente_o_malformado_sale_2(self):
        tmp, repo = _job_repo(files={"roto.json": "{no es json"})
        with tmp:
            self.assertEqual(_job(repo, "--config", "ghost.json")[0], 2)
            self.assertEqual(_job(repo, "--config", "roto.json")[0], 2)

    def test_contrato_inexistente_sale_2(self):
        tmp, repo = _job_repo()
        with tmp:
            code, _, err = _job(repo, "--contract", ".github/no-existe.json")
        self.assertEqual(code, 2)
        self.assertIn("no-existe.json", err)

    def test_declaracion_que_no_es_objeto_sale_2_y_no_revienta(self):
        tmp, repo = _job_repo()
        with tmp:
            (repo / CONTRACT).write_text(json.dumps({"docs_links": None}), encoding="utf-8")
            code, _, err = _job(repo)
        self.assertEqual(code, 2)
        self.assertIn("no es un objeto", err)
        self.assertNotIn("Traceback", err)

    def test_sin_declaracion_en_el_contrato_el_job_no_inventa_configuracion(self):
        tmp, repo = _job_repo()
        with tmp:
            (repo / CONTRACT).write_text(json.dumps({}), encoding="utf-8")
            code, _, err = _job(repo)
        self.assertEqual(code, 2)
        self.assertIn("docs_links", err)

    def test_fichero_de_raiz_residual_se_nombra_como_segunda_fuente(self):
        tmp, repo = _job_repo(files={"docs-links.json": json.dumps(CONFIG)})
        with tmp:
            found = CLI._check_docs_links(repo, {"docs_links": dict(CONFIG, job="docs-links")})
        self.assertTrue(any("docs-links.json" in f for f in found), found)

    def test_sin_fichero_de_raiz_no_hay_hallazgo_de_segunda_fuente(self):
        tmp, repo = _job_repo()
        with tmp:
            found = CLI._check_docs_links(repo, {"docs_links": dict(CONFIG, job="docs-links")})
        self.assertFalse(any("segunda fuente" in f for f in found), found)


HUB_LINKS = ("# Guía\n\n## Páginas\n\n[Mapa](codebase/map.md)\n"
             "[Decisión](adr/001-decision.md)\n")
REACH_FILES = {
    "README.md": "# P\n\nVea la [guía](docs/CODEBASE-GUIDE.md#páginas).\n",
    HUB: HUB_LINKS,
    "docs/codebase/map.md": "# Mapa\n\n[Volver](../CODEBASE-GUIDE.md)\n",
    "docs/adr/001-decision.md": ("# Decisión\n\n[Transitiva](002-transitiva.md)\n"
                                 "[Volver](../CODEBASE-GUIDE.md)\n"),
    "docs/adr/002-transitiva.md": "# Transitiva\n\n[Volver](../CODEBASE-GUIDE.md)\n",
    "docs/adr/003-huerfano.md": "# Huérfano\n\n[Volver](../CODEBASE-GUIDE.md)\n",
}
REACH = {"from": [HUB], "scope": ["docs"]}


def _reach_findings(files=None, config=None, declaration=None):
    """Hallazgos del gate con la declaración `reachable` activa."""
    with tempfile.TemporaryDirectory() as d:
        for rel, text in {**REACH_FILES, **(files or {})}.items():
            path = Path(d) / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        reachable = dict(REACH, **(config or {})) if declaration is None else declaration
        return GATE.check(Path(d), {**CONFIG, "reachable": reachable})


class DocsLinksReachableTests(unittest.TestCase):
    """HR-80 (#1379): todo .md del alcance declarado tiene que ser alcanzable
    desde los puntos de entrada siguiendo enlaces markdown."""

    def test_huerfano_falla_nombrando_el_fichero(self):
        found = _reach_findings()
        self.assertTrue(any("docs/adr/003-huerfano.md" in f for f in found), found)

    def test_enlazado_directamente_pasa(self):
        hub = HUB_LINKS + "[Huérfano](adr/003-huerfano.md)\n"
        found = _reach_findings({HUB: hub})
        self.assertEqual([f for f in found if "003-huerfano" in f], [])

    def test_enlazado_transitivamente_pasa(self):
        """El hub no enlaza la segunda decisión: la primera la lleva."""
        found = _reach_findings()
        self.assertFalse(any("002-transitiva" in f for f in found), found)

    def test_excepcion_declarada_por_fichero_pasa(self):
        found = _reach_findings(config={"exceptions": ["docs/adr/003-huerfano.md"]})
        self.assertEqual([f for f in found if "003-huerfano" in f], [])

    def test_excepcion_declarada_por_directorio_pasa(self):
        found = _reach_findings(config={"exceptions": ["docs/adr/"]})
        self.assertEqual([f for f in found if "003-huerfano" in f], [])

    def test_excepcion_declarada_por_glob_pasa(self):
        found = _reach_findings(config={"exceptions": ["docs/adr/*.md"]})
        self.assertEqual([f for f in found if "003-huerfano" in f], [])

    def test_sin_declaracion_no_cambia_nada(self):
        """Sin `reachable` el gate mide lo que medía: el huérfano no se ve."""
        with tempfile.TemporaryDirectory() as d:
            for rel, text in REACH_FILES.items():
                path = Path(d) / rel
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(text, encoding="utf-8")
            self.assertEqual(GATE.check(Path(d), CONFIG), [])

    def test_enlace_roto_lo_sigue_cazando_el_gate_que_ya_lo_cazaba(self):
        found = _reach_findings({"README.md": "# P\n\n[x](docs/ghost.md)\n"})
        self.assertTrue(any("ghost.md" in f for f in found), found)
        self.assertTrue(any("003-huerfano" in f for f in found), found)

    def test_declaracion_incompleta_falla_nombrando_la_clave(self):
        found = _reach_findings(declaration={})
        self.assertTrue(any("from" in f for f in found), found)

    def test_declaracion_que_no_es_objeto_falla(self):
        found = _reach_findings(declaration=[])
        self.assertTrue(any("reachable" in f for f in found), found)

    def test_punto_de_entrada_inexistente_falla(self):
        found = _reach_findings(declaration={"from": ["docs/no-existe.md"], "scope": ["docs"]})
        self.assertTrue(any("no-existe.md" in f for f in found), found)

    def test_alcance_que_no_existe_falla_cerrado(self):
        found = _reach_findings(config={"scope": ["docs/no-existe"]})
        self.assertTrue(any("no-existe" in f for f in found), found)

    def test_alcance_fuera_del_repositorio_falla(self):
        found = _reach_findings(config={"scope": ["/etc"]})
        self.assertTrue(any("/etc" in f for f in found), found)

    def test_el_alcance_por_defecto_es_docs(self):
        found = _reach_findings(declaration={"from": [HUB]})
        self.assertTrue(any("003-huerfano" in f for f in found), found)

    def test_un_fichero_fuera_del_alcance_no_se_exige(self):
        files = {"notas/004-fuera.md": "# Fuera\n"}
        found = _reach_findings(files)
        self.assertFalse(any("004-fuera" in f for f in found), found)

    def test_un_fichero_que_no_es_markdown_no_se_exige(self):
        files = {"docs/datos.txt": "texto\n"}
        found = _reach_findings(files)
        self.assertFalse(any("datos.txt" in f for f in found), found)

    def test_los_hallazgos_salen_ordenados_por_ruta(self):
        files = {"docs/aaa-huerfano.md": "# A\n", "docs/zzz-huerfano.md": "# Z\n"}
        found = [f for f in _reach_findings(files) if "huérfano" in f]
        self.assertEqual(len(found), 3, found)
        self.assertTrue(found[0].split()[1] < found[1].split()[1] < found[2].split()[1], found)

    def test_la_cli_acepta_la_clave_reachable(self):
        with tempfile.TemporaryDirectory() as d:
            for rel, text in REACH_FILES.items():
                path = Path(d) / rel
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(text, encoding="utf-8")
            found = CLI._check_docs_links(Path(d), {"docs_links": {**CONFIG, "reachable": REACH}})
        self.assertFalse(any("clave desconocida" in f for f in found), found)
        self.assertTrue(any("003-huerfano" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
