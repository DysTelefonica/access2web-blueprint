#!/usr/bin/env python3
# test_docs_decisions.py — registro de decisiones por consumer (HR-67, #713)
"""Formato de cada decisión, sustitución enlazada, referencias que existen e
índice alineado con el directorio. Solo stdlib."""

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


GATE = _load("check_docs_decisions", ASSET / "check_docs_decisions.py")
CLI = _load("ci_pattern_cli_decisions", CLI_PATH,
            importlib.machinery.SourceFileLoader("ci_pattern_cli_decisions", str(CLI_PATH)))


def decision(estado="adoptada", extra="", fecha="2026-10-06", decide="el dueño", motivo="Porque sí."):
    return (f"# Una decisión\n\n- estado: {estado}\n- fecha: {fecha}\n- decide: {decide}\n"
            f"- motivo: {motivo}\n{extra}")


A, B = "2026-10-06-alfa", "2026-10-10-beta"
INDEX = f"# Decisiones\n\n| Decisión | Estado |\n|---|---|\n| [{A}]({A}.md) | adoptada |\n"
FILES = {"docs/decisions/README.md": INDEX, f"docs/decisions/{A}.md": decision()}


def findings(files=None, drop=()):
    with tempfile.TemporaryDirectory() as d:
        merged = {**FILES, **(files or {})}
        for rel, text in merged.items():
            if rel in drop:
                continue
            path = Path(d) / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        return GATE.check(Path(d), {})


def with_beta(beta, index_extra=f"| [{B}]({B}.md) | x |\n"):
    return {f"docs/decisions/{B}.md": beta, "docs/decisions/README.md": INDEX + index_extra}


class DocsDecisionsTests(unittest.TestCase):
    def test_conforme_da_verde(self):
        self.assertEqual(findings(), [])

    def test_falta_un_campo_nombra_fichero_y_campo(self):
        found = findings({f"docs/decisions/{A}.md": "# X\n\n- estado: adoptada\n"})
        self.assertTrue(any(A in f and "decide" in f for f in found), found)

    def test_estado_fuera_de_la_lista_es_rojo(self):
        found = findings({f"docs/decisions/{A}.md": decision("quizá")})
        self.assertTrue(any("estado" in f and "quizá" in f for f in found), found)

    def test_fecha_no_iso_y_motivo_vacio_son_rojos(self):
        self.assertTrue(any("fecha" in f for f in findings({f"docs/decisions/{A}.md": decision(fecha="ayer")})))
        self.assertTrue(any("motivo" in f for f in findings({f"docs/decisions/{A}.md": decision(motivo="")})))

    def test_sustituida_sin_enlace_es_roja(self):
        found = findings(with_beta(decision("sustituida")))
        self.assertTrue(any(B in f and "sustituida-por" in f for f in found), found)

    def test_sustituida_hacia_una_inexistente_es_roja(self):
        found = findings(with_beta(decision("sustituida", "- sustituida-por: fantasma\n")))
        self.assertTrue(any("fantasma" in f for f in found), found)

    def test_sustituida_bien_enlazada_es_verde(self):
        beta = decision("sustituida", f"- sustituida-por: {A}\n")
        self.assertEqual(findings(with_beta(beta)), [])

    def test_ciclo_de_sustitucion_y_autosustitucion_son_rojos(self):
        files = with_beta(decision("sustituida", f"- sustituida-por: {A}\n"))
        files[f"docs/decisions/{A}.md"] = decision("sustituida", f"- sustituida-por: {B}\n")
        self.assertTrue(any("ciclo" in f for f in findings(files)))
        selfref = {f"docs/decisions/{A}.md": decision("sustituida", f"- sustituida-por: {A}\n")}
        self.assertTrue(any("ciclo" in f for f in findings(selfref)))

    def test_aplazada_necesita_condicion_de_revision(self):
        self.assertTrue(any("revision" in f for f in findings(with_beta(decision("aplazada")))))
        ok = decision("aplazada", "- revision: tras la release 5\n")
        self.assertEqual(findings(with_beta(ok)), [])

    def test_una_cita_a_una_decision_inexistente_es_roja_y_a_una_existente_verde(self):
        found = findings({"docs/guia.md": f"Ver decisión: {A} y decisión: nada.\n"})
        self.assertTrue(any("docs/guia.md:1" in f and "nada" in f for f in found), found)
        self.assertEqual(findings({"docs/guia.md": f"Ver decisión: {A}.\n"}), [])

    def test_indice_y_directorio_no_pueden_divergir(self):
        extra = {f"docs/decisions/{B}.md": decision()}
        self.assertTrue(any(B in f and "índice" in f for f in findings(extra)))
        ghost = {"docs/decisions/README.md": INDEX + "| [fantasma](fantasma.md) | x |\n"}
        self.assertTrue(any("fantasma" in f for f in findings(ghost)))

    def test_directorio_sin_decisiones_falla_cerrado(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertTrue(any("ninguna decisión" in f for f in GATE.check(Path(d), {})))

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        buf = io.StringIO()
        with tempfile.TemporaryDirectory() as d, contextlib.redirect_stderr(buf):
            self.assertEqual(CLI._check_docs_decisions(Path(d), {}), [])
        self.assertIn("HR-67: docs_decisions no declarado", buf.getvalue())


if __name__ == "__main__":
    unittest.main()
