#!/usr/bin/env python3
# test_docs_commands.py — comandos documentados, extraídos y probados (HR-65, #711)
"""Los comandos que la documentación manda ejecutar se extraen (no se
enumeran), se clasifican y se prueban. Solo stdlib."""

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


GATE = _load("check_docs_commands", ASSET / "check_docs_commands.py")
CLI = _load("ci_pattern_cli_commands", CLI_PATH,
            importlib.machinery.SourceFileLoader("ci_pattern_cli_commands", str(CLI_PATH)))
TOOL = '''import sys
args = sys.argv[1:]
if args[:1] == ["status"] and len(args) == 1:
    print("ok")
elif args[:1] == ["init"] and "--name" in args and "--dry-run" in args:
    print("init")
else:
    print("uso invalido", file=sys.stderr)
    sys.exit(2)
'''
README = ("# Tool\n\nPara ver el estado: `tool status`.\n\n```sh\n"
          "tool init --name <proyecto>\ntool deploy | tee log.txt\n```\n")
CONFIG = {"cli_name": "tool", "cli_command": ["python3", "tool.py"],
          "executed_verbs": ["status"], "dry_run_flag": "--dry-run"}


def _run(files=None, config=None):
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d)
        for rel, text in {"tool.py": TOOL, "README.md": README, **(files or {})}.items():
            (repo / rel).parent.mkdir(parents=True, exist_ok=True)
            (repo / rel).write_text(text, encoding="utf-8")
        return GATE.check(repo, {**CONFIG, **(config or {})})


class DocsCommandsTests(unittest.TestCase):
    def test_corpus_valido_es_verde_y_cuenta_los_tres_niveles(self):
        findings, summary = _run()
        self.assertEqual(findings, [])
        self.assertEqual(summary, {"executed": 1, "parse-only": 1, "presence-only": 1})

    def test_un_flag_inexistente_es_rojo_nombrando_fichero_y_linea(self):
        findings, _ = _run({"README.md": README + "\nOtro: `tool status --bogus`.\n"})
        self.assertTrue(any("README.md:" in f and "tool status --bogus" in f for f in findings), findings)

    def test_anadir_un_comando_a_un_doc_lo_anade_al_corpus_sin_tocar_el_test(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "docs").mkdir()
            (Path(d) / "docs" / "a.md").write_text("`tool status` y `tool nuevo`\n", encoding="utf-8")
            found = GATE.extract(Path(d), {"cli_name": "tool", "roots": ["docs"]})
        self.assertEqual([c["command"] for c in found], ["tool status", "tool nuevo"])

    def test_un_bloque_anidado_y_uno_de_tildes_extraen_sus_comandos(self):
        text = "````md\n```sh\ntool uno\n```\ntool dos\n````\n\n~~~\ntool tres\n~~~\n"
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "README.md").write_text(text, encoding="utf-8")
            found = GATE.extract(Path(d), {"cli_name": "tool", "roots": ["README.md"]})
        self.assertEqual([c["command"] for c in found], ["tool uno", "tool dos", "tool tres"])

    def test_extraer_cero_comandos_falla_cerrado(self):
        findings, _ = _run({"README.md": "# sin comandos\n"})
        self.assertTrue(any("ningún comando" in f for f in findings), findings)

    def test_el_nombre_de_la_cli_es_parametro_y_el_asset_no_nombra_ninguna_herramienta(self):
        findings, summary = _run({"README.md": "`otra status`\n"},
                                 {"cli_name": "otra", "cli_command": ["python3", "tool.py"]})
        self.assertEqual((findings, summary["executed"]), ([], 1))
        self.assertNotIn("gentle-ai", (ASSET / "check_docs_commands.py").read_text(encoding="utf-8"))

    def test_placeholder_en_parse_only_se_rellena_con_un_valor_ficticio(self):
        self.assertEqual(_run()[0], [])

    def test_sin_dry_run_declarado_los_comandos_con_placeholder_son_solo_presencia(self):
        _findings, summary = _run(config={"dry_run_flag": ""})
        self.assertEqual(summary, {"executed": 1, "parse-only": 0, "presence-only": 2})

    def test_exclusion_obsoleta_es_stale(self):
        findings, _ = _run(config={"exclude": ["tool ya-no-existe"]})
        self.assertTrue(any("STALE" in f and "tool ya-no-existe" in f for f in findings), findings)

    def test_sin_declaracion_avisa_por_stderr_y_sale_verde(self):
        buf = io.StringIO()
        with tempfile.TemporaryDirectory() as d, contextlib.redirect_stderr(buf):
            self.assertEqual(CLI._check_docs_commands(Path(d), {}), [])
        self.assertIn("HR-65: docs_commands no declarado", buf.getvalue())

    def test_declaracion_incompleta_falla_nombrando_la_clave(self):
        with tempfile.TemporaryDirectory() as d:
            found = CLI._check_docs_commands(Path(d), {"docs_commands": {"cli_name": "tool"}})
        self.assertTrue(any("cli_command" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
