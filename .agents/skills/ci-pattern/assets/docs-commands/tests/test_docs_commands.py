#!/usr/bin/env python3
# test_docs_commands.py — comandos documentados, extraídos y probados (HR-65, #711)
"""Los comandos que la documentación manda ejecutar se extraen (no se
enumeran), se clasifican y se prueban. Solo stdlib."""

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
JOB = ASSET / "run_docs_commands.py"
CONTRACT = ".github/ci-pattern-adoption.json"


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
SLOW = "import time\ntime.sleep(30)\n"
README = ("# Tool\n\nPara ver el estado: `tool status`.\n\n```sh\n"
          "tool init --name <proyecto>\ntool deploy | tee log.txt\n```\n")
CONFIG = {"cli_name": "tool", "cli_command": ["python3", "tool.py"],
          "executed_verbs": ["status"], "dry_run_flag": "--dry-run"}
# CLI que no existe en el runner: es el caso de APAP#2324 sobre `make` ausente.
AUSENTE = {"cli_name": "herramienta-ficticia", "cli_command": ["herramienta-ficticia"],
           "executed_verbs": ["--help"], "dry_run_flag": "--dry"}
DOS_COMANDOS = ("# Tool\n\n`herramienta-ficticia subcomando --dry`\n\n"
                "Y tambien: `herramienta-ficticia otro --help`.\n")


def _run(files=None, config=None):
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d)
        for rel, text in {"tool.py": TOOL, "README.md": README, **(files or {})}.items():
            (repo / rel).parent.mkdir(parents=True, exist_ok=True)
            (repo / rel).write_text(text, encoding="utf-8")
        return GATE.check(repo, {**CONFIG, **(config or {})})


def _timeout_de(seconds):
    """Baja el reloj del gate para no pagar TIMEOUT segundos en cada test."""
    original, GATE.TIMEOUT = GATE.TIMEOUT, seconds
    return original


def _counts(out):
    return {k: v.strip() for k, v in
            (pair.split("=") for pair in out.split(": ", 1)[1].strip().split(", "))}


def _job_repo(decl=None, files=None):
    """Repo temporal con el contrato de adopción y `decl` en `docs_commands`."""
    tmp = tempfile.TemporaryDirectory()
    repo = Path(tmp.name)
    for rel, text in {"tool.py": TOOL, "README.md": README, **(files or {})}.items():
        (repo / rel).parent.mkdir(parents=True, exist_ok=True)
        (repo / rel).write_text(text, encoding="utf-8")
    (repo / CONTRACT).parent.mkdir(parents=True, exist_ok=True)
    (repo / CONTRACT).write_text(json.dumps(
        {"docs_commands": dict(CONFIG, job="docs-commands", **(decl or {}))}),
        encoding="utf-8")
    return tmp, repo


def _job(repo, *args):
    proc = subprocess.run([sys.executable, str(JOB), str(repo), *args],
                          capture_output=True, text=True, timeout=60, check=False)
    return proc.returncode, proc.stdout, proc.stderr


class DocsCommandsTests(unittest.TestCase):
    def test_corpus_valido_es_verde_y_cuenta_los_tres_niveles(self):
        findings, summary = _run()
        self.assertEqual(findings, [])
        self.assertEqual(summary, {"executed": 1, "parse-only": 1, "presence-only": 1})

    def test_un_flag_inexistente_es_rojo_nombrando_fichero_y_linea(self):
        findings, _ = _run({"README.md": README + "\nOtro: `tool status --bogus`.\n"})
        self.assertTrue(any("README.md:" in f and "tool status --bogus" in f for f in findings), findings)

    def test_cli_ausente_es_un_hallazgo_de_entorno_y_no_uno_por_comando(self):
        findings, _ = _run({"README.md": DOS_COMANDOS, "tool.py": TOOL}, AUSENTE)
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("herramienta-ficticia", findings[0])

    def test_el_hallazgo_de_binario_ausente_dice_que_instalar_la_cli_en_el_runner(self):
        findings, _ = _run({"README.md": DOS_COMANDOS, "tool.py": TOOL}, AUSENTE)
        self.assertIn("entorno", findings[0])
        self.assertIn("runner", findings[0])
        self.assertIn("instale", findings[0].lower())

    def test_binario_ausente_no_cuenta_como_documentacion_mala(self):
        findings, _ = _run({"README.md": DOS_COMANDOS, "tool.py": TOOL}, AUSENTE)
        self.assertNotIn("no se pudo ejecutar", findings[0])
        self.assertNotIn("[Errno 2]", findings[0])

    def test_cli_del_repo_sin_permiso_de_ejecucion_no_dice_que_falta(self):
        with tempfile.TemporaryDirectory() as d:
            repo = Path(d)
            (repo / "cli-opaca").write_text("#!/bin/sh\nexit 0\n", encoding="utf-8")
            (repo / "cli-opaca").chmod(0o644)
            (repo / "README.md").write_text("`cli-opaca check`\n", encoding="utf-8")
            findings, _ = GATE.check(repo, {"cli_name": "cli-opaca",
                                            "cli_command": ["cli-opaca"],
                                            "executed_verbs": ["check"],
                                            "dry_run_flag": "--dry"})
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("no es ejecutable", findings[0])
        self.assertNotIn("no está en el `PATH`", findings[0])

    def test_timeout_es_un_hallazgo_distinto_del_binario_ausente(self):
        original = _timeout_de(1)
        try:
            timeout, _ = _run({"README.md": README, "slow.py": SLOW},
                              {"cli_command": ["python3", "slow.py"]})
            ausente, _ = _run({"README.md": DOS_COMANDOS, "tool.py": TOOL}, AUSENTE)
        finally:
            GATE.TIMEOUT = original
        self.assertTrue(any("timeout" in f for f in timeout), timeout)
        self.assertNotEqual(timeout, ausente)

    def test_el_hallazgo_de_timeout_dice_que_hacer(self):
        original = _timeout_de(1)
        try:
            findings, _ = _run({"README.md": README, "slow.py": SLOW},
                               {"cli_command": ["python3", "slow.py"]})
        finally:
            GATE.TIMEOUT = original
        self.assertTrue(any("no terminó" in f and "exclude" in f for f in findings), findings)

    def test_el_rechazo_del_parser_sigue_rojo_con_fichero_y_linea(self):
        findings, _ = _run({"README.md": README + "\nOtro: `tool status --bogus`.\n"})
        self.assertTrue(any(f.startswith("HR-65: README.md:") for f in findings), findings)

    def test_el_rechazo_del_parser_no_se_confunde_con_las_demas_causas(self):
        findings, _ = _run({"README.md": README + "\nOtro: `tool status --bogus`.\n"})
        (rejection,) = [f for f in findings if "rechazado por el parser" in f]
        for causa in ("entorno", "timeout", "no se pudo ejecutar"):
            self.assertNotIn(causa, rejection)

    def test_el_rechazo_del_parser_dice_que_hacer(self):
        findings, _ = _run({"README.md": README + "\nOtro: `tool status --bogus`.\n"})
        self.assertTrue(any("rechazado por el parser" in f and "corrija" in f
                            for f in findings), findings)

    def test_los_tres_hallazgos_dicen_que_hacer_y_son_distintos(self):
        original = _timeout_de(1)
        try:
            rechazo, _ = _run({"README.md": README + "\nOtro: `tool status --bogus`.\n"})
            timeout, _ = _run({"README.md": README, "slow.py": SLOW},
                              {"cli_command": ["python3", "slow.py"]})
        finally:
            GATE.TIMEOUT = original
        ausente, _ = _run({"README.md": DOS_COMANDOS, "tool.py": TOOL}, AUSENTE)
        # Cada causa dice qué hacer, y el remedio de una no vale para otra.
        self.assertEqual(len(ausente), 1, ausente)
        self.assertIn("instale", ausente[0].lower())
        self.assertIn("exclude", timeout[0].lower())
        self.assertIn("corrija", rechazo[0].lower())
        for hallazgo in (ausente[0], timeout[0], rechazo[0]):
            self.assertTrue(hallazgo.startswith("HR-65:"), hallazgo)
        self.assertEqual(len({ausente[0], timeout[0], rechazo[0]}), 3)

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

    def test_el_job_lee_el_contrato_de_adopcion_sin_fichero_en_la_raiz(self):
        tmp, repo = _job_repo()
        with tmp:
            code, out, _ = _job(repo)
        self.assertEqual(code, 0, out)
        self.assertEqual(_counts(out), {"executed": "1", "parse-only": "1", "presence-only": "1"})

    def test_el_contrato_manda_y_no_el_fichero_de_la_raiz(self):
        # El fichero de raíz declara otra CLI: si el job lo leyera, no extraería nada.
        tmp, repo = _job_repo(files={"docs-commands.json": json.dumps(
            {"cli_name": "otra-cli", "cli_command": ["otra-cli"], "dry_run_flag": "--help"})})
        with tmp:
            code, out, _ = _job(repo)
        self.assertEqual(code, 0, out)
        self.assertEqual(_counts(out), {"executed": "1", "parse-only": "1", "presence-only": "1"})

    def test_config_obsoleto_aun_funciona_y_avisa_por_stderr(self):
        tmp, repo = _job_repo(files={"viejo.json": json.dumps(CONFIG)})
        with tmp:
            code, out, err = _job(repo, "--config", "viejo.json")
        self.assertEqual(code, 0, out)
        self.assertEqual(_counts(out), {"executed": "1", "parse-only": "1", "presence-only": "1"})
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
            (repo / CONTRACT).write_text(json.dumps({"docs_commands": "make"}), encoding="utf-8")
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
        self.assertIn("docs_commands", err)

    def test_clave_desconocida_en_la_declaracion_es_hallazgo(self):
        with tempfile.TemporaryDirectory() as d:
            found = CLI._check_docs_commands(Path(d), {
                "docs_commands": dict(CONFIG, job="docs-commands", inventada="x")})
        self.assertTrue(any("clave desconocida" in f and "inventada" in f for f in found), found)

    def test_fichero_de_raiz_residual_se_nombra_como_segunda_fuente(self):
        tmp, repo = _job_repo(files={"docs-commands.json": json.dumps(CONFIG)})
        with tmp:
            found = CLI._check_docs_commands(repo, {"docs_commands": dict(CONFIG, job="docs-commands")})
        self.assertTrue(any("docs-commands.json" in f for f in found), found)

    def test_sin_fichero_de_raiz_no_hay_hallazgo_de_segunda_fuente(self):
        tmp, repo = _job_repo()
        with tmp:
            found = CLI._check_docs_commands(repo, {"docs_commands": dict(CONFIG, job="docs-commands")})
        self.assertFalse(any("segunda fuente" in f for f in found), found)


if __name__ == "__main__":
    unittest.main()
