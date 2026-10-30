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
# APAP#2324: `{"cli_name": "make", "cli_command": ["make"]}` sobre un runner sin
# `make`. El nombre es ficticio porque `make` sí está en el PATH de este runner y
# el caso se repetiría como «la CLI está» (#1038). Sin `dry_run_flag` ni
# `executed_verbs` con valor, `_tier` devuelve presence-only para todo comando.
SIN_EJECUCION = {"cli_name": "herramienta-ausente", "cli_command": ["herramienta-ausente"],
                 "executed_verbs": [], "dry_run_flag": ""}
DOC_SIN_EJECUCION = ("# Tool\n\nPara compilar: `herramienta-ausente build`.\n\n```sh\n"
                     "herramienta-ausente test | tee salida.txt\n```\n")
# --- Capa de existencia ---------------------------------------------------
#
# El caso vivo (DysTelefonica/access2web-blueprint#1374): la documentacion cita
# `make test` con `--cov` y da el `Makefile` como fuente, pero el target `test`
# dejo de medir cobertura. El gate validaba el comando en su forma (`pytest
# --help` sale 0 sin mirar el resto) y no miraba la existencia.
MAKEFILE = ("PYTEST ?= $(PYTHON) -m pytest\n\n.PHONY: test\n\n"
            "test:\n\t$(PYTEST) -c app/pyproject.toml --rootdir=app -m \"not integration\"{extra}\n")
SIN_COBERTURA = MAKEFILE.format(extra="")
CON_COBERTURA = MAKEFILE.format(extra=" --cov --cov-fail-under=69")
FAKE_PYTEST = "import sys\nprint('uso: pytest [opciones]')\n"
BP_CONFIG = {"cli_name": "pytest", "cli_command": ["python3", "fake_pytest.py"],
             "executed_verbs": [], "dry_run_flag": "--help",
             "existence": {"makefile": "Makefile"}}
# Documentos con la CLI (para que el gate tenga algo que extraer) y la cita que
# se mide, en la misma linea.
DOC_BP = ("| Medición local | `make test` (`pytest -c app/pyproject.toml --rootdir=app "
          "-m \"not integration\" --cov --cov-fail-under=69`) | `Makefile` |\n")
DOC_MAKE = "Mida con `pytest app/tests` y luego `make {target}`.\n"
DOC_SCRIPT = "Prepare el repo con `pytest app/tests` y luego `./scripts/{script}`.\n"
DOC_PATH = "Prepare el repo con `pytest app/tests` y luego `{tool}`.\n"
DOC_OTRO = "Mida con `pytest app/tests`, `make test` y revise `ruff check --fix`.\n"
# Una cita con ruta, en un documento que NO esta en la raiz del repo (#1385).
DOC_CITA = "Mida con `pytest app/tests` y lea `{cita}`.\n"




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

    def test_presence_only_por_construccion_con_cli_ausente_avisa_y_no_roja(self):
        # APAP#2324: sin `dry_run_flag` ni `executed_verbs` declarados, `_tier`
        # devuelve presence-only para todo comando, así que el gate no ejecuta
        # nada y una CLI ausente no puede estar ocultando ninguna medición: el
        # hallazgo se mantiene como aviso y el job sale 0 (#1038).
        tmp, repo = _job_repo(SIN_EJECUCION, {"README.md": DOC_SIN_EJECUCION})
        with tmp:
            code, out, err = _job(repo)
        self.assertEqual(code, 0, out + err)
        self.assertEqual(_counts(out), {"executed": "0", "parse-only": "0",
                                        "presence-only": "2"})
        self.assertIn("entorno", err)
        self.assertIn("instale", err.lower())
        self.assertEqual(len(err.strip().splitlines()), 1, err)
        self.assertNotIn("entorno", out)

    def test_el_aviso_se_distingue_del_hallazgo_por_canal_y_por_forma(self):
        tmp, repo = _job_repo(SIN_EJECUCION, {"README.md": DOC_SIN_EJECUCION})
        with tmp:
            _, _, err_aviso = _job(repo)
        tmp, repo = _job_repo({**SIN_EJECUCION, "dry_run_flag": "--dry"},
                              {"README.md": DOC_SIN_EJECUCION})
        with tmp:
            _, out_hallazgo, err_hallazgo = _job(repo)
        self.assertTrue(err_aviso.startswith("AVISO"), err_aviso)
        self.assertEqual(err_hallazgo, "", err_hallazgo)
        (hallazgo,) = [línea for línea in out_hallazgo.splitlines() if línea.startswith("HR-65:")]
        self.assertIn("entorno", hallazgo)

    def test_con_dry_run_declarado_la_cli_ausente_sigue_siendo_roja(self):
        tmp, repo = _job_repo({**SIN_EJECUCION, "dry_run_flag": "--dry"},
                              {"README.md": DOC_SIN_EJECUCION})
        with tmp:
            code, out, err = _job(repo)
        self.assertEqual(code, 1, out + err)
        self.assertEqual(_counts(out.splitlines()[0]), {"executed": "0", "parse-only": "1",
                                                        "presence-only": "1"})
        (hallazgo,) = [línea for línea in out.splitlines() if línea.startswith("HR-65:")]
        self.assertIn("entorno", hallazgo)

    def test_con_executed_verbs_declarados_la_cli_ausente_sigue_siendo_roja(self):
        # `build` es un verbo que sí aparece en la documentación del fixture, así
        # que el gate tenía una medición que perder.
        tmp, repo = _job_repo({**SIN_EJECUCION, "executed_verbs": ["build"]},
                              {"README.md": DOC_SIN_EJECUCION})
        with tmp:
            code, out, err = _job(repo)
        self.assertEqual(code, 1, out + err)
        self.assertEqual(_counts(out.splitlines()[0]), {"executed": "1", "parse-only": "0",
                                                        "presence-only": "1"})
        (hallazgo,) = [línea for línea in out.splitlines() if línea.startswith("HR-65:")]
        self.assertIn("entorno", hallazgo)

    def test_una_cli_que_arranca_y_revienta_sigue_siendo_roja(self):
        # `which` la da por lanzable y revienta al arrancar (shebang roto): es el
        # camino de `CliAusenteError`, que #1038 no relaja. Con `executed_verbs`
        # declarado el gate iba a ejecutar, así que la ausencia sí oculta algo.
        tmp, repo = _job_repo({**SIN_EJECUCION, "executed_verbs": ["build"]},
                              {"README.md": DOC_SIN_EJECUCION})
        cli = repo / SIN_EJECUCION["cli_name"]
        cli.write_text("#!/interprete-que-no-existe\n", encoding="utf-8")
        cli.chmod(0o755)
        with tmp:
            code, out, err = _job(repo)
        self.assertEqual(code, 1, out + err)
        self.assertTrue(any("no arranca" in línea for línea in out.splitlines()), out)

    def test_el_resumen_es_identico_con_y_sin_cli_en_presence_only(self):
        con_cli, con = _run({"README.md": DOC_SIN_EJECUCION},
                            {**SIN_EJECUCION, "cli_command": ["python3", "tool.py"]})
        sin_cli, sin = _run({"README.md": DOC_SIN_EJECUCION}, SIN_EJECUCION)
        self.assertEqual(con, {"executed": 0, "parse-only": 0, "presence-only": 2})
        self.assertEqual(con, sin)
        self.assertEqual((con_cli, sin_cli), ([], []))

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
        self.assertEqual(_counts(out), {"executed": "1", "parse-only": "1",
                                        "presence-only": "1"})

    def test_el_contrato_manda_y_no_el_fichero_de_la_raiz(self):
        # El fichero de raíz declara otra CLI: si el job lo leyera, no extraería nada.
        tmp, repo = _job_repo(files={"docs-commands.json": json.dumps(
            {"cli_name": "otra-cli", "cli_command": ["otra-cli"], "dry_run_flag": "--help"})})
        with tmp:
            code, out, _ = _job(repo)
        self.assertEqual(code, 0, out)
        self.assertEqual(_counts(out), {"executed": "1", "parse-only": "1",
                                        "presence-only": "1"})

    def test_config_obsoleto_aun_funciona_y_avisa_por_stderr(self):
        tmp, repo = _job_repo(files={"viejo.json": json.dumps(CONFIG)})
        with tmp:
            code, out, err = _job(repo, "--config", "viejo.json")
        self.assertEqual(code, 0, out)
        self.assertEqual(_counts(out), {"executed": "1", "parse-only": "1",
                                        "presence-only": "1"})
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


class ExistenceTests(unittest.TestCase):
    """La capa de existencia: lo que el comando documentado invoca tiene que existir.

    Los tres niveles de ejecucion prueban la FORMA del comando; esta capa prueba la
    EXISTENCIA de lo que invoca. Es la que caza el caso vivo de BP: la documentacion
    cita `make test` con `--cov` y el target ya no lo mide (#1374)."""

    def _bp(self, doc=DOC_BP, makefile=SIN_COBERTURA, files=None, config=None):
        return _run({"docs/adopcion-fase-2-aislamiento.md": doc, "Makefile": makefile,
                     "fake_pytest.py": FAKE_PYTEST, **(files or {})},
                    {**BP_CONFIG, **(config or {})})

    def test_el_flag_que_el_target_deja_de_ejecutar_es_rojo(self):
        # Caso vivo de BP: la documentacion dice que `make test` mide con `--cov` y
        # el target dejo de medirlo. La forma la aprueba el parser (`pytest --help`
        # sale 0 sin mirar el resto); la existencia, no.
        findings, _ = self._bp()
        self.assertTrue(any("adopcion-fase-2-aislamiento.md:" in f and "--cov" in f for f in findings), findings)

    def test_el_hallazgo_nombra_el_flag_que_falta_y_dice_que_hacer(self):
        findings, _ = self._bp()
        (hallazgo,) = [f for f in findings if "--cov" in f]
        self.assertIn("--cov-fail-under=69", hallazgo)
        self.assertIn("no hace lo que la documentacion dice", hallazgo)
        self.assertIn("corrija", hallazgo)
        self.assertIn("Makefile", hallazgo)

    def test_el_flag_que_el_target_si_ejecuta_no_es_hallazgo(self):
        findings, _ = self._bp(makefile=CON_COBERTURA)
        self.assertEqual(findings, [])

    def test_un_target_de_make_inexistente_es_rojo(self):
        findings, _ = self._bp(doc=DOC_MAKE.format(target="no-existe"))
        self.assertTrue(any("no esta en el Makefile" in f and "no-existe" in f for f in findings), findings)

    def test_un_target_de_make_que_existe_no_es_hallazgo(self):
        findings, _ = self._bp(doc=DOC_MAKE.format(target="test"))
        self.assertEqual(findings, [])

    def test_un_script_local_inexistente_es_rojo(self):
        findings, _ = self._bp(doc=DOC_SCRIPT.format(script="no-existe.sh"))
        self.assertTrue(any("no tiene el script" in f and "scripts/no-existe.sh" in f for f in findings), findings)

    def test_un_script_local_que_existe_no_es_hallazgo(self):
        # La cita se resuelve contra el directorio del documento que la escribe
        # (#1385): este documento vive en `docs/`, asi que `./scripts/x.sh` es
        # `docs/scripts/x.sh`.
        findings, _ = self._bp(doc=DOC_SCRIPT.format(script="preflight.sh"),
                               files={"docs/scripts/preflight.sh": "#!/bin/sh\n"})
        self.assertEqual(findings, [])

    def test_un_documento_en_la_raiz_sigue_resolviendo_desde_la_raiz(self):
        # El documento esta en la raiz: su directorio es la raiz del repo, y la
        # cita se resuelve igual que antes (#1385).
        findings, _ = _run({"README.md": README + "\nPrepare con `./scripts/preflight.sh`.\n",
                            "scripts/preflight.sh": "#!/bin/sh\n"},
                           {**CONFIG, "existence": {"makefile": "Makefile"}})
        self.assertFalse(any("no tiene el script" in f for f in findings), findings)

    def test_un_ejecutable_del_path_que_no_resuelve_es_rojo(self):
        findings, _ = self._bp(doc=DOC_PATH.format(tool="herramienta-ficticia-xyz"),
                               config={"existence": {"path_executables":
                                                     ["herramienta-ficticia-xyz"]}})
        self.assertTrue(any("no se resuelve" in f and "herramienta-ficticia-xyz" in f for f in findings), findings)

    def test_un_ejecutable_del_path_que_resuelve_no_es_hallazgo(self):
        findings, _ = self._bp(doc=DOC_PATH.format(tool="python3 tool.py"),
                               config={"existence": {"path_executables": ["python3"]}})
        self.assertEqual(findings, [])

    def test_sin_declarar_existence_la_capa_no_mide_nada(self):
        # La capa la activa la declaracion: un consumer que no la declara no cambia.
        findings, _ = self._bp(doc=DOC_SCRIPT.format(script="no-existe.sh"),
                               config={"existence": None})
        self.assertFalse(any("no tiene el script" in f for f in findings), findings)

    def test_la_capa_de_existencia_no_depende_de_la_cli_en_el_runner(self):
        # La capa no ejecuta nada: mide lo que el comando invoca aunque la CLI falte.
        findings, _ = self._bp(doc=DOC_SCRIPT.format(script="no-existe.sh"),
                               config={"cli_command": ["python3", "no-existe.py"]})
        self.assertTrue(any("no tiene el script" in f for f in findings), findings)

    def test_el_exclude_exime_un_comando_de_existencia(self):
        findings, _ = self._bp(doc=DOC_SCRIPT.format(script="no-existe.sh"),
                               config={"exclude": ["./scripts/no-existe.sh"]})
        self.assertFalse(any("no tiene el script" in f for f in findings), findings)
        self.assertFalse(any("STALE" in f for f in findings), findings)

    # --- La cita se resuelve contra el documento que la cita (#1385) ----------
    #
    # El caso vivo de BP: `docs/03-aplicaciones/brass/prd/capacidades.md` cita
    # `../capacidades.md` y el fichero esta en el directorio de ese documento,
    # no en la raiz del repo. La capa lo buscaba desde la raiz y daba 234 rojos
    # falsos. Los enlaces markdown ya se resuelven contra el directorio de su
    # fichero; la capa de existencia se iguala a ellos.

    def _cita(self, cita, doc="docs/prd/guia.md", files=None, config=None):
        return _run({doc: DOC_CITA.format(cita=cita),
                     "fake_pytest.py": FAKE_PYTEST, **(files or {})},
                    {**BP_CONFIG, **(config or {})})

    def test_el_script_se_busca_junto_al_documento_que_lo_cita(self):
        findings, _ = self._cita("../capacidades.md",
                                  files={"docs/capacidades.md": "# capacidades\n"})
        self.assertEqual(findings, [], findings)

    def test_una_cita_que_sale_del_repositorio_es_rojo(self):
        findings, _ = self._cita("../../../fuera.sh")
        self.assertTrue(any("fuera del repositorio" in f for f in findings), findings)

    def test_una_ruta_sin_punto_guion_sigue_resolviendose_en_el_path(self):
        # La capa solo mide scripts con `./` o `../`. Una ruta desnuda se
        # resuelve en el PATH, como antes: el defecto no la cambia.
        findings, _ = self._cita("guia.sh", files={"docs/prd/guia.sh": "#!/bin/sh\n"},
                                  config={"existence": {"makefile": "Makefile",
                                                       "path_executables": ["guia.sh"]}})
        self.assertTrue(any("no se resuelve" in f and "guia.sh" in f for f in findings), findings)
        self.assertFalse(any("no tiene el script" in f for f in findings), findings)

    def test_una_cita_que_es_un_directorio_no_es_un_script(self):
        findings, _ = self._cita("../prd", files={"docs/prd/guia.sh": "#!/bin/sh\n"})
        self.assertTrue(any("no es un script" in f for f in findings), findings)

    # --- Una cita vale si existe contra el documento O contra la raiz (#1385) ---
    #
    # Un comando entre comillas invertidas se ejecuta desde la raiz del repo, no
    # desde la carpeta del documento que lo muestra: `docs/guia.md` que dice
    # «ejecute `./scripts/check.sh`» apunta a `scripts/check.sh`. La regla nueva
    # acepta la cita si el fichero existe resuelto contra el documento O contra
    # la raiz; si sale del repo por las dos vias, sigue en rojo.

    def test_una_cita_de_la_raiz_desde_un_documento_anidado_es_valida(self):
        findings, _ = self._cita("./scripts/check.sh", doc="docs/guia.md",
                                  files={"scripts/check.sh": "#!/bin/sh\n"})
        self.assertEqual(findings, [], findings)

    def test_una_cita_que_no_existe_ni_en_el_documento_ni_en_la_raiz_es_rojo(self):
        findings, _ = self._cita("./scripts/check.sh", doc="docs/guia.md")
        self.assertTrue(any("no tiene el script" in f and "docs/scripts/check.sh" in f
                            and "scripts/check.sh" in f for f in findings), findings)

    def test_una_cita_que_sale_del_repo_por_las_dos_vias_sigue_en_rojo(self):
        findings, _ = self._cita("../../../fuera.sh", doc="docs/prd/guia.md")
        self.assertTrue(any("fuera del repositorio" in f for f in findings), findings)

    def test_clave_desconocida_en_existence_es_hallazgo(self):
        findings, _ = self._bp(config={"existence": {"inventada": "x"}})
        self.assertTrue(any("clave desconocida" in f and "inventada" in f for f in findings), findings)

    def test_la_existencia_no_se_cruza_con_un_comando_que_no_es_la_cli(self):
        # La atribucion de flags es de la misma linea y solo con la CLI declarada.
        findings, _ = self._bp(doc=DOC_OTRO)
        self.assertEqual(findings, [])


if __name__ == "__main__":
    unittest.main()
