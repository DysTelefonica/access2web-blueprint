#!/usr/bin/env python3
# ci-pattern asset tests — la configuracion del gate sale de la base (#1270)
"""Un PR no puede reescribir la configuracion que lo mide.

En `pull_request_target` el workflow es el de la base, pero el checkout de la
cabeza del PR sigue trayendo los ficheros de configuracion del consumer: el
contrato de adopcion, el `ci-pattern.yaml`, el host-contract y la politica del
gate. Si el gate los lee de ahi, un PR puede debilitar su propia puerta.

La regla: la CONFIGURACION sale de la base (sin credenciales), y el ESTADO (el
manifiesto de gobernanza y los ficheros gobernados) sigue siendo el del PR.
"""
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent.parent
SKILL_ROOT = ASSET.parents[0]
TEMPLATE = ASSET / "templates" / ".github" / "workflows" / "pr-contract.yml.tmpl"
PORTING_GUIDE = SKILL_ROOT / "references" / "porting-guide.md"
CLI = SKILL_ROOT / "assets" / "bin" / "ci-pattern"

#: Ficheros de configuracion que gobiernan un gate. Ninguno puede leerse de la
#: cabeza del PR: si se lee de ahi, el PR escribe las reglas que lo miden.
CONFIG_FILES = (
    ".github/ci-pattern-adoption.json",
    ".github/pr-contract.policy.json",
    ".github/host-contract.json",
    "ci-pattern.yaml",
)


def guide_fragment() -> str:
    for block in re.findall(r"```yaml\n(.*?)```", PORTING_GUIDE.read_text(encoding="utf-8"), re.S):
        if "governance-seals:" in block:
            return block
    raise AssertionError("el porting-guide ya no publica el fragmento del job `governance-seals`")


class GateConfigurationFromBaseTests(unittest.TestCase):
    def test_no_configuration_is_read_from_the_pr_head(self):
        """Ninguna configuración del gate se lee de `pr-head`."""
        text = guide_fragment()
        for cfg in CONFIG_FILES:
            with self.subTest(configuracion=cfg):
                self.assertNotIn(f"pr-head/{cfg}", text)
                self.assertNotRegex(text, rf"pr-head/[^\s\"']*{re.escape(cfg.split('/')[-1])}")

    def test_the_fragment_reads_the_contract_from_the_base(self):
        text = guide_fragment()
        self.assertIn("--config-from base", text)
        self.assertIn("path: base", text)
        self.assertRegex(text, r"ref: \$\{\{ github\.event\.pull_request\.base\.sha \}\}")

    def test_the_base_checkout_leaves_no_credentials(self):
        text = guide_fragment()
        base = text[text.index("path: base"):]
        self.assertIn("persist-credentials: false", base[:400])

    def test_the_cli_takes_the_contract_from_the_config_directory(self):
        """`--config-from` carga el contrato de ahi: el repo sin contrato ya no es un exit 2."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp) / "repo"
            (repo / ".github").mkdir(parents=True)
            config = Path(tmp) / "config" / ".github"
            config.mkdir(parents=True)
            (config / "ci-pattern-adoption.json").write_text("{}", encoding="utf-8")
            without = subprocess.run([sys.executable, str(CLI), "adoption", "check",
                                      "--phase", "3", str(repo)], capture_output=True, text=True)
            with_config = subprocess.run([sys.executable, str(CLI), "adoption", "check",
                                          "--phase", "3", str(repo), "--config-from",
                                          str(Path(tmp) / "config")],
                                         capture_output=True, text=True)
        self.assertIn("contrato de adopción ausente", without.stdout + without.stderr)
        self.assertNotIn("contrato de adopción ausente", with_config.stdout + with_config.stderr)


if __name__ == "__main__":
    unittest.main()
