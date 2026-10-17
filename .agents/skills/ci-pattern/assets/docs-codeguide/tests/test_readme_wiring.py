#!/usr/bin/env python3
# test_readme_wiring.py — el cableado que el README manda copiar (HR-63, #709)
"""El README es la plantilla de cableado del job: define el job
``docs-codeguide`` y la lista ``required_jobs`` que lo exige. Stdlib."""

import unittest
from pathlib import Path

README = (Path(__file__).resolve().parent.parent / "README.md").read_text(encoding="utf-8")


class ReadmeWiringTests(unittest.TestCase):
    def test_la_plantilla_define_el_job_docs_codeguide(self):
        self.assertRegex(README, r"(?m)^docs-codeguide:\n\s+runs-on:")

    def test_la_plantilla_lo_mete_en_required_jobs_y_en_el_needs_del_agregador(self):
        self.assertRegex(README, r'"required_jobs": \[[^\]]*"docs-codeguide"')
        self.assertRegex(README, r"needs: \[[^\]]*docs-codeguide")


if __name__ == "__main__":
    unittest.main()
