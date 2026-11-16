#!/usr/bin/env python3
# ci-pattern asset tests — ci-red: el doble de gh (#1508, HR-89)
"""Guarda: ``gh`` resuelve al doble de ``_gh_stub.py``, nunca al ``gh`` real."""
import shutil
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _gh_stub import HostTestCase  # noqa: E402


class GhStubGuardTests(HostTestCase):
    def test_gh_resolves_to_the_stub_never_to_the_real_gh(self):
        host = self.host()
        resolved = shutil.which("gh", path=host.env()["PATH"])
        self.assertIsNotNone(resolved, "gh no resuelve en el PATH del test")
        self.assertEqual(Path(resolved).resolve().parent, host.bin.resolve(),
                         f"gh resuelve a {resolved}, no al doble: el test llamaría al gh real")


if __name__ == "__main__":
    unittest.main()
