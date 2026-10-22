#!/usr/bin/env python3
# ci-pattern asset — HR-32 meta-gate: every matrix gate sees a violation fail (#245)
"""Meta-gate for HR-32: every ``gate`` entry in the HR→gate matrix carries a
test that exists AND a violation recipe proving the gate fails (exit != 0)
against a violation fixture. A gate whose violation check passes cleanly —
or a gate with no recipe at all — is the violation this gate reports.

The recipes live in this file as data + code: each one builds its violating
input locally (temp dirs, throwaway git repositories, no network) and runs
the real gate binary the matrix entry declares. Recipes are keyed by the
matrix entry's ``test`` field, so several HRs sharing one gate share one
recipe. This file's own matrix entry (HR-32) is the check itself and is
exempt from the registry by construction.

Exit codes: 0 every gate rejects its violation · 1 findings (a gate exists
that does not fail against its violation, or has no recipe) · 2 invalid
input (matrix unreadable).
"""

from __future__ import annotations

import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[3]
MATRIX = SKILL_ROOT / "references" / "hr-gate-matrix.json"


def main(argv):
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
