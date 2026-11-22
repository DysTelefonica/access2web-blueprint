type: patch

El ratchet de ruff extendido (#1750) deja de marcar `gate-liveness/check_gate_liveness.py` con SIM201: el patrón `not ... in` que lo causaba desaparece al identificar solo los jobs de primer nivel. La entrada obsoleta se retira del baseline (`testing/vendorable-ruff-extended-baseline.json`); el baseline no crece.
