type: patch

HR-3 (#763): `testing/suites/vendorable-ruff` mide todos los assets vendorizables con el ruff extendido de los consumers (`S,ERA,ARG,FAST,N,C901,PLR,SIM,RET,TRY,PTH`) contra un baseline shrink-only por fichero y regla (`testing/vendorable-ruff-extended-baseline.json`): un hallazgo nuevo, un recuento que sube o una entrada obsoleta (`STALE`) dejan la suite en rojo; `--update-baseline` solo puede bajar.
