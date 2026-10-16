#!/usr/bin/env bash
# test_docs_codeguide.sh — runner bash de la suite del gate docs-codeguide
# (issue #528, T10, A2). Descubre y ejecuta los test_*.py bajo
# ``tests/`` y propaga el primer exit != 0. Sin red, sin secretos: el
# gate es stdlib-only y la suite usa ``unittest`` + ``tempfile``.

set -u

THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

fail=0
for t in "$THIS_DIR"/test_*.py; do
    [[ -e "$t" ]] || continue
    if python3 "$t"; then
        echo "PASS $t"
    else
        echo "FAIL $t"
        fail=1
    fi
done

exit "$fail"