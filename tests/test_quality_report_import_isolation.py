"""Los importadores de `quality_report` tienen que pasar ejecutados solos (BP#1353).

`scripts/quality_report.py` importa a sus hermanos por nombre de módulo desde
DQH 2.2, así que quien lo carga necesita `scripts/` en `sys.path`. Tres
ficheros de test lo cargan con `importlib`; solo uno se ocupaba de eso. Los
otros dos pasaban en la suite completa únicamente porque
`tests/test_quality_report.py::_load()` —que sí inserta `scripts/`— se
recogía antes y dejaba el intérprete ensuciado.

El síntoma era invisible para CI y carísimo de diagnosticar: el test no se
podía ejecutar solo, y cuando fallaba en la suite completa señalaba al
fichero equivocado, porque la causa estaba en el orden de recogida y no en
el test que fallaba.

Este fichero no reimplementa la propiedad: la mide desde fuera, lanzando cada
fichero en un intérprete nuevo, que es la única forma de observar "pasa
ejecutado solo" sin depender del propio proceso que lo está leyendo.

El `1 passed` del resumen se exige a propósito. Un `skip` también sale con
código 0, y un test saltado que evita un `ModuleNotFoundError` es verde sin
medir nada.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent

#: Los dos ficheros que cargan `quality_report` sin poner `scripts/` en `sys.path`.
#: El tercero, `tests/test_quality_report.py`, sí lo hace y no necesita el arreglo.
IMPORTING_FILES = (
    "tests/lanzadera/test_quality_report_smoke.py",
    "tests/test_ci_workflow.py",
)

#: El nodo concreto de cada fichero, para que un skip del resto no oculte el que importa.
PINNED_NODE = {
    "tests/lanzadera/test_quality_report_smoke.py": "test_gate_order_includes_legacy_hashes",
    "tests/test_ci_workflow.py": "test_gate_order_is_pinned_in_code",
}


def _run_alone(rel: str) -> subprocess.CompletedProcess[str]:
    """Ejecuta un fichero de test en un intérprete limpio, como haría un desarrollador."""
    return subprocess.run(
        [
            sys.executable,
            "-m",
            "pytest",
            "-c",
            "app/pyproject.toml",
            "--rootdir=app",
            "-p",
            "no:cacheprovider",
            "-q",
            f"{rel}::{PINNED_NODE[rel]}",
        ],
        capture_output=True,
        text=True,
        cwd=ROOT,
        timeout=300,
    )


@pytest.mark.parametrize("rel", IMPORTING_FILES)
def test_the_pinned_node_passes_in_a_clean_interpreter(rel: str) -> None:
    """El nodo que carga `quality_report` se sostiene sin nadie delante."""
    completed = _run_alone(rel)

    assert completed.returncode == 0, (
        f"`{rel}::{PINNED_NODE[rel]}` falla ejecutado solo, así que en la suite completa "
        "solo pasa porque otro test metió `scripts/` en `sys.path` antes. Un test que "
        "depende del orden de recogida no es una red: si cae, señala al fichero "
        "equivocado.\n"
        f"rc={completed.returncode}\n{completed.stdout[-1500:]}{completed.stderr[-800:]}"
    )


@pytest.mark.parametrize("rel", IMPORTING_FILES)
def test_the_pinned_node_runs_instead_of_skipping(rel: str) -> None:
    """Un `skip` sale con código 0: verde sin medir nada no es verde."""
    completed = _run_alone(rel)

    assert re.search(r"\b1 passed\b", completed.stdout), (
        f"`{rel}::{PINNED_NODE[rel]}` no llegó a ejecutarse: el resumen no dice `1 passed`. "
        "Si el arreglo fuera saltarse el import, el rojo se camufla de verde.\n"
        f"{completed.stdout[-800:]}"
    )


@pytest.mark.parametrize("rel", IMPORTING_FILES)
def test_running_the_file_leaves_no_state_behind(rel: str, tmp_path: Path) -> None:
    """Al terminar no puede quedar `scripts/` en `sys.path` ni el módulo cacheado.

    El síntoma original tenía dos mitades, y la segunda es la que envenena a
    los demás: dejar `quality_report` en `sys.modules` hace que el siguiente
    que lo importe reciba la copia en memoria en vez de la del disco, así que
    un cambio futuro en el fichero no se vería.

    La medición tiene que mirar el intérprete **después** de la corrida, y
    eso solo se puede desde dentro del proceso que corrió pytest. Por eso el
    wrapper de abajo llama a `pytest.main` y luego interroga su propio
    `sys.path` y `sys.modules`.
    """
    scripts = str(ROOT / "scripts")
    args = ["-q", "-p", "no:cacheprovider", f"{rel}::{PINNED_NODE[rel]}"]
    wrapper = tmp_path / "run_then_probe.py"
    wrapper.write_text(
        "import sys\n"
        "import pytest\n"
        f"CODE = pytest.main({args!r})\n"
        "LEAKED_MODULE = 'quality_report' in sys.modules\n"
        f"LEAKED_PATH = {scripts!r} in [p.rstrip('/') for p in sys.path]\n"
        "print('PROBE', LEAKED_MODULE, LEAKED_PATH)\n"
        "sys.exit(CODE)\n",
        encoding="utf-8",
    )
    completed = subprocess.run(
        [sys.executable, str(wrapper)], capture_output=True, text=True, cwd=ROOT, timeout=300
    )
    leaked = re.search(r"^PROBE (\w+) (\w+)$", completed.stdout, re.M)

    assert leaked is not None, (
        f"el wrapper no llegó a sondear: no hay linea `PROBE` en la salida.\n"
        f"stdout={completed.stdout[-800:]}\nstderr={completed.stderr[-800:]}"
    )
    module_leaked, path_leaked = leaked.group(1), leaked.group(2)

    assert module_leaked != "True", (
        "tras correr el fichero, `quality_report` sigue en `sys.modules`: el siguiente "
        "que lo importe recibe una copia en memoria y no verá los cambios del disco. "
        "Cargarlo con `monkeypatch.setitem` lo deshace (BP#1353)."
    )
    assert path_leaked != "True", (
        "tras correr el fichero, `scripts/` sigue en `sys.path`: un `sys.path.insert` "
        "a pelo se propaga a todo lo que corra después. `monkeypatch.syspath_prepend` "
        "lo deshace (BP#1353)."
    )
