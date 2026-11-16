type: patch

**Qué.** Las tres suites que aún ponían un `gh` falso sin extensión primero en el PATH (`test_release_workflow_template.py`, `test_sync_governance_labels.py` y `test_gh_api_retry.sh`) ya no pueden llegar al `gh` real: la guarda de `_fake_gh.py` o una guarda equivalente en bash se niega a ejecutar antes de lanzar ningún proceso.
**Causa.** En Windows, un `gh` sin extensión no lo encuentra el PATH y gana `gh.exe`, el real. El #1609 cubrió las suites que usan `CI_PATTERN_GH`, pero estas tres quedaron fuera.
**Arreglo.** Las suites de Python llaman a `guard_fake_gh` antes de cada ejecución que llama a `gh`. El asset de bash (`gh-api-retry.sh`) llama a `gh` por nombre y no lee `CI_PATTERN_GH`, así que su suite resuelve el mismo PATH que recibe el asset con una guarda propia. Un test por suite demuestra que, con un `gh` real primero en el PATH, la suite se niega y ese `gh` no recibe ninguna llamada.
**A partir de ahora.** Los tests del asset no dependen del PATH para elegir el `gh`. Un consumer que vendoriza estas suites debe re-vendorizarlas.
