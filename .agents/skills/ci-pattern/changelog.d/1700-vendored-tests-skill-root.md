type: patch

Los tests vendorizados de `assets/prerelease/tests` dejan de asumir el árbol del catálogo (#1700). `test_no_cost_pr.py` busca `references/porting-guide.md` relativo a la raíz de la skill, no en una ruta fija `personal/ardelperal/ci-pattern/`, así que pasa también en un consumer que vendoriza la skill en `.agents/skills/ci-pattern/`. `test_provision_night_runner.py` compara los pins con el `tests.yml` del catálogo solo si existe: en un consumer sin ese fichero se salta con un motivo explícito, y en el catálogo su ausencia sigue siendo un fallo. Refs #1700.
