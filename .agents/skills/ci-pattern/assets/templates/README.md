# assets/templates — plantillas canónicas de issue, PR, docs y release

Las plantillas que el gate de contrato de PR lee **como dato** son las del propio
patrón, en su repositorio. La IA que adopta el patrón en un consumer las deriva
de ahí (fase 3 del porting-guide, G3.2), nunca de memoria:

- `.github/PULL_REQUEST_TEMPLATE.md` — cuerpo del PR: secciones **Chain Context**
  (HR-54) y «Tests que prueban el cierre» (HR-53), presupuesto de revisión y la
  excepción de tamaño como campo de datos.
- `.github/ISSUE_TEMPLATE/` — formularios de issue (el contrato canónico de
  issue: sus seis secciones obligatorias).
- `assets/pr-contract/check_pr_contract.py` — el gate que los lee: sus
  constantes nombran los campos exactos que tienen que aparecer.

**Estado:** este directorio queda reservado para una extracción verificada de
esos activos como plantillas parametrizadas; esa extracción no ha aterrizado y
no se inventa aquí. `ci-pattern adopt` y `ci-pattern update` están **retirados**
(exit 2): la adopción la gobierna `ci-pattern adoption check` sobre el contrato
del consumer.

Excepción ya aterrizada: el camino de release (#1015), que no existía como
activo del patrón y sí como plantillas:

- `.github/workflows/release.yml.tmpl` — el workflow de release del consumer.
  Se dispara con un tag estable; `preflight` corre los dos gates de
  `assets/release-gate/`, `release` publica en el entorno protegido `release`
  y `verify` comprueba la salida. Tres valores a cambiar, todos en su bloque
  `env`: la raíz donde el propagador materializó la skill, el basename del
  workflow de CI y la ruta del script de verificación.
- `scripts/verify-release-output.sh.tmpl` — la comprobación de salida: la
  revisión que sirve el endpoint de salud del consumer debe ser el SHA del tag.

Ambas pasan el gate `assets/workflow-policy/` del propio canon, que corre sobre
`tests/test_release_workflow_template.py`.
