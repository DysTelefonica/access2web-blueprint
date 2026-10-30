# assets/templates — plantillas canónicas de issue, PR, docs y release

Las plantillas que el gate de contrato de PR lee **como dato** son las del propio
patrón, en su repositorio. La IA que adopta el patrón en un consumer las deriva
de ahí (fase 3 del porting-guide, G3.2), nunca de memoria:

- `.github/PULL_REQUEST_TEMPLATE.md` — cuerpo del PR: secciones **Chain Context**
  (HR-54) y «Tests que prueban el cierre» (HR-53, aviso), presupuesto de revisión
  y la excepción de tamaño como etiqueta `size:exception` (no como campo de datos).
- `.github/ISSUE_TEMPLATE/` — formularios de issue (la plantilla canónica de
  issue: sus seis secciones recomendadas; ningún gate las exige, team-skills#1093 R2).
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

Segunda excepción, la del commit del catálogo (#1062):

- `.github/workflows/operating-doc-drift.yml.tmpl` — el workflow de deriva del
  documento operativo. Se dispara en `schedule` y `workflow_dispatch` y no
  publica check requerido (HR-30). **No lleva el commit del catálogo escrito**:
  lo resuelve en runtime del sello `catalog @ <sha40>` de `.team-skills.yaml` y
  exporta a `$GITHUB_OUTPUT` para el `ref:` del checkout, y falla en voz alta
  nombrando el dato que falta. Cinco valores a cambiar, todos en su bloque
  `env`: repositorio del catálogo, credencial de lectura, manifiesto, ruta de la
  CLI dentro del checkout y fichero de parámetros.

La vigila `tests/test_operating_doc_drift_template.py`: diez casos, incluido el
negativo que mete un `ref:` literal en una copia y comprueba que el detector lo
ve, y la ejecución real del paso con `bash` sobre un consumer de mentira.

El mismo detector, aplicado al fragmento de `governance-seals` que publica
`references/porting-guide.md`, vive en `tests/test_governance_seals_job.py`
(#1062): el fragmento resuelve el sello de `.team-skills.yaml` en vez de llevar
`pin: <sha literal>`, y el paso final recibe su directorio por `env`.
