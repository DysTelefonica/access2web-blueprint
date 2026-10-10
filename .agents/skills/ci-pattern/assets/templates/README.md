# assets/templates — plantillas canónicas de issue y PR

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
