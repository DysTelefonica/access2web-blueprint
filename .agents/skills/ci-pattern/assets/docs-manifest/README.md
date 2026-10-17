# docs-manifest — lista canónica de documentos y de sus secciones

Asset de `ci-pattern` (HR-70, DysTelefonica/team-skills#719). `required-docs.json`
es la fuente única de qué documentos debe tener un consumer, para qué sirve cada
uno, qué secciones lleva y qué HR lo exige. `check_docs_manifest.py` (stdlib, sin
red) comprueba que cada documento activado existe y que cada sección obligatoria
existe como título y no está vacía; el rojo nombra fichero y sección.

| Campo | Significado |
|---|---|
| `path` | Ruta del documento (admite `*`, p. ej. `docs/codebase/*.md`) |
| `purpose` | Para qué sirve |
| `sections` | Títulos obligatorios (la numeración `N.` del título no cuenta) |
| `required_by` | HR que lo exige |
| `route` | Opcional `{when, read}`: fila del enrutado «Dónde trabajar» de `AGENTS.md` (HR-69) |
| `activated_by` | Clave del contrato que lo activa (`docs_map`…); sin ella, siempre |

Activación: `docs_manifest` en el contrato de adopción. Sin declaración, `adoption
check --phase 5` avisa por `stderr` y sale 0. El consumer puede **añadir**
documentos y secciones con `docs_manifest.extend`; cualquier otra clave (p. ej.
`exclude`) es rojo: lo del canon no se quita. Cada HR que exija un documento nuevo
lo registra aquí, y el enrutado de `AGENTS.md` se genera de este mismo fichero.
