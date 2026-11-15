---
name: release-notes
description: "Trigger: notas de versión, notas de la release, redactar la release, texto de la release, registro de cierres, cobertura de las notas. Redacta las notas de una versión a partir del rango entre dos tags, agrupadas por interés del lector, con cada entrada en antes, ahora y dónde, y comprueba que ninguna PR, commit o issue del rango queda sin contar."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.0"
  last_verified: 2026-10-29
  scope: ['universal', 'docs']
  auto_invoke: ['drafting release notes', 'checking release notes coverage']
  tiers: ['universal', 'docs']
---

# release-notes

Notas de versión que cualquiera puede leer y saber qué hay de nuevo: cada cosa integrada, la de seguridad, la de producto, los errores y la funcionalidad, sin dejar nada del rango sin contar. El modelo es la release v1.2.1 de `DysTelefonica/cadete` (`docs/releases/v1.2.1.md`). Vale para repositorios web/contenedor y para proyectos VBA/Access; cada perfil define sus fuentes de datos en `references/`.

## §1 Activation

Cargue esta skill cuando:

- Vaya a redactar las notas de una versión, de un hotfix o de un rc que pasa a producción.
- Revise unas notas ya escritas antes de publicar la GitHub Release.
- Necesite comprobar que unas notas cuentan todo el rango entre dos tags.

**Do NOT load when:**

- Decida qué versión sale o cuándo: eso lo decide el mantenedor.
- Publique la release o mueva el tag: lo hace el workflow o el procedimiento de release del repositorio.
- Escriba un fragmento de `changelog.d/` de una skill del catálogo: eso es el PR de release del catálogo.

## §2 Hard Rules

- **HR-1** — Las notas MUST cubrir el rango `<tag anterior>..<tag nuevo>` entero. Antes de entregarlas, ejecute `assets/release_coverage.py check`; un exit distinto de 0 MUST bloquear la entrega. Ninguna PR, commit sin PR o issue cerrada del rango puede faltar.
- **HR-2** — Las entradas MUST agruparse por interés del lector, en este orden: Seguridad, Errores de producto corregidos, Funcionalidad nueva o ampliada, Repositorio, pruebas y CI. El último grupo MUST ir separado y abrir con «No cambian la aplicación.». Un grupo sin entradas se omite.
- **HR-3** — Cada entrada de los tres primeros grupos MUST llevar cuatro datos: qué ocurría en la versión anterior, qué ocurre en la nueva, dónde (módulo, formulario, pantalla o ruta) y la PR y la issue que lo cierran. En funcionalidad nueva, «En <anterior>» puede decir «no existía». Una entrada sin «dónde» o sin referencia es reject.
- **HR-4** — El Resumen MUST abrir las notas en pocas líneas, una por grupo, y MUST avisar de migraciones, cambios de configuración y cambios incompatibles. Quien lea solo el Resumen MUST saber qué cambia.
- **HR-5** — La Procedencia MUST dar la versión anterior con su commit, el commit del tag nuevo, el artefacto publicado (imagen y digest, o fichero distribuido y su SHA-256) y el artefacto de marcha atrás. Un dato que aún no existe (digest antes del build) MUST decir dónde figurará.
- **HR-6** — La tabla «PR integradas» y el «Registro de cierres» MUST cuadrar entre sí y con el rango. Una PR intermedia de una cadena MUST figurar como integrada por su punta («entra con #N»).
- **HR-7** — «Cómo desplegar» MUST numerar los pasos, cada uno con su comprobación; MUST incluir la copia previa cuando hay migración, las pruebas tras el despliegue y la marcha atrás.
- **HR-8** — «Pendiente tras esta versión» MUST listar lo detectado y no resuelto, cada punto con su issue o con «sin issue» y el motivo.
- **HR-9** — En un repositorio con `ci-pattern`, las notas MUST partir de su plantilla `RELEASE-NOTES.md.tmpl` y conservar sus secciones y su orden. Esta skill MUST NOT copiar ni sustituir esa plantilla.
- **HR-10** — Las exclusiones de la comprobación (`--exclude-pr`, `--exclude-issue`) MUST limitarse a la PR que añade las propias notas y a su issue, y MUST declararse en el PR de las notas.
- **HR-11** — La skill MUST NOT publicar la release, crear ni mover tags, ni elegir el número de versión.
- **HR-12** — Las notas MUST escribirse en el idioma de los artefactos del repositorio (castellano peninsular formal en la flota) y MUST NOT inventar datos: una cifra, un hash o un resultado de prueba sin fuente se omite o se marca como pendiente.

## §3 Decision Gates

| Condition | Action |
|---|---|
| El repositorio tiene `ci-pattern` (`.agents/skills/ci-pattern/assets/templates/RELEASE-NOTES.md.tmpl`) | Perfil web/contenedor; parta de esa plantilla (HR-9) y siga `references/profile-web.md`. |
| El repositorio es web/contenedor sin `ci-pattern` | Perfil web/contenedor con la lista de secciones de §4, paso 4. |
| El repositorio es VBA/Access (`*.accdb`, `.dysflow/project.json`) | Perfil VBA/Access: `references/profile-vba.md` y su esqueleto `assets/RELEASE-NOTES-vba.md.tmpl`. |
| `collect` lista commits sin PR | Cuéntelos en su grupo citando su issue o su SHA corto; en VBA es lo habitual. |
| Una PR está cerrada sin merge y su contenido entró por otra | Cítela como integrada por la punta que `collect` indica en `integrated_by` (HR-6). |
| `check` sale con 1 | Añada a las notas lo que nombra cada línea `FALTA` y repita; no entregue con fallos. |
| `check` sale con 2 | Corrija la entrada (rango o notas ilegibles); no hay veredicto. |
| Un dato de Procedencia no existe aún | Diga dónde figurará (GitHub Release, manifiesto); no lo invente (HR-12). |

## §4 Execution Steps

1. **Fijar el rango.** Pida al mantenedor el tag anterior y el nuevo (o el commit que se etiquetará). Trabaje sobre un clon con ambos.
2. **Recoger los datos.** Ejecute `python3 assets/release_coverage.py collect --repo-dir <clon> --from <anterior> --to <nuevo> --github <owner/name> --out <dir temporal>/range.json`. Lea en el perfil las demás fuentes: cuerpos de las PR, issues, artefactos y migraciones.
3. **Clasificar cada cambio** en uno de los cuatro grupos de HR-2. Una PR que toca varios grupos aparece en cada uno, con la parte que le corresponde.
4. **Redactar** en este orden: Resumen, Procedencia, los cuatro grupos numerados, PR integradas, Qué lo prueba, Cambios incompatibles, Cómo desplegar, Registro de cierres con «Pendiente tras esta versión». Con `ci-pattern`, los cuatro grupos ocupan el lugar de «Qué se corrige» y el resto de secciones conserva el nombre y el orden de la plantilla.
5. **Escribir cada entrada** con el formato de HR-3: `### <n.m> <título> — #PR (cierra #issue)`, y las líneas «En <anterior>», «En <nuevo>» y «Dónde».
6. **Comprobar la cobertura.** Ejecute `python3 assets/release_coverage.py check --range <range.json> --notes <notas>` (con `--exclude-pr` y `--exclude-issue` solo según HR-10) hasta exit 0.
7. **Entregar** las notas en el fichero que el repositorio espera (`docs/releases/<tag>.md`, `RELEASE-<tag>.md`) por PR, con la salida de `check` en el cuerpo. La publicación sigue el procedimiento de release del repositorio (HR-11).

## §5 Output Contract

| Key | Type | Description |
|---|---|---|
| `status` | `"success" \| "blocked" \| "failed"` | `success` solo con `check` en exit 0. |
| `profile` | `"web" \| "vba"` | Perfil aplicado. |
| `range` | string | `<anterior>..<nuevo>` con los dos SHA. |
| `notes_path` | string | Ruta del fichero de notas producido. |
| `coverage_exit` | integer | Exit de `release_coverage.py check` (0, 1 o 2). |
| `missing` | string[] | Líneas `FALTA` de la última ejecución; vacío con exit 0. |
| `exclusions` | string[] | PR e issues excluidas y su motivo (HR-10); vacío si no hay. |
| `pending_data` | string[] | Datos de Procedencia o de pruebas que aún no existen y dónde figurarán. |
| `next_recommended` | `"open_pr" \| "fix_missing" \| "ask_maintainer" \| "none"` | Siguiente paso. |

## §6 Anti-patterns

| Symptom | Fix |
|---|---|
| «Varias correcciones» o «mejoras internas» como entrada | Una entrada por cambio, con antes, ahora y dónde (HR-3). |
| Las notas listan las PR por número sin decir qué cambia para el usuario | Agrupar por interés del lector (HR-2); la tabla de PR va al final. |
| Un cambio de CI mezclado con los errores de producto | Grupo «Repositorio, pruebas y CI», marcado como «No cambian la aplicación.». |
| El Resumen calla una migración o un cambio de configuración | Avisarlo en el Resumen y detallarlo en «Cómo desplegar» (HR-4, HR-7). |
| Se omite una PR intermedia de una cadena | Citarla como integrada por su punta (HR-6). |
| En VBA se cuentan solo las PR y faltan los commits empujados a la rama | Ejecutar `collect` y cubrir cada commit sin PR (perfil VBA). |
| Se copia la plantilla de `ci-pattern` dentro de la skill o del repositorio | Partir de la plantilla vendorizada (HR-9). |
| Se excluye de la comprobación una PR incómoda | Solo la PR de las notas y su issue (HR-10). |
| Un hash o un digest escrito de memoria | Calcularlo de la fuente o marcarlo pendiente (HR-12). |

## §7 Resources

| Recurso | Uso |
|---|---|
| `assets/release_coverage.py` | `collect` (rango) y `check` (cobertura). Python 3 stdlib; `git` y, sin `--prs-json`, `gh`. |
| `assets/RELEASE-NOTES-vba.md.tmpl` | Esqueleto del perfil VBA/Access, con las secciones de la plantilla de `ci-pattern`. |
| `references/profile-web.md` | Fuentes y Procedencia del perfil web/contenedor. |
| `references/profile-vba.md` | Fuentes, Procedencia y despliegue del perfil VBA/Access. |

Pruebas de la comprobación: `testing/suites/release-notes/test_release_coverage.py` en el catálogo.
