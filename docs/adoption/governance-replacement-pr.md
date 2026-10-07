# Fase 3 — Reemplazo de gobernanza: qué se retira, qué se instala y de dónde sale

> **Qué es.** La evidencia del PR de la fase 3 de la adopción del patrón de CI en este repositorio: la procedencia de cada plantilla, las retiradas ejecutadas, las desviaciones conscientes y lo que queda pendiente con su seguimiento.

> **Qué no es.** Un cambio de los gates de calidad del repositorio: los 14 `scripts/check_*.py`, el plugin de cobertura y sus umbrales quedan como están (decisión ratificada 6).

## Procedencia de las plantillas canónicas

El patrón todavía no trae sus plantillas de serie: `assets/templates/` está vacío a propósito y lo sigue la fricción `DysTelefonica/team-skills#376`. Esta adopción no espera a esa wave (no tiene fecha) y deriva las plantillas de dos fuentes que sí existen: la adopción cerrada de referencia y el contrato que los gates del patrón leen como dato.

| Fichero instalado | Origen | Identidad del origen | Qué se copió |
|---|---|---|---|
| `.github/ISSUE_TEMPLATE/issue-canonical.yml` | `DysTelefonica/cadete` | `main` en `b9fc2af2c244c9d484628de8a74246552a59544b`, blob `e11be9867f0fe4efff72a57e9b465c5ed94dbf0f` | Las seis secciones, sus `id` y su orden; la forma de `config.yml` |
| `.github/ISSUE_TEMPLATE/config.yml` | `DysTelefonica/cadete` | mismo commit, blob `8005e3226730ef74f37ae9614ba94a1bb879b4a0` | Los dos campos, tal cual |
| `.github/PULL_REQUEST_TEMPLATE.md` | `DysTelefonica/cadete` | mismo commit, blob `b6ad14023992f4c7f7381f8b85010e8f4a177427` | El campo de excepción de tamaño, `Closes #` y la checklist |

Lo que exige cada gate del patrón, y por eso está donde está:

| Pieza de la plantilla | Contrato que la lee |
|---|---|
| Las seis secciones del cuerpo de la issue, con sus nombres exactos | `ci-pattern/SKILL.md:700-704`: contrato documental (la contraparte issue-side no es de gate) |
| Encabezado `Chain Context` con sus ocho campos y un 📍 | `assets/pr-contract/check_pr_contract.py` (HR-54) |
| Encabezado `Tests que prueban el cierre` | el mismo gate (HR-53) |
| `Closes #N` / `Refs #N` | el mismo gate (HR-6, HR-7) |
| `size-exception-reason:` en una sola línea | el mismo gate (HR-8) y el gate del propio repositorio, `scripts/check_pr_size.py`, que además exige la etiqueta `size:exception` |
| Regla de nombre de rama | `scripts/check_branch_name.py:26`, la del repositorio |
| `status:approved` por un mantenedor | decisión ratificada (2026-10-07); la etiqueta existe en el host |

## Qué retira esta fase

| # | Artefacto | Acción | Estado |
|---|---|---|---|
| 1 | `.github/ISSUE_TEMPLATE/bug.yml` | Retirado en la primera rebanada, junto con la instalación del formulario canónico (G3.1) | Fuera del árbol |
| 2 | `.github/ISSUE_TEMPLATE/docs.yml` | Igual | Fuera del árbol |
| 3 | `.github/ISSUE_TEMPLATE/feature.yml` | Igual | Fuera del árbol |
| 4 | `docs/calidad-de-codigo-y-ci.md` | Retirado en esta rebanada | Era documentación escrita a mano y ya derivaba (afirmaba que no hay `package.json` mientras `.github/dependabot.yml` declara el ecosistema `npm` sobre `/e2e`) |
| 5-10 | etiquetas `bug`, `chore`, `ci`, `documentation`, `enhancement`, `refactor` | **Acción de mantenedor** (abajo) | 0 issues abiertas cada una |
| 11 | etiqueta `feat` | **Acción de mantenedor**: reetiquetar y después borrar | 5 issues abiertas: #226, #252, #253, #254, #260 |
| 12 | etiqueta `test` | **Acción de mantenedor**: reetiquetar y después borrar | 2 issues abiertas: #255, #256 |

Ningún agente aplica etiquetas ni borra etiquetas en este repositorio: los comandos van en el cuerpo del PR para que los ejecute un mantenedor.

Lo que **no** se retira y conviene saber: `size:exception` se conserva (decisión ratificada 4: el gate de tamaño del repositorio exige la etiqueta además del campo de datos, y retirarla obligaría a cambiar ese gate). Las etiquetas de producto (`app/<slug>`, `cross-cutting`, `epic:*`, `dependencies`, `python:uv`, `duplicate`, `good first issue`, `help wanted`, `invalid`, `question`, `wontfix` y las `status:*`) se conservan.

## Qué instala esta fase

| Fichero | Papel |
|---|---|
| `.github/ISSUE_TEMPLATE/issue-canonical.yml` y `config.yml` | El formulario canónico; sin formulario en blanco |
| `.github/PULL_REQUEST_TEMPLATE.md` | La plantilla canónica del PR |
| `.github/host-contract.json` | Las reglas del host con su clase, releídas de la API |
| `docs/adoption/branch-protection.json` y `docs/adoption/rulesets.json` | Las instantáneas literales de esa relectura |
| `.governance-manifest.json` | La identidad de los artefactos del patrón: ruta y sha256 |
| `AGENTS.md` | El bloque de gobernanza del flujo de entrega y del merge, por roles |

`phases."3".pattern_paths` declara los seis ficheros del manifiesto; el gate comprueba que cada uno está cubierto por él y que su sha256 coincide con el árbol.

## Desviaciones conscientes respecto de la adopción de referencia

| Desviación | Por qué |
|---|---|
| La plantilla de PR **sí** trae `## Chain Context` con los campos | HR-54 vigente en el patrón; la plantilla de Cadete es anterior a esa regla |
| El formulario no trae campo de aplicación | La dimensión de aplicación ya vive en las etiquetas `app/<slug>` del host (8) y en el prefijo del título, como ratificó el operador |
| `AGENTS.md` y `.github/ci-pattern-adoption.json` **no** entran en el manifiesto de gobierno, aunque Cadete los congele | `AGENTS.md` lo reescribe cada corrida de propagación del catálogo (el bloque de slice se reestampa) y el contrato lo edita cada fase: congelar sus hashes convertiría cada propagación o cada fase siguiente en un fallo de fase 3. Su identidad la gobiernan sus propios mecanismos (marcadores del slice y el gate de cada fase) |
| Las instantáneas de esta fase son dos (`branch-protection` y `rulesets`) | Es lo que la fase 3 pide. Las de `repo` y `labels` completan el juego del gate de deriva en la fase 4 |
| `chain:partial` se declara `documented-only` | La etiqueta no existe todavía en el host; la crea un mantenedor. Cuando exista, su clase pasa a `host-enforced` |

## Pendientes declarados, con su seguimiento

1. **La etiqueta `chain:partial` no existe en el host.** La crea un mantenedor con el comando del cuerpo del PR. Mientras tanto la cadena se declara como dato en los cuerpos.
2. **El documento retirado está citado en 53 sitios de 19 ficheros** (medido con `grep -rn`, contando navegación, tablas y comentarios). En `AGENTS.md` se repunta la lectura obligatoria a los propios gates, porque es el fichero que esta fase sustituye; el resto de las citas —`README.md`, `DOCS.md`, `CODEBASE-GUIDE.md`, `docs/testing/*`, `openspec/**`, `skills/*`, `app/README.md`— las reapunta el PR de la fase 5 al documento operativo generado, que es su sustituto. Ancla: `#781`.
3. **El patrón no trae las plantillas de serie.** `DysTelefonica/team-skills#376` lo sigue; esta fase deriva y declara la procedencia.

## Validación ejecutada

```text
$ python3 <skill>/assets/bin/ci-pattern adoption check --phase 3 .
adoption check: fases 0..3 en verde sobre .
EXIT=0

# la identidad del manifiesto se verifica contra el árbol: si un fichero del
# manifiesto cambia y el manifiesto no, la fase 3 falla nombrando el fichero

$ python3 <skill>/assets/host-readback/check_host_drift.py \
    --contract .github/host-contract.json \
    --snapshot repo=<instantánea> --snapshot labels=<instantánea> \
    --snapshot branch-protection=docs/adoption/branch-protection.json \
    --snapshot rulesets=docs/adoption/rulesets.json
VERDICT: PASS
EXIT=0
# con 'chain:partial' como única regla documented-only listada

$ python3 -c "import yaml; yaml.safe_load(open('.github/ISSUE_TEMPLATE/issue-canonical.yml'))"
YAML OK   # seis secciones, nombres exactos, en orden
```
