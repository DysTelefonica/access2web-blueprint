# Parámetros del patrón de CI

Todo lo que varía de un repo a otro vive en esta tabla. Adopte el patrón
extrayendo estos valores primero; ninguna regla de la skill depende de los
valores concretos del consumer de origen (`ardelperal/APAP_WEB`), que figuran
solo como ejemplo verificado.

## Los veinticinco parámetros

| # | Parámetro | Valor en ardelperal/APAP_WEB (consumer de origen; ejemplo verificado) | Dónde vive | Quién lo lee |
|---|---|---|---|---|
| 1 | Regex de rama | `^(feat\|fix\|chore\|docs\|refactor\|test)/<N>-<kebab-slug>$` | Protección de rama + gate de nombre de rama | El gate deriva la issue del número `<N>`; el generador `assets/branch-name.sh` la lee del gate (`--gate-script`) o la recibe por `--pattern`, y valida el nombre antes de crear la rama (HR-35) |
| 2 | Presupuesto de revisión | `400` líneas (adiciones + eliminaciones, sin lockfiles) | Constante del gate de tamaño | `scripts/check_pr_size.py` |
| 3 | Etiquetas canónicas | `chain:partial`, `size:exception`, `type:*`, `status:approved` | Configuración de labels del repo | Gates de issue-spec y de tamaño |
| 4 | Contextos de estado por SHA | `release/smoke-production`, `release/e2e-production` | Lista cerrada de contextos | `scripts/check_release_evidence.py` |
| 5 | Rutas sensibles | Globs de auth, sesión, CSRF, migraciones y runtime (49 líneas) | `.github/release-e2e-paths.txt` | `scripts/check_release_e2e_required.py` |
| 6 | URL de salud | `https://<host>/healthz` (expone `.revision` del deploy) | Constante del smoke | `scripts/production_smoke.py` |
| 7 | User-Agent del smoke | Agente propio con reintentos y host de la redirección | Constante del smoke | `scripts/production_smoke.py` (403 de Cloudflare sin él) |
| 8 | Política de merge | Commit de fusión, rama remota conservada, auto-merge + update-branch activados, `strict: true` | Settings del repo | Gate de evidencia del deploy |
| 9 | Ruta de post-mortems | `docs/postmortems/<date>-<slug>.md` | Convención de docs del repo | HR-21 (post-mortem blameless); HR-19 (notas de release); HR-48 (la documentación del proyecto vive en su repo, y el gate de las fases 3 y 5 lo comprueba) |
| 10 | Playbook de release | `RELEASE-<TAG>.md` en la raíz del repo | Raíz del repo | HR-22 |
| 11 | Runbook e2e de producción | `docs/runbooks/e2e-production.md` | `docs/runbooks/` | Exclusión de §1; registro del veredicto (HR-10); el runbook es parte del arnés de evidencia y queda sujeto a HR-42 (premisas con comando de siembra o verificación) |
| 23 | Ubicación del playbook de CI del consumer de origen | `ardelperal/APAP_WEB@HEAD` — `odd/skill-ci-portable/ci-playbook.md` (415 líneas, «Regla 1» a «Regla 19»; no viaja con la skill). Verificado en solo lectura contra la API del host (2026-10-05). Las citas «playbook regla N» de §2 se refieren a ese documento | Referencias del consumer de origen | HR-5, HR-6, HR-7, HR-9 (citas de playbook) |
| 12 | Skills locales de testing y seguridad | `apap-testing-strategy`, `apap-security` | Catálogo local del consumer | Exclusiones de §1 |
| 13 | Runner de tests | `pytest` | Configuración de la suite local | HR-4 (paridad de preflight); §6 |
| 14 | Eventos autorizados a publicar checks requeridos | `pull_request` (el único evento que evalúa el PR); `push`, `workflow_dispatch` y `schedule` publican bajo otro nombre de contexto o fallan | Disparadores del workflow de cada gate | HR-30 (publicación de nombres requeridos); HR-17 (fallback manual) |
| 15 | Dato mutable → disparador o reevaluación en merge | Cuerpo del PR → sin disparador (editar el cuerpo no dispara el gate); etiquetas del PR → `labeled`/`unlabeled`; estado y etiquetas de la issue enlazada → sin disparador en el host: reevaluar en el momento del merge | Disparadores de los workflows de los gates + procedimiento de merge | HR-31 (verde sobre datos mutables); HR-6 (higiene del autor) |
| 16 | Actores exentos y patrones de rama de plataforma | Actores exentos: el bot de dependencias del host y el actor de propagación del catálogo; patrones de rama de plataforma: `revert-<n>-<rama>` (botón Revert), el de cada ecosistema de actualización de dependencias habilitado y `skill-fleet/<consumer>` (propagación del catálogo); toda exención exige actor Y origen en el mismo repositorio, nunca solo el patrón | Lista de datos leída por el gate de nombre y el gate de trazabilidad | HR-33 (identidad verificable); gates de nombre e issue-spec |
| 17 | Contrato de host y drift check | `assets/host-readback/host-contract.example.json` (etiquetas, checks requeridos, métodos de merge, rulesets con su `enforcement`, flags de protección; cada regla `host-enforced` o `documented-only`) contrastado con instantáneas JSON de la API (repo, labels, branch-protection, rulesets) capturadas con GETs de solo lectura | `assets/host-readback/` | HR-34 (drift check; clasificación de reglas); paso 2 de la adopción §4 |
| 18 | Rutas de alto riesgo para revisión | Lista de globs mantenida como dato, con la misma forma que el parámetro 5 («rutas sensibles»); valores concretos aún sin fijar (el gate de evidencia es asset futuro, listado en S2) | Fichero de datos del consumer | HR-35 (evidencia de revisión); gate futuro |
| 19 | Campo de evidencia de revisión | `review-evidence:` en el cuerpo del PR, una sola línea, con `lens=`, `verdict=` y `reference=` | Cuerpo del PR | HR-35; gate futuro de evidencia |
| 20 | Ventana de la métrica de salud del presupuesto | 30 PRs fusionados (decisión del operador, 2026-10-02) | Métrica de salud (asset futuro) | HR-36 |
| 21 | Umbral de tasa de excepciones | 40% de los PRs de la ventana (decisión del operador, 2026-10-02) | Métrica de salud (asset futuro) | HR-36 |
| 22 | Manifiestos que disparan el gate de avisos en PR | Lista de rutas de manifiestos de dependencias (p. ej. `requirements*.txt`, `pyproject.toml`); valores concretos aún sin fijar en el consumer | Fichero de datos del consumer | HR-39 (gate programado + PRs que tocan manifiestos) |
| 24 | Ruta del registro operativo de incidentes | `docs/incidents/` — el mismo repositorio de la aplicación; un repositorio de operaciones aparte es opcional | `ci-pattern.yaml` (`P49_incident_log_dir`), con default del esquema | HR-21 (detalle técnico del incidente: IPs, hostnames, comandos, logs, cronología y estado de cada ticket); gate de fase 5 |
| 25 | Gestor de secretos declarado | «el gestor de secretos del consumer» (p. ej. Vault, OpenShift Secrets, 1Password) | `ci-pattern.yaml` (`P50_secrets_manager`), con default del esquema | HR-21 (credenciales y tokens: nunca en git, ni en un repositorio privado) |

El nombre de rama (parámetro 1) se **genera**, nunca se prescribe de memoria:
`assets/branch-name.sh` construye `<tipo>/<N>-<slug>`, normaliza el slug
 de
forma determinista (minúsculas, sin acentos, guiones), valida contra la regex
—leída del `check_branch_name.py` del destino o pasada por `--pattern`— e
incluye auto-test ejecutable sin el repo destino (HR-35). Las opciones
(`--pattern`, `--max-len`, `--gate-script`) se aceptan en cualquier orden,
antes o después del nombre: `--max-len` debe ser un entero positivo (se
rechaza con mensaje explícito) y los nombres con salto de línea, retorno de
carro o vacíos se rechazan como no representables como rama. La plantilla de
encargo que verifica el resto de parámetros en vivo vive en
`references/delegation-template.md` (HR-36).

Parámetros secundarios documentados en las fuentes: host de producción, versión
de Python y lista de checks requeridos (tres en este repo: `branch-name`,
`required`, `pr-size / pr-size`).

## Policy file de gates dormibles (HR-18)

Un gate que pierde su justificación se duerme tras un policy file versionado;
nunca se elimina el candado construido. Esquema canónico del fichero —el que
implementa y valida `scripts/preflight.py` de `ardelperal/APAP_WEB`
(`ALLOWED_GATE_KEYS` / `VALID_ENFORCEMENT`, issue #1168) con su suite
(`tests/test_ci_gate_policy.py`)—, una entrada por gate dormible:

```json
{
  "gates": {
    "<nombre-del-gate>": {
      "enforcement": "dormant",
      "reason": "motivo de la dormición con su evidencia",
      "dormant_since": "AAAA-MM-DD"
    }
  }
}
```

Contrato del policy file (validado fail-loud: cualquier desvío levanta
`GatePolicyError` y el paso falla con código de salida propio; nada se
adivina ni se aplica en silencio):

- Nivel superior: un objeto con un objeto `gates`; una entrada por gate.
- `enforcement` es obligatorio y admite exactamente `dormant` y `enforcing`;
  otro valor o su ausencia invalida el fichero completo.
- `reason` es obligatorio y no vacío cuando `enforcement` es `dormant`: es el
  campo de datos que sostiene la dormición.
- `dormant_since` es opcional; si está, cadena no vacía (fecha de la
  transición a `dormant`).
- Ninguna otra clave está permitida: una clave desconocida rechaza el
  fichero completo.
- Fichero ausente equivale a política vacía: todos los gates quedan
  `enforcing` (default-deny: sin política no hay exención).
- Comportamiento: el gate dormido se ejecuta siempre y sus hallazgos se
  imprimen (nunca un falso verde silencioso); `dormant` devuelve 0 y
  etiqueta la salida como informativa; `enforcing`, un gate no listado o una
  política ausente devuelven el código de salida del propio gate.
- La re-activación es un cambio de datos de este fichero
  (`dormant` → `enforcing`) que pasa por review como cualquier PR.

Extensiones declaradas en versiones anteriores de esta tabla pero no
implementadas hoy por el validador de referencia —por HR-34 son
`documented-only`, no contrato de gate, hasta que el validador las
acepte—: `findings_exit_code` (limitar la supresión de `dormant` a un único
código de salida del gate), `activation_snapshot`, `grandfathered_entries`
y el campo `relaxation` (motivo + referencia obligatorios para toda
relajación respecto de la rama base). No escriba esas claves en el fichero:
hoy lo invalidarían como claves desconocidas.

Nota de reconciliación (2026-10-02, hueco G9 del inventario de activos
portable): este doc describía un esquema dormido (`policy_version`,
`activation_snapshot`, `grandfathered_entries`) que ningún fichero ni
validador implementa; el esquema canónico es el de arriba, que sí tiene
validador y tests.

Fuente del patrón: `grandfather-size-exceptions.json` de
Gentleman-Programming/gentle-ai (idea T1 del benchmark); regla R15 destilada
en la auditoría de origen (veredicto por gate en
`references/gate-verdicts.md`). Implementación y validador de referencia:
`scripts/preflight.py` de `ardelperal/APAP_WEB`.

## Policy file del agregador de jobs requeridos (HR-15, HR-29)

El agregador (`assets/required-jobs/`) lee su propia política: jobs conocidos
y skips aceptados por evento. Forma mínima del fichero (JSON):

```json
{
  "required_jobs": ["<job>"],
  "events": {
    "<evento>": { "accepted_skips": [] }
  }
}
```

Contrato de la política del agregador:

- `required_jobs` y `events` son obligatorias; `accepted_skips` lo es en cada
  evento declarado. Una clave desconocida invalida el fichero completo.
- Relajación respecto a la copia de la rama base: un job retirado de
  `required_jobs`, una entrada añadida a `accepted_skips` o un evento añadido.
  Endurecer no exige nada.
- Toda relajación `MUST` llevar el campo `relaxation` con `reason` y
  `reference`, cadenas no vacías y distintas de las que ya lleve la copia
  base: un campo heredado no ampara una relajación nueva (HR-15). El campo
  se valida siempre que esté presente: uno malformado no puede quedar
  esperando a "cubrir" una relajación futura.
- El gate recibe la copia base por `--base-policy`. Sin ese argumento, la
  salida declara que la comparación no se hizo (HR-3).
- Diferencia con el policy file de gates dormibles: la ejecución se evalúa con
  la política del PR, no con la de la rama base. HR-29 obliga a añadir job,
  `needs` y política en el mismo PR, y la copia base no conocería el job
  nuevo. La copia base solo decide si hay relajación.

## Scripts de implementación de referencia

El agregador de jobs requeridos ya está distribuido como asset portátil de
esta skill: `assets/required-jobs/` (gate fail-closed + comprobación de
paridad contra el workflow del consumer + política de ejemplo + suite; véase
su `README.md`). El resto sigue
versionado en el repositorio de origen del patrón, `ardelperal/APAP_WEB`
(bajo `scripts/` y `.github/`), pendiente de publicarse como asset portátil;
adáptelos a los parámetros de arriba, no al revés:

| Script | Función |
|---|---|
| `scripts/preflight.py` | Lee el job de lint del workflow y lo reproduce en local, paso a paso |
| `scripts/check_pr_size.py` | Presupuesto de revisión con excepción declarada en el cuerpo |
| `scripts/check_issue_specs.py` | Trazabilidad determinista: rama, etiquetas, cierre de issue |
| `scripts/check_release_evidence.py` | Evaluador fail-closed de un contexto de estado por SHA |
| `scripts/check_release_e2e_required.py` | Decide si el rango toca rutas sensibles |
| `scripts/production_smoke.py` | Smoke sin autenticación, con User-Agent propio y reintentos |
| `.github/release-e2e-paths.txt` | Lista de rutas sensibles, un glob por línea |

## Formato de `ci-pattern.yaml`

El validador (`ci-pattern params validate`) rechaza cualquier entrada ambigua o
mal tipada. Reglas que el adoptante debe respetar:

- **Las regex van SIEMPRE entrecomilladas** (`P01_branch_name_pattern: "^(feat\|fix\|…)$"`):
  sin comillas, una expresión que empiece por `[` se interpreta como lista y el
  parámetro se rechaza por tipo incorrecto.
- **Los paths son rutas relativas POSIX del repo**: sin `/` inicial, sin `..`, sin
  letra de unidad de Windows (`C:\…`) y sin separador ni ruta UNC de Windows
  (`\\servidor\…`). La ruta de salud de P19 es la excepción tipada:
  `P19_health_path` es una **ruta URL** que empieza por `/` (p. ej. `/healthz`).
- **Las listas tipan sus elementos** (`item_type` en el esquema) y se validan
  elemento a elemento; los objetos declaran sus `fields` y rechazan campos
  desconocidos.
- **Ni claves ni subclaves duplicadas**: la subclave repetida es entrada inválida
  (exit 2), no una sobrescritura silenciosa.
