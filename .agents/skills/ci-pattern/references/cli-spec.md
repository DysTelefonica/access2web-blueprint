# Especificación de la CLI `ci-pattern`

Adoptador/actualizador/verificador determinista del patrón de CI. Diseñada para
que una IA la invoque con una frase («actualizá el sistema de gobernanza»,
«adoptá el patrón en este repo», «verificá el cumplimiento») y ejecute sin
imaginar nada: cada comando lee datos estructurados y emite una salida terse,
agrupada y accionable.

- **Implementación:** `assets/bin/ci-pattern` (relativo a la raíz de la skill;
  canónica en `DysTelefonica/team-skills`, `~/personal-skills`, y resuelta igual
  en cualquier mirror). Python 3 stdlib exclusivamente — sin venv, sin dependencias
  de terceros, ejecutable con `python3` y desde CI.
- **Estado de esta wave:** entregados `params validate`, `verify`, `status`, el
  formato de manifiesto y su lectura/escritura, y `adoption check` (#276). `adopt`
  y `update` están retirados: la adopción la gobierna una IA con el contrato de
  adopción y su gate por fase — este documento lo declara, no lo simula.

---

## §1 Comandos

| Comando | Estado | Qué hace |
|---|---|---|
| `ci-pattern params validate <file>` | **Entregado** | Valida un fichero de parámetros contra `assets/parameters.schema.json` (§5). |
| `ci-pattern verify <repo>` | **Entregado** | Comprobación de cumplimiento dirigida por el manifiesto (§4). |
| `ci-pattern status <repo>` | **Entregado** | Resumen de una línea por clase: manifiesto presente, ficheros OK/modificados/ausentes, gates cableados, bloques de slice. Exit: 0 sólo si está limpio, 1 con hallazgos, 3 sin manifiesto (§2). |
| `ci-pattern manifest show <repo>` | **Entregado** | Imprime el manifiesto leído (lectura; la escritura vive en la librería y la ejercitan los tests). |
| `ci-pattern adoption check (--phase N \| --all) <repo>` | **Entregado** | Gate por fase de la adopción (#276): valida el contrato `.github/ci-pattern-adoption.json` del consumer y verifica cada fase del porting-guide contra el árbol real. Fase 0 inventario exhaustivo (adopted/retired sin tercera opción; los targets `adopted` pertenecen al catálogo canónico declarado en `assets/adoption/pattern-catalog.json`, sin exigir existencia —#306/#311—; cobertura ejecutada sobre las rutas gobernadas del catálogo, los scripts descubiertos por REFERENCIA o declarados en `declared_governed_paths` —#311— y las etiquetas del host vía `host_labels` —#312—; `removal_confirmed`); fase 1 evidencia de propagación; fase 2 aislamiento (EJECUTA el gate HR-28); fase 3 sustitución de gobierno e identidad del manifiesto (#277) —las instantáneas del host aceptan el **403 de plan** («Upgrade to GitHub Pro…») como evidencia de capacidad no disponible solo si el contrato declara esa área `unavailable` en `host_capabilities`; con `available` es drift (exit 1) y sin declaración —o con un valor fuera del vocabulario `available`/`unavailable`— es duda (exit 2), #330— y HR-48 sobre `AGENTS.md` (ninguna instrucción manda la documentación fuera del repo); fase 4 los cinco gates del patrón en verde; fase 5 documento operativo con gate de deriva (#279) que incluye la sección «Dónde vive la documentación» y HR-48 sobre `AGENTS.md`, el documento y los ficheros que estos enlazan o citan (#319); EJECUTA además el gate de HR-21 sobre los post-mortems (parámetro 9) y el registro operativo (`P49_incident_log_dir`, por defecto `docs/incidents/`): bloquea secretos, datos personales y evidencia sin `sha256` ni ubicación, y permite IPs, hostnames y comandos internos: el sujeto son los ficheros de TEXTO del registro (`.md`, `.log`, `.txt`, `.json`, `.yaml`, saltando binarios) y superar el límite de escaneo no da verde parcial —sale exit 2 nombrando el recuento y el límite (#322); fase 6 cadena de aceptación (número, SHA de 40 hex, `conclusion: success`). Ninguna fase pasa si la anterior no pasa. Exit: 0 cumple · 1 hallazgos nombrando cada artefacto o condición que falta · 2 contrato ilegible, esquema inválido o sujeto vacío. |
| `ci-pattern adoption generate-doc <repo> --out <ruta>` | **Entregado** | Documento operativo del consumer (#279): flujo issues → ramas → commits → PR → checks → merge → release → fricciones, GENERADO desde sus datos (ci-pattern.yaml, host-contract, required-jobs, formularios de issue). Determinista byte a byte; nunca a stdout; entrada ilegible o ausente → exit 2. La deriva la detecta el gate de la fase de documentación operativa (#279). |
| `ci-pattern adopt <repo>` | **Retirado** | Redirige a `adoption check` (exit 2 con mensaje accionable). Esta CLI nunca copia plantillas ni escribe el manifiesto: la adopción la gobierna una IA con el contrato de datos. |
| `ci-pattern update <repo>` | **Retirado** | Redirige a `adoption check` (exit 2 con mensaje accionable). La re-sincronización se gobierna por el manifiesto y el veredicto de `adoption check`. |

Flags globales: `--json` (salida máquina en `verify`, `status` y
`manifest show`), `--slice-canonical <dir>` en `verify` para contrastar los
bloques de slice contra el catálogo canónico.

`verify <repo>` extiende su veredicto con la clase `adoption` (#276): si el
consumer declara `.github/ci-pattern-adoption.json`, `verify` exige que
`adoption check --all` esté en verde — hallazgo `phase-incomplete` (y nota de
fases no evaluadas) hace exit 1; un contrato ilegible hace exit 2; sin
contrato, `verify` no cambia.

## §2 Códigos de salida

| Código | Significado |
|---|---|
| `0` | Limpio: sin hallazgos o validación correcta. |
| `1` | Hallazgos: la verificación encontró desviaciones o la validación rechazó el fichero. |
| `2` | Error de uso o entrada inválida: comando o argumentos desconocidos; manifiesto sin nada verificable; `P23_required_jobs` mal tipado (no lista de cadenas no vacías); claves de `files` que no son rutas relativas seguras (absolutas, con `..` o con letra de unidad Windows). Sin traceback. |
| `3` | Recurso ausente o ilegible: manifiesto o fichero de parámetros inexistente/corrupto. |

## §3 Formato del manifiesto (`.governance-manifest.json`)

Vive en la raíz del repo adoptante. La escritura vive en la librería y los
tests ejercitan el round-trip; el flujo de adopción vigente es el contrato
de adopción + `adoption check` (#276).

```json
{
  "schema_version": 1,
  "adopted_at": "2026-10-02T12:00:00Z",
  "canonical_version": "0.4",
  "canonical_source": {
    "repo": "DysTelefonica/team-skills",
    "path": "personal/ardelperal/ci-pattern"
  },
  "parameters": {
    "P06_review_budget_lines": 400,
    "P10_label_approval": "status:approved"
  },
  "files": {
    "scripts/check_pr_size.py": "sha256:…",
    ".github/workflows/pr-size.yml": "sha256:…"
  },
  "slice_blocks": {
    "APAP_WEB": "sha256:…"
  }
}
```

- `schema_version`: `1`. Un manifiesto con otra versión se rechaza con exit 3.
- `canonical_source`: repo + ruta del catálogo canónico del que se adoptó.
- `parameters`: mapa `P<NN>_<snake_name>` → valor validado contra el esquema.
- `files`: ruta repo-relativa → `sha256:<hex>` del contenido adoptado.
- `slice_blocks`: nombre del bloque de slice (`<!-- personal-skills:slice:NAME @ … -->`
  en `AGENTS.md`) → `sha256:<hex>` del contenido entre marcadores (sin las líneas
  de marcador, saltos incluidos tal cual).

## §4 Contrato de `verify`

`verify <repo>` es de solo lectura y NUNCA escribe ni corrige. Comprobaciones,
agrupadas por clase en la salida:

1. **`manifest`** — existe, es JSON válido, `schema_version: 1`. Ausente: exit 3
   con la instrucción de generarlo (la adopción la gobierna una IA con el
   contrato de adopción + `adoption check`, #276).
2. **`files`** — cada ruta del manifiesto existe y su sha256 coincide. Los hashes
   se calculan sobre el contenido con saltos normalizados a LF (regla única para
   ficheros y slices): una copia CRLF de un fichero LF (`core.autocrlf` en
   Windows) NO es hallazgo. Divergencia:
   hallazgo `locally-modified` (el contenido del repo manda; el manifiesto se
   actualiza vía `update`, no a mano). Ausencia: hallazgo `missing`. La CLI jamás
   sobrescribe ni restaura ficheros.
3. **`gates`** — cada job de `P23_required_jobs` debe aparecer como clave directa del
   bloque top-level `jobs:` en algún fichero de `.github/workflows/*.yml`
   (detección estructural por indentación: una clave anidada bajo `with:` u otro
   bloque no cuenta como job cableado). Un job sin cablear es hallazgo
   `gate-not-wired` con el nombre y la clase de hallazgo.

   Antes de las comprobaciones, `verify` valida la forma del manifiesto y falla
   cerrado con exit 2 (§2) si no hay nada verificable (`files` y `slice_blocks`
   vacíos y `P23_required_jobs` vacío), si `P23_required_jobs` no es una lista de
   cadenas no vacías, si alguna entrada no es un job id válido de GitHub
   (criterio más abajo), o si alguna clave de `files` no es una ruta relativa
   segura (criterio más abajo). La validación de claves es total y previa: se
   comprueban TODAS las claves de `files` antes de leer cualquier fichero y el
   error lista todas las inválidas, no solo la primera (#232).

   **Job ids válidos (#232).** Cada entrada de `P23_required_jobs` debe casar
   `^[A-Za-z_][A-Za-z0-9_-]*$` (la gramática de job ids de GitHub Actions):
   empieza por letra o `_`, y sigue con letras, dígitos, `_` o `-`. Una entrada
   inválida (por ejemplo con un espacio o un Unicode invisible como
   `unit\u200bsuite`) sale con exit 2 nombrándola; jamás produce un falso
   `gate-not-wired`. Válidos: `lint`, `lint_job-1`, `Unit Suite` NO (espacio).

   **Predicado de ruta relativa segura (#182, precisado en #232).** Una clave de
   `files` es una ruta relativa segura si y solo si:

   - no es vacía;
   - `PurePosixPath(key)` no es absoluta, y la clave no empieza por `/` ni `\`;
   - ningún componente es `..` y la clave no contiene `..` como subcadena;
   - no empieza por letra de unidad Windows (`^[A-Za-z]:`).

   Las formas que POSIX normaliza — `foo//bar`, `foo/./bar` — son ACEPTADAS:
   `PurePosixPath` las colapsa (`foo/bar`) y no pueden escapar del repositorio.
   Ejemplos aceptados: `scripts/check_pr_size.py`, `foo//bar`, `foo/./bar`.
   Ejemplos rechazados (exit 2, nombrando la clave): `/etc/hostname` (absoluta),
   `../x` y `a/../b` (escapada), `C:\x` y `C:/x` (letra de unidad).
4. **`slice_blocks`** — para cada nombre: el bloque existe en `AGENTS.md` y su
   hash coincide con el manifiesto (`locally-modified`, `missing` en su defecto).
   Con `--slice-canonical <slices>`, resuelve el cuerpo canónico de forma
   determinista, replicando la composición de `propagate-team-skills.ps1` sin
   tocarlo: localiza el consumer en `<slices>/../fleet/registry.json` por su
   `name` (el del bloque `personal-skills:slice:<name>`) y compone
   `slices/partials/<primary_type>.md` + el fragmento que resuelve
   `Resolve-SliceFragment` (primero `slices/<canonical_name>/` y luego
   `slices/<name>/`); cada fuente se recorta por la derecha (TrimEnd), las
   piezas vacías no cuentan (truthiness del propagador) y las presentes se unen
   con una línea en blanco. El bloque del consumer lleva el marcador
   `@ v<sha7>` con `sha7 = sha256(cuerpo compuesto)[:7]`, y el bloque se monta
   como `sliceOpen + "\\n" + cuerpo + "\\n" + sliceClose` (convención fijada por
   test). Hallazgos: `marker-mismatch` si el marcador del bloque no es la
   versión de la composición canónica; `stale-vs-canonical` si el cuerpo
   difiere de ella (el consumer quedó desfasado tras un avance del catálogo);
   `canonical-unavailable` si falta `fleet/registry.json`, el consumer no está
   en él, o no hay partial ni fragment. Sin ese flag, la salida lo declara
   como advertencia: la comparación canónica no se ejecutó.

Salida no-JSON: cabecera por clase y un hallazgo por línea
`<clase>: <detalle accionable>`. `--json`: `{"repo", "clean", "findings":
[{"class", "code", "detail"}], "canonical_comparison": "performed"|"skipped",
"head_sha", "cli_version", "schema_version", "manifest_sha256"}`. Los cuatro
últimos campos atan la evidencia a una revisión concreta (HR-16):
`head_sha` es el HEAD del repo verificado (`null` si no hay git ni commits),
`cli_version` la versión de la skill (frontmatter de `SKILL.md`),
`schema_version` la del manifiesto y `manifest_sha256` el sha256 del fichero
`.governance-manifest.json` tal cual está en disco. El JSON es determinista:
dos ejecuciones sobre el mismo commit y manifiesto son idénticas byte a byte
(sin timestamps). Exit `0` limpio, `1` con hallazgos.

## §5 Parámetros y su validación

El contrato de datos es `assets/parameters.schema.json` (P01–P48); el fichero del
adoptante es `ci-pattern.yaml`. El parser acepta un subconjunto estricto de YAML:
mapas anidados a cualquier profundidad, listas en bloque (`- item`) o inline
(`[a, b]`) a cualquier nivel, escalares (entero, flotante, booleano, cadena con o
sin comillas) y comentarios `#`. Un mapa se declara en líneas indentadas; `{}`
declara el mapa vacío y cualquier otro mapa inline es error con número de línea.
Tabulación, clave duplicada, subclave duplicada,
indentación ambigua o mezcla mapa/lista en una misma clave son error con número
de línea. Las regex se escriben SIEMPRE entrecomilladas: sin comillas, una
expresión que empiece por `[` se interpreta como lista.

Cada entrada del esquema declara: `id`, `name`, `type`, `default`, `required`,
`consumed_by` (evidencia `fichero:línea` del origen) y restricciones del tipo
(`min`, `choices`, `pattern`, `value_type`, `item_type`, `fields`). Tipos:
`regex`, `string`, `label`, `path`, `url_path`, `integer`, `float`, `boolean`,
`list`, `map`, `enum`, `object`.

`params validate <file>` rechaza con exit 1 y un mensaje por error, siempre
nombrando el parámetro (id + name) y la forma esperada:

- **Clave desconocida** — «no está en el esquema; revise el nombre o elimínelo».
- **Faltante obligatorio** — «obligatorio, sin default seguro; establézcalo».
- **Tipo incorrecto** — «debe ser <type>, se recibió <valor de tipo X>».
- **Regex inválida** — el error de `re.compile` del parámetro.
- **Etiqueta mal formada** — debe ser `prefijo:valor` (p. ej. `status:approved`).
- **Ruta ilegal** — absoluta, con `..`, con letra de unidad de Windows (`C:\…`),
  con separador o ruta UNC de Windows (`\\servidor\…`): el manifiesto y los
  paths operan con rutas relativas POSIX.
- **Elemento de lista mal tipado** — las listas declaran `item_type` y cada
  elemento se valida contra él (error con la posición 1-based).
- **Campo de objeto inválido** — los objetos con `fields` validan cada campo
  declarado, rechazan campos desconocidos y objetos vacíos.
- **Subclave duplicada** — exit 2 (entrada inválida): antes se sobrescribía en
  silencio.

Valida también los defaults declarados en el esquema: el propio esquema no puede
contener un valor que su validador rechazaría (barrido cubierto por test;
`P19_health_path` usa `url_path`, no `path`).

## §6 Contrato de idempotencia (vigente para cualquier re-sincronización)

`update` está retirado (#276); el contrato rige para cualquier mecanismo de
re-sincronización con la canónica que un consumer aplique sobre ficheros
adoptados.

Re-ejecutar la herramienta sobre un repo ya adoptado no altera bytes generados
ni duplica bloques. Cada fichero generado se registra por sha256 en el manifiesto;
`update` solo reescribe cuando el hash canónico cambió, y jamás toca un fichero
modificado localmente sin reportarlo como hallazgo `locally-modified`.

## §7 Lo que la CLI NUNCA hace

- Nunca ejecuta `git merge`, `git push`, ni publica nada.
- Nunca toca producción (ninguna llamada de red salvo la que el operador pase
  por argumento; `verify` es 100 % filesystem).
- Nunca borra ficheros que ella misma no escribió; jamás ejecuta `rm` ni
  equivalentes.
- Nunca sobrescribe un fichero modificado localmente sin reportarlo como
  hallazgo.
- Nunca decide el contenido de los parámetros: los lee, los valida y falla
  fuerte; los valores los elige el operador.

## §8 Frontera honesta (juicio humano/IA)

Estos puntos quedan fuera del determinismo, con procedimiento acotado:

- **Los valores de los parámetros** (qué presupuesto, qué etiquetas): el esquema
  valida la forma, no la intención. Decisión del operador.
- **Los conflictos de contenido** tras un `update` (fichero modificado en el
  destino y canónica avanzada): la CLI los reporta; la disposición es humana.
- **La disposición de los veredictos** de `verify`: limpio/no limpio es
  determinista, qué corregir y en qué orden es juicio del agente con la lista
  accionable en la mano.
- **La completitud de una spec, la causa raíz de un rojo, si una edición es
  material**: según `references/porting-guide.md` y el playbook del origen.

## §9 Receta de uso para una IA (flujo de 5 comandos)

Diseñada para que la salida de cada paso alimenta el siguiente, sin prosa
intermedia:

```bash
# 0. Forma portable (equivale a invocar `ci-pattern …` directo; el script
#    lleva shebang y está en modo 100755):
python3 <skill>/assets/bin/ci-pattern status <repo>

# 1. Estado actual del repo destino (¿hay manifiesto? ¿está limpio?)
ci-pattern status <repo>

# 2. Validar los parámetros del adoptante ANTES de escribir nada
ci-pattern params validate ci-pattern.yaml

# 3. Gate por fase de la adopción: el consumer declara su contrato y el
#    gate verifica cada fase del porting-guide sobre el árbol real (#276)
ci-pattern adoption check --all <repo>

# 4. Verificación de cumplimiento tras cualquier cambio
ci-pattern verify <repo> --json
```

Frases de disparo del usuario: «actualizá el sistema de gobernanza»,
«adoptá el patrón en este repo», «verificá el cumplimiento del patrón de CI».
El flujo hoy entrega 1, 2 y 5; 3 y 4 requieren la wave del scaffolder.

---

## §10 Entrega de la copia runtime (`install-ci-pattern-cli.sh`)

Esta skill vive en la canónica `personal/ardelperal/ci-pattern/` del catálogo y
está DELIBERADAMENTE sin graduar a `skills/`: la graduación activaría la
propagación de flota vía el reconciliador, que el operador prohibió. Consecuencia:
la copia runtime (`~/.agents/skills/ci-pattern/`) NO la gestiona
`testing/suites/refresh-personal-symlinks/refresh-personal-symlinks.sh` (no tiene
fila en `managed-skills.tsv`) y se desfasa silenciosamente con cada cambio
canónico. Esa copia se instala y actualiza con el instalador dedicado del repo
canónico, idempotente, con la raíz resuelta desde su propia ubicación:

```bash
bash bin/install-ci-pattern-cli.sh            # instala/actualiza la copia runtime
bash bin/install-ci-pattern-cli.sh --dry-run  # previsualiza sin mutar (rsync -n)
bash bin/install-ci-pattern-cli.sh --check    # drift check: exit 0 limpio, exit 1 + lista de diferencias
bash bin/install-ci-pattern-cli.sh --force    # autoriza borrados no vacíos detectados
```

**Borrados.** El instalador previsualiza los borrados previstos con
`rsync -n` antes de cualquier escritura y exige `--force` para
autorizarlos. La previsualización es de sólo lectura y aparece en la
salida **antes** de la línea de resumen post-escritura, de modo que un
install con `--force` exitoso no obliga a aceptar borrados a ciegas:
el operador ve la lista primero y el resumen después. Sin `--force`,
el instalador rehúsa con exit 1 y deja el runtime intacto. En
`--check` y `--dry-run` no se muta nada: exit 1 + lista si hay drift,
exit 0 si está limpio. **Revise la lista de borrados antes de aceptar
`--force`; no lo invoque a ciegas.**

**Preflight del escritor (obligatorio):** tras cualquier cambio canónico de la
skill, ejecute `bash bin/install-ci-pattern-cli.sh` (instala la copia runtime) y
`bash bin/install-ci-pattern-cli.sh --check` (debe salir 0) antes de dar el cambio
por hecho. Es preflight del escritor, no gate de CI: la superficie de drift es la
copia runtime local (`~/.agents/skills/ci-pattern/`), estado de máquina que CI no
observa — cablear `--check` como gate de CI ahí sería un verde ficticio (síntoma
«gate de CI sobre estado local»; fix: el preflight del escritor, §6 de la skill).
El bytecode generado (`__pycache__/`) queda excluido en ambos lados. La suite de
comportamiento del instalador vive en `testing/suites/install-ci-pattern-cli/`,
cableada en el job `unit` de `.github/workflows/tests.yml`.

## §8 `report-friction` (#198, slice 1: clasificación y saneado)

`report-friction` prepara el reporte de una fricción de la skill detectada en un consumer. Clasificación EXPLÍCITA por `--kind` (`skill-bug` | `rule-unfulfillable` | `doc-drift` | `consumer-local`): `skill-bug`, `rule-unfulfillable` y `doc-drift` tienen destino `team-skills`; `consumer-local` se queda en el tracker del consumer. Nunca se deduce del texto.

Campos: `--gate` (o `--hr`), `--skill-version`, `--skill-sha`, `--command`, `--observed`, `--expected`, `--repro`; opcional `--redact <término>` (repetible). Por defecto escribe el cuerpo saneado con `--dry-run <file>` sin tocar la red; `--publish` llega con el slice 2 (deduplicación + gh autenticado).

Saneado determinista: tokens (`ghp_*`, `gho_*`, `github_pat_*`, `AKIA*`, `sk-*`), emails con host, hosts de URLs, rutas absolutas POSIX/Windows, **FQDN sueltos** (≥2 puntos con última etiqueta alfabética; los nombres de fichero con extensión conocida sobreviven), términos `--redact` (repetible, case-insensitive; un flag sin valor es error de uso, exit 2) y bloques de código del consumer (`<consumer-code omitted>`). TODO campo libre que entra al cuerpo se sanea (`--gate`, `--hr`, `--skill-version`, `--skill-sha`, `--command`, `--observed`, `--expected`, `--repro`). Huella de deduplicación estable: `sha256(kind|gate|firma normalizada)` — rutas y dígitos fuera, minúsculas; dos entornos con el mismo defecto comparten huella. Marcador en el cuerpo: `<!-- ci-pattern-friction: <huella> -->`. El cuerpo sigue el formulario `engineering.yml`. El texto depende del destino: `consumer-local` se redacta para el tracker del consumer; el resto para esta skill.

(Publicación y deduplicación por huella: contrato en el slice 2 de #198 — PR #261/#264.)
