# docs-codeguide — gate fail-closed del codeguide de un repo

Asset ejecutable de `ci-pattern` (issue DysTelefonica/team-skills#528,
sección A2). Sustituye al camino manual que comprueba una sola raíz
(`internal/components/`) por una verificación configurable en los dos
sentidos: el codeguide cita solo rutas que existen, y cada módulo del
code-root está nombrado en alguna página.

Binario (exit 0/1/2), stdlib exclusivamente, mismo script en local y CI
(TS#517 AC2). El veredicto es fail-closed: una entrada vacía en
cualquier lado es un FAIL — el comentario de la issue lo prohíbe:
«mejor que el espejo gentle-ai», que solo cubría `internal/components/`.

## Contenido

| Fichero | Papel |
|---|---|
| `check_docs_codeguide.py` | Gate: extrae las rutas del codeguide, enumera los módulos del code-root y emite el veredicto fail-closed |
| `tests/test_check_docs_codeguide.py` | Suite ejecutable (solo stdlib): 49 casos con path verde y path de fallo en cada sentido, en el fail-closed, en los bloques cercados, en las varias raíces y en globs, marcadores y llaves |

## Sentidos de la verificación

| Sentido | Origen | Destino | Verifica |
|---|---|---|---|
| Adelante | `docs/codebase/**/*.md` (default `--docs-path`) | repo | Cada ruta citada existe como fichero o directorio |
| Atrás | `--code-root` (default `internal/components`) | `docs/codebase/**/*.md` | Cada módulo de primer nivel está nombrado por su nombre, completo y en contexto de código (backticks, bloque cercado) o como segmento de una ruta |

Un módulo cuenta si:

- Es un fichero cuyo nombre casa con `--code-glob` (módulos planos,
  estilo `src/modules/users.py`); el nombre que se reporta es el
  `stem`.
- O es un directorio que contiene al menos un fichero que case con
  `--code-glob` (módulos carpeta, estilo
  `internal/components/<name>/*.go`); el nombre que se reporta es el
  nombre del directorio.

## Variante público vs Privado Free

| Pieza | Público | Privado Free |
|---|---|---|
| `check_docs_codeguide.py` | Idéntico | Idéntico |
| `tests/test_check_docs_codeguide.py` | Idéntico | Idéntico |

El script no usa la API de GitHub ni secretos: recorre `docs/codebase/**`
y el code-root con `pathlib`, extrae las rutas con un regex sobre
Markdown y compara contra el árbol del repo en disco. Sin red, sin
tokens, sin estado externo.

## Destino en el consumer

Copie el directorio a la ruta equivalente del repo consumer (por
ejemplo `scripts/docs-codeguide/`) y versione allí los argumentos.
No hay nada más que instalar: el gate es stdlib-only.

## Cableado del job

```yaml
docs-codeguide:
  runs-on: ubuntu-latest
  steps:
    - uses: actions/checkout@v4
    - name: Verificar sincronia codeguide <-> codigo
      run: |
        python3 scripts/docs-codeguide/check_docs_codeguide.py . \\
          --docs-path docs/codebase \\
          --code-root internal/components --code-glob "*.go"
```

El job tiene que ser un check **requerido**: declárelo en la lista de la
política del agregador (`required-jobs`) y en su `needs`. La fase 5 de
`ci-pattern adoption check` (HR-63, `docs_map.job`) falla si el workflow no
define el job o si `.github/required-jobs-policy.json` no lo lista.

```json
{ "required_jobs": ["unit", "docs-codeguide"] }
```

```yaml
required:
  needs: [unit, docs-codeguide]
```

Salida esperada en path verde (exit 0):

```text
docs-codeguide: OK (9 pages, code-root=internal/components, glob=*.go)
```

Salida en path rojo (exit 1), cada ausente en su propia línea de
`stderr`:

```text
FAIL docs -> repo: missing internal/components/ghost/foo.go
FAIL repo -> docs: missing orphan_module
```

Salida en path inválido (exit 2), el path del repo no existe:

```text
FAIL input: repo path is not a directory: /path/that/does/not/exist
```

## Contrato del veredicto

| Clave del dict | Significado |
|---|---|
| `missing_in_repo` | Rutas citadas que NO existen; lista vacía = sentido adelante OK |
| `missing_in_codeguide` | Módulos del code-root que NO aparecen en el codeguide; lista vacía = sentido atrás OK |

Si cualquiera de las dos listas tiene entradas, el veredicto es
`exit 1`. Si el path del repo no es un directorio, `exit 2`.

El dict es serializable a JSON sin transformaciones: reporteros y
herramientas externas pueden consumirlo directamente. La entrada
`<empty docs/codebase/>` o `<empty internal/components/>` nombra el
lado vacío cuando la entrada está sin contenido — el fail-closed
prevé este caso explícitamente.

## Extracción de rutas

`extract_path_refs` reconoce cadenas entre comillas invertidas que
contengan al menos una barra y empiecen por una letra (el patrón del
codeguide estilo Alan: `internal/components/engram/`). Las tablas
Markdown `` | `path` | ... | `` se extraen igual que en prosa; una
misma ruta citada varias veces en una página colapsa a una sola
entrada.

Los bloques de código cercados (tres o más comillas invertidas o `~~~`) no
cuentan como citas: se descartan enteros antes de extraer, así que un ejemplo
con barras no se toma por ruta ni invierte el emparejamiento del código en
línea posterior (issue #725). Un bloque sin cerrar llega hasta el final de la
página.

### Globs, marcadores y llaves (issue #737)

Una cita es válida si cumple una de estas reglas:

| Cita | Válida si | Rojo, y lo que nombra |
|---|---|---|
| Ruta literal (`app/src/main.py`) | existe como fichero o directorio | la ruta |
| Glob (`scripts/check_*.py`, `*`, `?`, `[…]`) | casa con al menos un fichero o directorio | `cita (glob sin coincidencias)` |
| Segmento marcador (`app/src/modules/<app>/`) | existe el prefijo anterior al primer marcador (puede llevar glob) | `cita (prefijo inexistente: <prefijo>)` |
| Alternancia (`x/{a,b}.py`) | es válida cada expansión | `cita (<expansión>: <motivo>)` |

Lo que sigue a un marcador no se valida: el gate no sabe qué valor toma
`<app>`. Los comandos con espacios siguen siendo citas a revisar.

Quedan fuera: URLs (`http://`, `https://`, `mailto:`, `ftp://`,
`file://`), tokens entre comillas sin barra (comandos, flags) y
enlaces Markdown `[texto](destino.md)` (los enlaces relativos a otras
páginas del codeguide no son rutas de repo).

## Configuración por CLI

| Argumento | Default | Efecto |
|---|---|---|
| `repo` (posicional) | obligatorio | Ruta al directorio raíz del repo |
| `--docs-path` | `docs/codebase` | Ruta al codeguide |
| `--code-root` | `internal/components` | Ruta a los módulos de primer nivel; repetible |
| `--code-glob` | `*.go` | Glob que debe casar un módulo para contar; repetible |

`--code-glob` se aplica dos veces: ficheros a primer nivel del
code-root (módulos planos) y ficheros dentro de cada subdirectorio
del code-root (módulos carpeta).

### Varias raíces (#727)

Repita `--code-root` y `--code-glob` por pares, en el mismo orden:

```bash
python3 check_docs_codeguide.py . \
  --code-root app/modules --code-glob "*.py" \
  --code-root app/templates --code-glob "*.html"
```

El sentido adelante (docs → repo) se evalúa una sola vez; el atrás, por raíz.
Con más de una raíz cada módulo ausente se nombra `módulo (raíz)`; con una
sola queda el nombre desnudo. Un número distinto de `--code-root` y
`--code-glob` es exit 2. En código, `check_docs_codeguide_roots(repo, [(raíz,
glob), ...])`. En el contrato de adopción, `docs_map.code_roots` lista
`{root, glob}` y las claves escalares `code_root` y `code_glob` siguen valiendo
para una sola raíz; mezclar las dos formas falla cerrado.

## Límites de la verificación

- **Ruta citada que no es del repo.** Las URLs se ignoran; pero una
  cita como `` `api/spec.yaml` `` en un repo sin `api/` aparece en
  `missing_in_repo` y exige revisar la cita.
- **Rename de módulo.** El extractor del sentido atrás busca el
  nombre como cadena literal; un rename de `app` a `apphandler`
  exige actualizar las páginas que lo citan, mientras tanto el gate
  lo nombra en `missing_in_codeguide`.
- **Code-root vacío.** El veredicto es
  `missing_in_codeguide == ["<empty <code-root>/>"]` y exit 1: el
  caso «el repo aún no tiene el árbol» no se disfraza como OK.

## Tests

```bash
for t in scripts/docs-codeguide/tests/test_*.py; do python3 "$t"; done
```

Bloques cubiertos: extracción de rutas (vacía, backtick, tabla,
no-ruta, unicidad, URL externa); sentido adelante (codeguide limpio,
ruta inexistente, múltiples); sentido atrás (módulo cubierto,
huérfano); fail-closed (`docs/codebase/` y code-root vacíos, path
inválido); configuración (`--docs-path`, `--code-root`, `--code-glob`);
contrato del binario (`exit 0/1/2`); contrato del dict (claves
estables y serializable a JSON).