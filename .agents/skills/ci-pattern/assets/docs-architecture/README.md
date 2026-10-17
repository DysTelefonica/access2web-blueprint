# docs-architecture — el documento de arquitectura no puede mentir

Asset de `ci-pattern` (HR-68, DysTelefonica/team-skills#714). `check_docs_architecture.py`
comprueba `docs/architecture.md` contra el código. Stdlib, sin red, mismo script en
local y CI. Exit 0 limpio, 1 hallazgos (uno por línea de `stderr`), 2 entrada inválida.

## Formato del documento

Dos piezas, ambas obligatorias:

1. **Árbol de módulos**: una tabla bajo un título «Árbol de módulos» cuya primera
   columna es la ruta del módulo entre comillas invertidas, relativa al repo
   (`app/routes`, `web/page.html`), y cuya segunda es su responsabilidad.
2. **Capas y dependencias**: un bloque cercado con información `architecture`,
   JSON. `layers` lista cada capa con sus módulos; `allowed` dice de qué capas
   puede depender cada una. Lo que no se permite está prohibido.

````markdown
## Árbol de módulos

| Módulo | Responsabilidad |
|---|---|
| `app/routes` | Entrada HTTP |
| `app/domain` | Reglas de negocio |

## Capas y dependencias

```architecture
{"layers": [{"name": "web", "modules": ["app/routes"]},
            {"name": "core", "modules": ["app/domain"]}],
 "allowed": {"web": ["core"], "core": []}}
```
````

## Capas por patrón de ruta (bloque v2, #751)

Cuando una capa atraviesa módulos (en un slice hexagonal, `domain/`, `application/` y `adapters/`
viven dentro del mismo módulo), cada capa declara `paths` en lugar de `modules`:

````markdown
```architecture
{"layers": [{"name": "domain", "paths": ["app/modules/*/domain/**"], "allowed": ["domain"]},
            {"name": "delivery", "paths": ["app/modules/**"], "allowed": ["domain", "delivery"]}],
 "exempt": ["app/__init__.py", "app/static/**"]}
```
````

| Regla | Detalle |
|---|---|
| Patrones | `*` no cruza `/`, `**` cruza directorios, `?` es un carácter; rutas relativas al repo |
| Primera coincidencia | decide la capa el primer patrón que casa, en el orden de las capas; `exempt` gana a las capas |
| Cobertura | todo fichero fuente de los code-roots cae en una capa o en `exempt`; uno sin capa es rojo y se nombra |
| Patrón vacío | un patrón (o de `exempt`) que no casa ningún fichero es rojo |
| `allowed` | dentro de cada capa o como mapa de primer nivel; lo que no se permite está prohibido |
| Expansión | para los linters, los patrones se expanden a los paquetes y módulos concretos: el directorio entero cuando todos sus ficheros son de la capa, los ficheros si no (los exentos no cuentan). Un slice mixto genera contratos sobre sus subpaquetes y sus ficheros, sin falsos positivos entre capas del mismo módulo |
| Marcas | además del fence `architecture`, se acepta el primer bloque entre `<!-- architecture:begin -->` y `<!-- architecture:end -->` (fence `json`) |
| Árbol | las filas pueden terminar en `/`; un módulo está descrito si tiene fila él o alguno de sus descendientes; una fila describe un camino que existe |

`independent`, `pure` y `tolerated` (#749) siguen igual y se escriben con los nombres del canon:
un consumer con otros nombres (`pure_layers` por `pure`, `forbidden_external` por
`forbidden_externals`, `baseline` por `tolerated`) los adapta al adoptar HR-68. Las capas por
`modules` (v1) siguen funcionando; no se mezclan `modules` y `paths` en una misma capa.

## Qué comprueba

| Sentido | Falla cuando |
|---|---|
| árbol → código | un módulo descrito no es un módulo de los code-roots |
| código → árbol | un módulo de los code-roots no está descrito |
| capas | un módulo del árbol está en ninguna o en varias capas; una capa declara un módulo fuera del árbol; una capa no tiene regla; una regla nombra una capa no declarada |
| fail-closed | falta el documento, el árbol está vacío, un code-root no tiene módulos, falta el bloque o no es JSON |

Un módulo es un fichero que casa el glob o un directorio con al menos un fichero que lo
casa, igual que en `docs-codeguide`; `__init__.py` no cuenta. Los bloques cercados (comillas invertidas o tildes)
no cuentan como tabla: el árbol de ejemplo dentro de un bloque no vale.

## Reglas de imports (Python y TypeScript)

El bloque `architecture` es la única fuente. De las capas y de `allowed` se deriva la
configuración de [import-linter](https://import-linter.readthedocs.io): un contrato
`forbidden` por cada par de capas distintas cuya dependencia no se permite. Con
`--linter-config` (en el contrato: `stack: python` y `linter_config`) el gate falla si:

| Caso | Hallazgo |
|---|---|
| la configuración versionada no coincide con la derivada del bloque | nombra el fichero; regenere con `--write-config` |
| falta la configuración | nombra el fichero |
| `lint-imports` no está instalado o no se ejecuta | fail-closed (HR-3), nunca pasa en silencio |
| un import viola una regla | `fichero:línea` y la regla incumplida |

Los módulos de una capa se nombran como rutas (`app/routes`) y se traducen a nombres
con puntos (`app.routes`); los imports dentro de una misma capa siempre se permiten.
La versión del linter se fija en el CI (`import-linter==2.15`).

Con `--stack typescript` (en el contrato, `stack: typescript`) el mismo bloque genera la
configuración de [dependency-cruiser](https://github.com/sverweij/dependency-cruiser)
(una regla `forbidden` por par no permitido, más `no-unresolvable-local-imports`, para
que un import local roto no pase por verde). Los casos son los de la tabla anterior;
además, un linter que no analiza ningún módulo es rojo (HR-3). El `fichero:línea` sale
de buscar el especificador del import en el fichero de origen. El CI fija
`dependency-cruiser@17.4.3` y `typescript@7.0.2`.

## Uso

```bash
python3 check_docs_architecture.py . --doc docs/architecture.md \
  --code-root app --code-glob "*.py" --code-root web --code-glob "*.html"
```

`--linter-config` valida los imports y `--write-config` escribe esa configuración desde el
bloque. `--code-root` y `--code-glob` van por pares, como en `docs-codeguide`. En el contrato de
adopción lo ejecuta la fase 5 cuando el consumer declara `architecture` (HR-68).

## Reglas adicionales del bloque (#749, solo Python)

Tres claves opcionales del bloque `architecture` permiten sustituir un verificador de capas
propio sin perder reglas. Sin ellas la configuración derivada no cambia.

| Clave | Forma | Qué se deriva |
|---|---|---|
| `independent` | `[{"modules": ["app/lan", "app/exp"]}]` | Un contrato `forbidden` por par ordenado del grupo, con `allow_indirect_imports = True` (importaciones directas): ningún módulo del grupo importa a otro. Lo que no está en el grupo (módulos transversales) no se restringe. |
| `pure` | `{"layers": ["domain"], "forbidden_externals": ["fastapi"]}` | Un recorrido AST (stdlib) de los módulos de esas capas: un `import` o `from … import` de un paquete externo prohibido es `fichero:línea: la capa pura «domain» importa el framework «fastapi»`, también bajo `TYPE_CHECKING`. Un fichero que no se puede analizar es rojo (HR-3). No exige tener instalados los paquetes. |
| `tolerated` | `[{"from": "delivery", "to": "di", "imports": ["origen -> destino"]}]` | `ignore_imports` del contrato de ese par. Una importación no listada sigue siendo violación y una entrada que ya no coincide con ninguna es `STALE en tolerated` (rojo): la deuda solo puede encoger. La lista enumerada es el recuento. |

`tolerated` solo admite pares cuya dependencia no está permitida en `allowed`. Los módulos de
`independent`, `pure` y `tolerated` se escriben con barras, como el resto del bloque. Los módulos
de las capas han de ser paquetes importables: un directorio sin `__init__.py` hace fallar a
import-linter con «Module … does not exist».

## Tests

```bash
python3 tests/test_check_docs_architecture.py
```
