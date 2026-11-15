# docs-commands — comandos documentados, extraídos y probados

Asset de `ci-pattern` (HR-65, DysTelefonica/team-skills#711). Generaliza la prueba
de invocaciones documentadas de gentle-ai: `check_docs_commands.py` (stdlib, sin
red) **extrae** de `README.md`, `AGENTS.md` y `docs/**` los comandos que empiezan por
el nombre de la CLI del consumer (`cli_name`, parámetro) y los prueba; añadir un
comando a un documento lo añade al corpus sin tocar ningún test. Con la
declaración `existence` el gate comprueba además que lo que el comando documentado
invoca existe (ver «La capa de existencia»).

| Nivel | Qué es | Qué se prueba |
|---|---|---|
| `executed` | Verbo de `executed_verbs`, sin placeholders ni shell compuesto | Se ejecuta en un directorio temporal; el parser no puede rechazarlo |
| `parse-only` | Lleva placeholders (`<valor>`, se rellenan con un valor ficticio) o verbo no seguro | Se ejecuta con `dry_run_flag`; el parser no puede rechazarlo |
| `presence-only` | Shell compuesto, o sin `dry_run_flag` declarado | Solo se cuenta, para que el nivel sea visible |

Un comando rechazado (`rejection_exit_codes`, por defecto 2) es rojo con
`fichero:línea`. Extraer cero comandos falla. `exclude` exime comandos
ilustrativos y solo puede encoger (una entrada obsoleta es `STALE`).

## La capa de existencia: el comando tiene que existir

Los tres niveles de arriba prueban la **forma** del comando: que el parser real lo
acepte. Con `dry_run_flag --help` eso es todo lo que se mide, y no basta: el caso vivo
de `DysTelefonica/access2web-blueprint` (`docs/adopcion-fase-2-aislamiento.md:41`)
cita `make test` con `--cov` y da el `Makefile` como fuente, pero el target `test` dejó
de medir cobertura (#1374). El comando se validaba en su forma (`pytest --help` sale 0
sin mirar el resto) y el gate no veía que lo que invocaba ya no estaba.

La capa de existencia comprueba que **lo que el comando invoca existe**, sin ejecutar
nada. La regla es estrecha y determinista —solo mira ficheros, Makefile y `PATH`— y la
activa la declaración `existence` del contrato, para que un consumer que no la declara
no cambie de comportamiento.

| Cita del documento | Qué se comprueba |
|---|---|
| `./scripts/x.sh` | El script existe en el repositorio, resuelto desde el directorio del documento que lo cita **o** desde la raíz del repo |
| `make <target>` | El target existe en el Makefile, y los flags que la documentación le atribuye están en su receta |
| `<ejecutable>` | El ejecutable que `path_executables` declara permitido resuelve en el `PATH` del runner |

Una cita con `./` o `../` puede ser relativa a dos sitios, y vale si el fichero existe
en **cualquiera** de los dos:

- **el documento que la escribe**: es la convención de los enlaces markdown, y el gate
  de `docs-links` ya resuelve así. En `DysTelefonica/access2web-blueprint` la capa daba
  234 rojos falsos sin esta vía: `docs/03-aplicaciones/brass/prd/capacidades.md` cita
  `../capacidades.md`, que vive en `docs/03-aplicaciones/brass/` (#1385);
- **la raíz del repo**: un comando entre comillas invertidas se ejecuta desde la raíz,
  no desde la carpeta del documento que lo muestra. `docs/guia.md` que dice «ejecute
  `./scripts/check.sh`» apunta a `scripts/check.sh` (#1385, seguimiento).

Tres casos, todos deterministas:

- el fichero existe resuelto por una de las dos vías → nada que decir;
- no existe por ninguna → rojo, nombrando las rutas que el gate ha mirado;
- sale del repositorio por las dos vías → rojo, con su propio motivo: un script de fuera
  del repo no es un script del repo. Si una sola vía sale del repo, la otra decide.

Una cita **sin** `./` ni `../` no se resuelve contra el repo: sigue siendo un
ejecutable del `PATH` que la política declara permitido, o no se mide. Un documento
en la raíz del repo tiene por directorio la raíz, así que ahí la cita se resuelve
como antes.

La declaración es `existence` en el bloque `docs_commands`: `makefile` es el Makefile
que se consulta (por defecto `Makefile`; un `make` con `-C` o `-f` no se mide, porque
consulta otro) y `path_executables` son los ejecutables del `PATH` que la documentación
puede citar. Sin esta lista el gate no mira el `PATH`, porque un ejecutable ausente en
el runner es un problema del runner, no de la documentación (el mismo criterio que el
hallazgo de entorno de arriba). La atribución de flags es de la misma línea y solo con
la CLI declarada, y una línea que cita dos targets no se cruza con ninguno (es
ambigua): un target de patrón (`%`) o un comando sin target no se comprueban. La capa
no depende de que la CLI esté donde corre el job —no ejecuta nada— y `exclude` exime
sus citas igual que exime un comando de ejecución.

## Tres causas, tres hallazgos

El veredicto distingue **por qué** falló el gate, porque arreglarlo es distinto en
cada caso (#1021). Los dos primeros son rojos para el pipeline. El de entorno también
sale rojo para el pipeline **salvo** que la política no declare ninguna forma de
ejecutar (#1038): entonces baja a aviso, porque no hay medición que la CLI ausente
pueda ocultar. Lo que cambia en los tres casos es el diagnóstico, y en el tercero
también la severidad, pero solo cuando el gate no iba a ejecutar nada.

| Hallazgo | Qué pasó | Qué hacer | Severidad |
|---|---|---|---|
| `rechazado por el parser (exit N)` | La CLI recibió el comando documentado y lo rechazó | Corregir el comando en `fichero:línea` o quitarlo de la documentación | rojo |
| `timeout: no terminó en Ns` | El runner cortó la ejecución; el parser no llegó a rechazarlo | Sacar el comando del corpus con `exclude` o hacerlo más rápido; si el comando está mal, corregirlo | rojo |
| `entorno: la CLI «X» no se puede ejecutar aquí` | La CLI no está donde corre el job | Instalar la CLI en la imagen del runner | rojo, o **aviso** si la política no declara `dry_run_flag` ni `executed_verbs` |

Sin `dry_run_flag` ni verbos en `executed_verbs`, `_tier` devuelve `presence-only`
para cualquier comando y el gate no lanza nada. En ese caso la ausencia de la CLI se
reporta igual, pero a `stderr` con prefijo `AVISO` y el job sale 0. Basta con que se
declare **una** de las dos claves para que vuelva a bloquear. Es el caso de
APAP_WEB#2324: 58 comandos que su runner nunca iba a ejecutar porque no tiene `make`.

El tercero sale **una vez**, no una por comando documentado: la disponibilidad de la
CLI se decide una sola vez por gate. Antes un único `except` cubría las tres causas y
devolvía `no se pudo ejecutar` para todas, así que un runner sin `make` repetía la
misma línea `[Errno 2]` por cada comando documentado (55 en
`DysTelefonica/APAP_WEB`) y el rojo señalaba «tu documentación está mal» sobre un
problema del runner.

Activación: `docs_commands` en el contrato de adopción. Sin declaración, `adoption
check --phase 5` avisa por `stderr` y sale 0. La fase 5 solo valida la declaración
y el job; los comandos los ejecuta el job, nunca la CLI del patrón como subproceso.

**La fuente única es ese bloque del contrato.** El job lo lee de ahí:

```bash
run_docs_commands.py <repo> [--contract .github/ci-pattern-adoption.json]
```

`--config <fichero>` sigue aceptando el JSON suelto de la raíz que el job usaba
antes (#1020), pero avisa por `stderr` de que está obsoleto y no cambia de
alcance: mientras exista, la fase 5 nombra el fichero de raíz si lo encuentra,
porque dos declaraciones del mismo dato derivan sin que nada las compare.
Ejecutar el job sin `--config` con el fichero viejo en la raíz es seguro: el
fichero queda muerto, no se lee.

## El runner necesita la CLI

El job ejecuta la CLI del consumer, así que `cli_command[0]` tiene que existir en la
imagen del runner: como fichero del repo, o resuelto por `PATH`. Lo que no es
fichero del repo lo busca el runner, no el contrato, y el contrato no lo declara.
Sin ella el job no mide nada, y en vez de callarse lo dice nombrando la CLI que
falta.

Si el job corre en un runner distinto del de desarrollo (self-hosted frente a
GitHub-hosted), la CLI se instala en la imagen de **aquel**: los dos runners no
garantizan la misma disponibilidad.

## Cableado del job requerido

```yaml
docs-commands:
  runs-on: ubuntu-latest
  steps:
    - uses: actions/checkout@v4
    - name: Comandos documentados
      run: python3 scripts/docs-commands/run_docs_commands.py .
```

```json
{ "required_jobs": ["unit", "docs-commands"] }
```
