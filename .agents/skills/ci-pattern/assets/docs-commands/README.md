# docs-commands — comandos documentados, extraídos y probados

Asset de `ci-pattern` (HR-65, DysTelefonica/team-skills#711). Generaliza la prueba
de invocaciones documentadas de gentle-ai: `check_docs_commands.py` (stdlib, sin
red) **extrae** de `README.md`, `AGENTS.md` y `docs/**` los comandos que empiezan por
el nombre de la CLI del consumer (`cli_name`, parámetro) y los prueba; añadir un
comando a un documento lo añade al corpus sin tocar ningún test.

| Nivel | Qué es | Qué se prueba |
|---|---|---|
| `executed` | Verbo de `executed_verbs`, sin placeholders ni shell compuesto | Se ejecuta en un directorio temporal; el parser no puede rechazarlo |
| `parse-only` | Lleva placeholders (`<valor>`, se rellenan con un valor ficticio) o verbo no seguro | Se ejecuta con `dry_run_flag`; el parser no puede rechazarlo |
| `presence-only` | Shell compuesto, o sin `dry_run_flag` declarado | Solo se cuenta, para que el nivel sea visible |

Un comando rechazado (`rejection_exit_codes`, por defecto 2) es rojo con
`fichero:línea`. Extraer cero comandos falla. `exclude` exime comandos
ilustrativos y solo puede encoger (una entrada obsoleta es `STALE`).

## Tres causas, tres hallazgos

El veredicto distingue **por qué** falló el gate, porque arreglarlo es distinto en
cada caso (#1021). Los tres son rojos: el fallo de entorno también para el pipeline,
lo que cambia es el diagnóstico, no la severidad.

| Hallazgo | Qué pasó | Qué hacer |
|---|---|---|
| `rechazado por el parser (exit N)` | La CLI recibió el comando documentado y lo rechazó | Corregir el comando en `fichero:línea` o quitarlo de la documentación |
| `timeout: no terminó en Ns` | El runner cortó la ejecución; el parser no llegó a rechazarlo | Sacar el comando del corpus con `exclude` o hacerlo más rápido; si el comando está mal, corregirlo |
| `entorno: la CLI «X» no se puede ejecutar aquí` | La CLI no está donde corre el job | Instalar la CLI en la imagen del runner |

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
