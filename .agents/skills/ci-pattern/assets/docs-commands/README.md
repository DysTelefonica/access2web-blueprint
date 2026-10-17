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

Activación: `docs_commands` en el contrato de adopción. Sin declaración, `adoption
check --phase 5` avisa por `stderr` y sale 0. La fase 5 solo valida la declaración
y el job; los comandos los ejecuta el job, nunca la CLI del patrón como subproceso.

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
