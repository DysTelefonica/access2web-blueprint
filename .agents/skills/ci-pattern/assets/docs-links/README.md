# docs-links — enlaces locales de la documentación

Asset de `ci-pattern` (HR-66, DysTelefonica/team-skills#712). `check_docs_links.py`
(stdlib, sin red) recorre `README.md`, `AGENTS.md` y `docs/**/*.md`, resuelve cada
enlace local contra el directorio de su fichero y falla nombrando `fichero:línea`
y destino. No cuentan URLs con esquema, `mailto:`, anclas de la misma página ni
bloques cercados. Un destino Markdown con `#ancla` exige ese título.

Además exige navegación (cada spoke enlaza al hub y el hub a cada spoke), que los
ficheros y frases retirados no reaparezcan y que `baseline` (deuda heredada) solo
encoja: una entrada que ya no coincide con un enlace roto es `STALE`.
Recorrer cero ficheros es un fallo.

Activación: `docs_links` en el contrato de adopción (claves en el esquema); sin
declaración, `adoption check --phase 5` avisa por `stderr` y sale 0.

## Alcanzabilidad: todo `.md` del alcance tiene que ser alcanzable (HR-80, #1379)

Un enlace que resuelve no dice si la documentación *completa* se ve. Un `.md` al
que nadie enlaza no rompe ningún enlace, así que ningún gate lo cazaba: la
auditoría de p3 en `DysTelefonica/cadete` encontró 45 documentos huérfanos
(`docs/skills-catalog.md`, `PRD_ENTITY_MAP.md`…) que un agente no encuentra «con
un solo vistazo».

La regla vive **en este gate**, no en uno nuevo: la clave opcional `reachable` del
bloque `docs_links` enciende la exigencia. La exige el mismo job `docs-links` y la
fase 5 de `adoption check`; no hay job nuevo ni declaración nueva.

```json
"docs_links": {
  "hub": "docs/CODEBASE-GUIDE.md",
  "spokes": "docs/codebase",
  "job": "docs-links",
  "reachable": {
    "from": ["README.md", "docs/CODEBASE-GUIDE.md"],
    "scope": ["docs"],
    "exceptions": ["docs/releases/"]
  }
}
```

| Clave | Qué significa |
|---|---|
| `from` | Puntos de entrada del grafo de enlaces. **Obligatoria**: sin ella la declaración está incompleta y el gate falla nombrando la clave. Cada punto de entrada tiene que existir. |
| `scope` | Rutas (fichero o directorio) donde se exige la completitud. Por defecto, `["docs"]`. |
| `exceptions` | Rutas o globs exentos: `docs/releases/` cubre todo lo que cuelga de él; `docs/adr/*.md` cubre los ficheros de un patrón. |

Qué mide:

- Todo `.md` bajo `scope` tiene que ser **alcanzable** desde `from` siguiendo
  enlaces markdown. El recorrido es **transitivo**: si el hub lleva a A y A lleva a
  B, B está alcanzable. Solo se sigue el enlace de dentro del repositorio; un
  enlace a otro sitio no trae la documentación de este repo.
- Lo que no se alcanza es un **documento huérfano** y sale en rojo con su ruta.
- El `.md` que nadie enlaza desde un `.md` inalcanzable no cuenta: la cadena
  empieza en los puntos de entrada, no en cualquier fichero.
- Un `.md` fuera del alcance no se exige; un fichero que no es Markdown, tampoco.

Es determinista y **falla cerrado** (HR-3): la declaración incompleta o mal formada
—`reachable` sin `from`, `from` que no es una lista de rutas, un punto de entrada
que no existe, un alcance que no existe— sale en rojo nombrando la clave. El rojo
nombra el fichero huérfano y los puntos de entrada desde los que no se alcanza.

La excepción es **explícita y declarativa**: una ruta mal escrita no desactiva
nada en silencio, porque los ficheros que pretendía cubrir salen en rojo como
huérfanos.

Mientras un consumer no declare `reachable`, esta regla no midió nada y el gate
sigue midiendo exactamente lo que medía: ningún consumer cambia de comportamiento
hasta declararla.

**La fuente única es ese bloque del contrato**, igual que en HR-65 (#1020):

```bash
run_docs_links.py <repo> [--contract .github/ci-pattern-adoption.json]
```

`--config <fichero>` sigue aceptando el `docs-links.json` de la raíz que el job
usaba antes, pero avisa por `stderr` de que está obsoleto y no cambia de alcance.
Mientras exista, la fase 5 lo nombra si lo encuentra: dos declaraciones del mismo
dato derivan sin que nada las compare. Igual que en HR-65, dejarlo en la raíz no
rompe el job —pasa a leerse solo el contrato—, pero tampoco debe quedarse.

## Cableado del job requerido

```yaml
docs-links:
  runs-on: ubuntu-latest
  steps:
    - uses: actions/checkout@v4
    - name: Enlaces locales de la documentación
      run: python3 scripts/docs-links/run_docs_links.py .
```

```json
{ "required_jobs": ["unit", "docs-links"] }
```

El job `docs-links` es un check requerido: figura en `required_jobs` de
`.github/required-jobs-policy.json` y en el `needs` del agregador; la fase 5 lo
comprueba cuando `docs_links.job` lo declara.
