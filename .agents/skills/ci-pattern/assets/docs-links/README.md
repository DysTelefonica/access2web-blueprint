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
