# issue-spec — gate de trazabilidad del cuerpo del PR

Asset canónico de `ci-pattern` (DysTelefonica/team-skills#633). Va a sustituir
las copias locales de cada consumer (`check_issue_spec.py` en BP,
`check_issue_specs.py` en APAP_WEB) por una sola fuente.

## Uso local (preflight)

```bash
python3 personal/ardelperal/ci-pattern/assets/issue-spec/check_issue_spec.py \
  local --body-file <ruta|-> [--repository owner/repo]
```

Sin payload, sin token y sin red. Exit codes: `0` sin hallazgos · `1` hallazgos ·
`2` entrada inválida. `--repository` activa la regla de repositorio; sin él,
los enlaces `owner/repo#N` no se juzgan.

## Reglas de forma

1. El cuerpo enlaza una issue con `Closes|Fixes|Resolves|Refs #N` fuera de
   código. Sin ella, el mensaje manda a `depends-on` de Chain Context: una PR de
   la que se depende no se cita aquí.
2. El cuerpo no cierra y referencia la misma issue a la vez.
3. Una referencia `owner/repo#N` apunta al repositorio del PR.

Las palabras de cierre valen en prosa, como en GitHub. Dentro de un span de
código o de un bloque cercado no cuentan. Un bloque con sangría de 4 espacios
no se quita: se cuenta, así que el fallo es ruidoso y no silencioso.

## Tests

```bash
python3 personal/ardelperal/ci-pattern/assets/issue-spec/tests/test_check_issue_spec.py
```

La suite la ejecuta `testing/run-unit-suites.sh`.
