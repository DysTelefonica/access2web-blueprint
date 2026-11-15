# checks-table — las tablas de checks escritas a mano, contra los jobs reales

Asset de `ci-pattern` (HR-82, DysTelefonica/team-skills#1382). La fase 5 ya mide la
tabla de checks del **documento operativo generado** (`docs/ci-pattern-flow.md`,
sección «5. Checks requeridos») contra los `required_jobs` de
`.github/required-jobs-policy.json`. Este asset mide las tablas que el consumer
**declara**: las que alguien escribe a mano.

Una tabla escrita a mano es tan vinculante como la generada, y no por eso puede
mentir: el caso vivo es `DysTelefonica/cadete`, cuya tabla de
`docs/codebase/testing-and-ci.md` atribuía `php-lint` a `ci.yml` (el job vive en
`php-lint-push.yml`), omitía `ci-images` (que sí es un job requerido) y una
auditoría le añadió un job que no define ningún workflow. Con la capa declarada, los
tres son rojos.

## Qué mide

| Rojo | Qué pasó | Qué hacer |
|---|---|---|
| job inventado | La fila cita un job que no define ningún workflow de `.github/workflows` | Borra la fila: una receta fabricada es peor que la ausencia |
| job movido | La fila atribuye el job a un workflow que no lo define | Corrige la atribución o saca la fila de esa tabla (el rojo nombra el workflow que lo define) |
| job requerido ausente | Un `required_jobs` de la política no aparece en ninguna tabla declarada | Documenta el job en la tabla del workflow que lo define |

La comparación es con los **workflows reales**: el gate lee
`.github/workflows/*.yml` con el lector estructural del canon (el mismo del gate de
política de workflows) y contrasta, para cada fila, el id del job y el nombre del
check publicado contra los jobs que el workflow declarado define. Un job cuenta
como documentado si el workflow declarado lo define **o** si publica ese nombre de
check.

## La declaración

En el contrato de adopción (`.github/ci-pattern-adoption.json`):

```json
"checks_tables": {
  "tables": [
    {
      "path": "docs/codebase/testing-and-ci.md",
      "workflow": "ci.yml",
      "section": "Los jobs de CI"
    }
  ]
}
```

- `path`: documento con la tabla (obligatorio).
- `workflow`: workflow que define los jobs de la tabla (obligatorio). Se acepta el
  nombre del fichero (`ci.yml`) o su ruta (`.github/workflows/ci.yml`).
- `section`: título de la sección donde está la tabla (opcional). Sin ella se lee el
  documento entero.

La clave es **opcional**: sin ella `adoption check --phase 5` avisa por `stderr` y
sale 0, así que ningún consumer cambia de veredicto por propagar el canon. Con ella,
la fase 5 mide cada tabla declarada.

## Cómo se reconoce una tabla de jobs

Una tabla markdown es una tabla de jobs si su **primera columna** se titula `Job`
(o `Jobs`, `Check`, `Checks`); la segunda columna es el **nombre del check** si se
titula `Check`, `Nombre del check`, `Contexto` o `Context`. Cada fila aporta el id
del job (entre backticks) y, si la hay, el nombre del check publicado. Un bloque de
tabla sin cabecera también cuenta: cada fila con la primera celda entre backticks es
una fila de job. Una tabla declarada sin filas de jobs es roja: una tabla sin jobs
no mide nada.

Solo lectura, sin red: el documento declarado, los workflows del repo y la política
de jobs requeridos. Si un workflow no se puede leer, se dice con su nombre: sin él no
se puede saber qué jobs existen.
