type: minor

HR-82 (#1382): las tablas de checks ESCRITAS A MANO se miden contra los jobs reales de los workflows. La fase 5 ya medía la tabla del documento operativo generado, pero una tabla escrita a mano es tan vinculante como esa y nadie la miraba: en `DysTelefonica/cadete` la tabla de `docs/codebase/testing-and-ci.md` atribuía `php-lint` a `ci.yml` (el job vive en `php-lint-push.yml`), omitía `ci-images` (que sí es un job requerido) y una auditoría le añadió un job que no define ningún workflow, con la fase 5 en verde.

La capa nueva se activa por la declaración `checks_tables` del contrato de adopción (`tables`: lista de `{path, workflow}` y `section` opcional). Por cada fila declarada contrasta el id del job y el nombre del check publicado contra los jobs que el workflow declarado define, leídos con el lector estructural del canon. Da tres rojos distintos porque el remedio es distinto: un job que ningún workflow define (fabricado), un job que vive en otro workflow (el rojo nombra el workflow que lo define) y un job requerido que no aparece en ninguna tabla declarada. Sin la declaración, `adoption check --phase 5` avisa por `stderr` y sale 0: ningún consumer cambia de comportamiento hasta que la declara.

Tests: `assets/adoption/tests/test_checks_tables.py` (el caso de cadete: job inventado, job movido de workflow y job requerido omitido).
