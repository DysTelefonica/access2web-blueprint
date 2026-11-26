type: minor

**Qué.** El gate `e2e-evidence` lee los globs de UI del `ci-pattern.yaml` del consumer (`P59_e2e_ui_globs`).
**Causa.** La UI de cadete vive en `www/**`, que no está en `DEFAULT_UI_GLOBS`. La única forma de cambiarla era `--ui-glob` en la línea de comandos, así que la lista no formaba parte del contrato del consumer.
**Arreglo.** Se añade el parámetro P59 al esquema, a `parameters.md` y al ejemplo. El gate lo lee con el mismo lector estructural que el gate de imágenes. `--ui-glob` manda sobre él, y sin el fichero o sin la clave rigen los defaults de siempre. Si no es una lista no vacía de cadenas, el gate sale con 2. Tests: P59 sustituye a los defaults (en ambos sentidos), la línea de comandos manda, y un P59 inválido sale con 2.
**A partir de ahora.** Un consumer cuya UI no esté en los defaults la declara en `P59_e2e_ui_globs` antes de cablear el gate. Hoy ningún consumer lo cablea. Fase 0 de CI/CD, punto 0.6.
