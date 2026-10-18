type: patch

`ci-pattern adoption generate-doc` deriva la forma del nombre de rama del documento operativo (sección `branch` y etapa 3 del ciclo de vida) de P01 (`branch_name_pattern`) en vez de un texto fijo. Un consumer que exige el número de issue (`<tipo>/<número>-<slug>`) obtiene la frase que le corresponde; sin P01 declarado, o con el default `<tipo>/<slug>`, el documento generado no cambia byte a byte (#842).
