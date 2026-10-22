type: patch

HR-68 (#748): el adaptador de imports de Python lee las violaciones de `lint-imports` aunque el linter parta su salida al ancho de la terminal. Se ejecuta con `COLUMNS` ancho y el análisis une la continuación de una viñeta partida (tras `->`, antes de `(l.N)` o en mitad de una ruta), de modo que la violación sale como `fichero:línea: import prohibido por la regla «…»` y no como «sin violaciones legibles».
