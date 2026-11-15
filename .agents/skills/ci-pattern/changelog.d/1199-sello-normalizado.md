type: minor

El sello de gobierno sella los workflows sobre su contenido NORMALIZADO (DysTelefonica/team-skills#1199). En un fichero bajo `.github/workflows/`, cada `uses: owner/repo@<sha de 40 hex>` y su comentario `# vX` pasan a un marcador fijo antes de hashear: un bump de Dependabot cambia el sha y el comentario, y eso ya no mueve el sello, así que un PR de Dependabot que solo bumpea una action deja de dar un rojo de gobernanza sin defecto detrás. El nombre de la action se conserva (cambiar qué action se usa sigue rompiendo el sello) y todo lo demás del workflow —los pasos, el `run:`, los `permissions:`, el `if:`, los `needs:`— también: en los consumers el sello protege a los guardianes del merge (`governed-merge.yml`, `force-push-guard.yml`, `push-compliance.yml`), y la cobertura de los pines la da `workflow-policy` (HR-26), que solo mira el formato del pin. El formato del manifiesto pasa a `schema_version: 2`: un manifiesto v1 no lo mide ningún gate (sale 3 con el comando de re-sellar en el mensaje) y `ci-pattern manifest seal` lo migra, de modo que el cambio de formato es explícito y no silencioso. Los tres consumers re-sellan en la misma tanda de propagación.

Lote de correccion (verificacion independiente de #1200): la normalizacion solo
toca la clave YAML `uses:` y solo su sha con un comentario final de version
(` # v1.2.3`, con espacio delante del `#` y nada detras). Lo que va dentro de un
bloque escalar (`run: |`) es shell, no YAML, y no se normaliza: tragarse esas
lineas dejaba carga util fuera del sello y hacia que un cambio de logica no
moviera el hash. Tampoco se normaliza un `uses:` con otra forma de pin (una
etiqueta, un `uses:` local, o un comentario que no es una version como
`@<sha>#algo` sin espacio): entonces el sha sigue dentro del sello. Y al migrar
el formato del manifiesto, los hashes del formato viejo se comprueban antes de
migrar, para que la deriva de un fichero declarado no se esconda bajo el cambio
de algoritmo; quien cambia un fichero lo declara con `--reeseal`. El
propagador (`scripts/propagate-team-skills.ps1`) deja de calcular el hash por su
cuenta y re-sella con `ci-pattern manifest seal`, que es el canon del hash.
