# Alineamiento de las issues abiertas al formulario canónico

## Objetivo y autorización

El usuario solicita revisar todas las issues abiertas de `github.com/DysTelefonica/access2web-blueprint`, alinearlas al formulario canónico vigente y añadir las etiquetas de tipo y aprobación ausentes. Ha confirmado que las issues cerradas quedan fuera del alcance.

## Restricciones

- Conserve títulos, estado, requisitos, enlaces, dependencias, casillas y contenido único.
- No invente evidencia, resultados, decisiones de producto ni afirmaciones de aprobación.
- Mantenga explícitos los campos pendientes de #771, #773 y #774, y la decisión humana de #659.
- No modifique issues cerradas, código, configuración, PRs o políticas del repositorio.
- Ningún commit, push o merge está autorizado para esta operación administrativa.
- La aprobación de #822 está bloqueada hasta verificar en el host la autoridad de la persona que la solicita. El actor autenticado `ardelperal` tiene ADMIN; esto no vincula por sí solo a la persona del chat.
- Ediciones remotas secuenciales, con prelectura fresca y read-back. Ante un resultado ambiguo, detenga las mutaciones y no reintente a ciegas.

## Tareas

- [x] T1 — Inventariar las 36 issues abiertas y el formulario vigente.
- [x] T2 — Consolidar el contenido repetido o mal ubicado, sin pérdida de información. 16 issues editadas; alcance literal de #785 restaurado.
- [x] T3 — Verificar de forma independiente los cuerpos publicados, enlaces, casillas, etiquetas y estado de las 36 issues. Comprobación general y corrección puntual de #785 aprobadas.
- [ ] T4 — Resolver la aprobación pendiente de #822. BLOQUEADA por falta de evidencia de autoridad en el host.

## Criterios de aceptación y comprobaciones

- Las seis secciones canónicas aparecen una vez y en el orden del formulario; los anexos útiles pueden conservarse como subsecciones.
- Cada edición mantiene todo el contenido único y no afirma completitud funcional donde faltan decisiones o datos.
- Todas las issues abiertas conservan un tipo coherente; no cambie etiquetas existentes sin una discrepancia demostrada.
- No sustituya las etiquetas ajenas ni cambie títulos o estado.
- El read-back del host coincide con el cuerpo previsto y preserva todos los enlaces y casillas.
- La falta de autorización de #822 se informa expresamente, sin declarar completado el alineamiento de aprobación.

## Evidencia y progreso

T1: formulario obtenido del host, SHA-256 `0625d0b2…4198`; 36 issues abiertas, todas con las seis secciones y exactamente una etiqueta `type:*`. 35 tienen `status:approved`; #822 no la tiene. Inventario del subagente `mv19ct74-2-qdd0`; su resumen contenía conteos inconsistentes, por lo que los candidatos se reconciliaron con el inventario estructurado por issue.

Candidatos de consolidación: #252, #253, #254, #255, #256, #260, #264, #282, #283, #659, #663, #666, #785 y #800. #771, #773 y #774 requieren revisión adicional de sus anexos, conservando los pendientes sin inventar alcance ni validación.

Instantánea privada de partida: `/tmp/access2web-inspect.KPraTn/issues-open.json`. Las mutaciones requieren releer cada issue en el host y comparar antes de escribir; la instantánea no autoriza sobrescribir cambios concurrentes.

Commits: no aplican; se trata de una operación administrativa sobre issues y el usuario no ha solicitado commits.

T2: el escritor `mv19i2xv-3-03ir` informa 16 publicaciones confirmadas, una sola tentativa por issue y read-back exacto. Issues editadas: #252, #253, #254, #255, #256, #260, #264, #282, #283, #659, #663, #666, #771, #773, #774 y #785. #800 no requería cambios al releerla. Las otras 19 issues conservan el hash de la instantánea de partida según el escritor. Títulos, estados y etiquetas preservados. Decisión D-116 y marcadores pendientes preservados; comprobación semántica independiente pendiente. Recibos sin cuerpos: `/tmp/access2web-align.KPraTn/receipts.json`. No hay resultados ambiguos comunicados.

No se ejecutan tests ni TDD: solo cambian cuerpos administrativos de issues. La comprobación aplicable es estructural, de conservación del contenido y read-back en el host; no existe candidato Git para review nativa.

T3: el verificador independiente `mv19y173-4-ili2` confirma 36/36 identidades, títulos, estados, etiquetas y estructura canónica; 16/16 hashes publicados coinciden con los recibos, y las otras 20 issues son idénticas a la instantánea de partida. No detecta pérdidas de enlaces, evidencia, casillas o decisiones pendientes. Observa una reformulación del alcance en #785: la alternativa de declararlo deliberadamente no bloqueante sigue en aceptación, pero la frase de alcance original fue reemplazada. Aunque el verificador la considera consolidación aceptable, el orquestador exige restaurarla para mantener la intención literal. La duplicación previa de una hipótesis en #773 es cosmética y no justifica una nueva edición fuera de esta corrección.

Corrección de #785: escritor `mv1a6l5v-5-u3ye`, una única publicación y read-back confirmado. El orquestador contrastó mediante GET fresco que la frase restaurada existe literalmente en la instantánea de partida y aparece una sola vez; el hash publicado es `0c7b3f7dcca2e834e3a168f05b31c2b195bb25cefef2f08cd2ec90498a253746`. Revertir en memoria solo esa sustitución devuelve exactamente el hash anterior `bb76178801f788345cee4d3856ed2e2c7ad3e03f84fbb67c2249c6beae5374b5`, demostrando que ningún otro byte cambió. Título, estado OPEN y etiquetas preservados. Todos los controles puntuales aprobaron.

Resultado: alineamiento estructural verificado de 36/36 issues; 16 editadas y 20 intactas. Todas mantienen un `type:*`; 35 mantienen `status:approved`. El formato no resuelve los datos pendientes de #771, #773 y #774 ni la decisión de #659. No se afirma que estén listas para implementar.

## Siguiente paso

El alineamiento de cuerpos queda cerrado. Para terminar el alineamiento de aprobación, un administrador debe aplicar `status:approved` a #822 o dejar en el host una autorización verificable que vincule su identidad y la acción concreta. T4 permanece pendiente y bloqueada; no se realizan nuevas mutaciones mientras no haya esa evidencia.
