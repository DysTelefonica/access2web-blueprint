type: patch

**Qué.** El resolvedor que da la invariante HR de un gate distribuido (`_hr_citing_gate`, en `assets/bin/ci-pattern`) lee ahora el bloque entero de la entrada HR de `SKILL.md`, con sus sub-reglas, y no solo su primera línea.
**Causa.** Las entradas HR son largas y envuelven el «gate: `assets/...`» en la línea siguiente, y las sub-reglas (HR-54.5, HR-54.11, HR-54.15) van indentadas: el resolvedor no las veía. Por eso el aviso de gates sin cablear nombraba «sin HR» 27 de los 42 gates del canon, cuando 15 de ellos sí tienen su HR escrita.
**Arreglo.** Se lee el bloque de la entrada (con sus sub-reglas) y se busca la cita del fichero y, si no hay, la de su carpeta. Lo que ya resolvía la primera línea no cambia, y un gate que ninguna entrada HR cita sigue sin HR: el aviso no inventa.
**A partir de ahora.** El aviso nombra la HR del gate citado en la continuación; los gates que siguen sin HR son los que ninguna entrada HR cita, que son los que cierra el paso 4 de #1755.
