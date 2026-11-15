type: patch

El sentido adelante de `docs-codeguide` deja de tratar como ruta del repo las referencias a otro repo (`owner/repo#N`), las acciones versionadas (`owner/repo@ref`) y los nombres de acción sin versión (`actions/checkout`) cuyo primer segmento no existe en la raíz del repo. Antes, citar `actions/checkout` o `DysTelefonica/team-skills#1270` en el codeguide daba un falso «missing». Un directorio muerto con barra final, o bajo un árbol que sí existe, sigue fallando (#1359, reincidencia de #1004).
