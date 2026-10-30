type: patch

El preflight de tooling separa «binario ausente» (aviso) de «contrato roto» (bloqueo) (#1172, decisión del dueño del 2026-10-24; auditoría de fricción de gates en team-skills#1093, comentario `6389037313`).

La auditoría clasificó 19 job-runs rojos del gate `preflight-tooling` en la ventana de 48 h: 8 de fricción y 11 de infraestructura. Los 11 de infraestructura eran el mismo caso —un preflight que exigía binarios que el runner no tiene— y el resultado era un PR en rojo que quien lo abre no puede arreglar: el fix aterriza en la definición del servicio del runner, no en el código.

Ahora el preflight distingue los dos fallos que antes compartían camino:

- **Binario ausente en el runner → AVISO.** El preflight lo dice por stderr, una línea por ausente, con el mensaje exacto `falta <cli> en el runner <nombre>`, **cómo se instala** y sale **0**. En GitHub Actions el aviso sale además anotado como `::warning::`, así que se ve en el PR sin ponerlo en rojo. La lista puede declarar la instrucción de instalación por entrada, con el comentario `# instalar: <comando>` (o `# install:`); sin ella, el aviso remite a la definición del servicio del runner.
- **Contrato roto → BLOQUEO.** Falta `--binaries`, la lista no existe, no es legible o está vacía, `--runner` sin nombre u opción no reconocida: sale **2**, como antes. Un contrato sin declarar no es un aviso: es un hueco.

El código de salida 1 (binario ausente) desaparece: la invocación heredada `--binaries FICHERO` a secas conserva su texto y ahora avisa con exit 0. La forma de la lista no cambia (una CLI por línea, `#` abre comentario, la línea en blanco se ignora); lo nuevo es el comentario opcional de instalación.

Tests: la suite del preflight pasa de 11 a 15 casos —el binario ausente avisa y sale 0 con el mensaje exacto, dice cómo instalarlo, respeta la instrucción declarada por la entrada, nombra todos los ausentes, la invocación heredada también avisa, el aviso sale anotado como warning de Actions, y los cuatro casos de contrato roto siguen saliendo con exit 2—. `shellcheck` sale limpio en el script y en la suite.
