type: patch

El aviso del rojo nocturno dice qué fallo hubo: `scripts/report-nightly-failure.sh` saca el resumen del cuerpo de los jobs del run (el job y el paso que salieron en rojo) cuando el llamante no trae uno, y cuando no puede consultarlos nombra el permiso que falta (`actions: read` en el token del job) en vez de dejar un «sin resumen» mudo (#1332).
